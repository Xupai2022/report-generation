from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path
from typing import List, Optional

from mss_ai_ppt_sample_assets.backend import config


class PreviewGenerationError(Exception):
    pass


INVALID_FS_CHARS = [":", "*", "?", "\"", "<", ">", "|"]


def sanitize_job_id(job_id: str) -> str:
    sanitized = job_id
    for ch in INVALID_FS_CHARS:
        sanitized = sanitized.replace(ch, "_")
    sanitized = sanitized.replace("\\", "_").replace("/", "_")
    return sanitized


class PPTPreviewGenerator:
    """Convert PPTX to slide images using real rendering engines.

    Priority:
      1. LibreOffice / OpenOffice (via `soffice` CLI, if installed)
    """

    def __init__(self, base_dir: Path = config.PREVIEWS_DIR):
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def _find_soffice(self) -> str:
        """Locate the soffice executable."""
        soffice_candidates = []
        env_path = os.getenv("LIBREOFFICE_PATH")
        if env_path:
            soffice_candidates.append(Path(env_path))

        # Common default installation paths on Windows
        soffice_candidates.extend(
            [
                Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
                Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
                Path(r"C:\Program Files\OpenOffice 4\program\soffice.exe"),
            ]
        )

        for cand in soffice_candidates:
            if cand.is_file():
                return str(cand)

        return shutil.which("soffice") or "soffice"

    def _with_libreoffice_pdf(self, ppt_path: Path, output_dir: Path) -> Path:
        """
        Use LibreOffice/OpenOffice in headless mode to render PPTX to a PDF file.

        Requires local installation of LibreOffice with `soffice` CLI available.
        """
        soffice = self._find_soffice()

        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output_dir),
                    str(ppt_path),
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
        except Exception as exc:
            raise PreviewGenerationError(
                "LibreOffice/OpenOffice (`soffice`) is required for PPT preview "
                "if PowerPoint is not available. Please install LibreOffice and, "
                "if needed, set environment variable LIBREOFFICE_PATH to the "
                "full path of soffice.exe."
            ) from exc

        pdf_files = sorted(output_dir.glob("*.pdf"))
        if not pdf_files:
            raise PreviewGenerationError("No PDF generated from LibreOffice export")
        # Usually only one PDF is generated
        return pdf_files[0]

    def _with_libreoffice_png(self, ppt_path: Path, output_dir: Path) -> List[Path]:
        """
        Use LibreOffice/OpenOffice in headless mode to render PPTX directly to PNG slides.

        This avoids dependence on Poppler/pdf2image and is more robust on Windows.
        """
        soffice = self._find_soffice()

        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            # Use Impress PNG filter + PageRange to export each slide as one PNG.
            # The upper bound (e.g. 1-50) is generous; extra pages are ignored.
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--convert-to",
                    "png:impress_png_Export:PageRange=1-50",
                    "--outdir",
                    str(output_dir),
                    str(ppt_path),
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
        except Exception as exc:
            raise PreviewGenerationError(
                "LibreOffice/OpenOffice (`soffice`) is required for PPT preview "
                "if PowerPoint is not available. Please install LibreOffice and, "
                "if needed, set environment variable LIBREOFFICE_PATH to the "
                "full path of soffice.exe."
            ) from exc

        images = sorted(output_dir.glob("*.png"))
        if not images:
            raise PreviewGenerationError("No PNG images generated from LibreOffice export")
        return images

    def _pdf_to_images(self, pdf_path: Path, output_dir: Path) -> List[Path]:
        """
        Convert a multi-page PDF into slide1.png, slide2.png, ... trying PyMuPDF (fitz),
        Poppler, then pdf2image, and finally LibreOffice as a last resort.

        Requires poppler binaries installed on the system. On Windows you may set
        POPPLER_PATH to the directory containing the poppler executables.
        """
        output_dir.mkdir(parents=True, exist_ok=True)

        # 1. Try PyMuPDF (fitz) - Robust, no external binaries required
        try:
            import fitz  # PyMuPDF

            doc = fitz.open(str(pdf_path))
            result: List[Path] = []
            
            # Use a zoom factor of 2.0 to get approx 144 DPI (72 * 2)
            mat = fitz.Matrix(2.0, 2.0)
            
            for i, page in enumerate(doc):
                pix = page.get_pixmap(matrix=mat)
                target = output_dir / f"slide{i+1}.png"
                pix.save(str(target))
                result.append(target)
            
            doc.close()
            
            if result:
                return sorted(result)

        except ImportError:
            # pymupdf not installed, fall back to others
            pass
        except Exception as exc:
            # If PyMuPDF fails, log/ignore and fall back
            # In a real app, logging 'exc' would be good.
            pass

        # Prefer Poppler's pdftoppm if available
        poppler_dir = os.getenv("POPPLER_PATH")
        pdftoppm_path: Optional[str] = None
        if poppler_dir:
            candidate = Path(poppler_dir) / ("pdftoppm.exe" if os.name == "nt" else "pdftoppm")
            if candidate.is_file():
                pdftoppm_path = str(candidate)
        if pdftoppm_path is None:
            which_pdftoppm = shutil.which("pdftoppm")
            if which_pdftoppm:
                pdftoppm_path = which_pdftoppm

        if pdftoppm_path:
            output_dir.mkdir(parents=True, exist_ok=True)
            try:
                subprocess.run(
                    [
                        pdftoppm_path,
                        "-png",
                        "-r",
                        "150",
                        str(pdf_path),
                        "slide",
                    ],
                    check=True,
                    cwd=str(output_dir),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
            except Exception as exc:
                raise PreviewGenerationError(
                    f"Poppler pdftoppm failed. PATH/POPPLER_PATH used. Underlying error: {exc}"
                ) from exc

            result: List[Path] = sorted(output_dir.glob("slide-*.png"))
            if not result:
                raise PreviewGenerationError(f"Poppler pdftoppm produced no images in {output_dir}")
            return result

        # Next try python-based pdf2image (if installed)
        try:
            from pdf2image import convert_from_path  # type: ignore
        except ImportError:
            convert_from_path = None

        if convert_from_path:
            try:
                images = convert_from_path(
                    str(pdf_path),
                    dpi=150,
                    fmt="png",
                    output_folder=None,
                    poppler_path=poppler_dir,
                )
            except Exception as exc:
                raise PreviewGenerationError(
                    f"Poppler is required for PDF preview via pdf2image. "
                    f"POPPLER_PATH={poppler_dir!r}. Underlying error: {exc}"
                ) from exc

            output_dir.mkdir(parents=True, exist_ok=True)
            result: List[Path] = []
            for idx, img in enumerate(images, start=1):
                target = output_dir / f"slide{idx}.png"
                img.save(target, "PNG")
                result.append(target)

            if not result:
                raise PreviewGenerationError("No images generated from PDF export")
            return result

        # Last resort: use LibreOffice to convert the PDF into PNG pages
        soffice = self._find_soffice()
        output_dir.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--convert-to",
                    "png:draw_png_Export:PageRange=1-50",
                    "--outdir",
                    str(output_dir),
                    str(pdf_path),
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
        except Exception as exc:
            raise PreviewGenerationError(
                "Failed to convert PDF to images via LibreOffice. Please install Poppler (pdftoppm) "
                "or python package pdf2image with Poppler binaries available."
            ) from exc

        result: List[Path] = sorted(output_dir.glob("*.png"))
        if not result:
            raise PreviewGenerationError("LibreOffice PDF->PNG produced no images")
        return result

    def _via_pdf_pipeline(self, ppt_path: Path, output_dir: Path) -> List[Path]:
        """
        Render PPT -> PDF -> PNG to ensure we get one image per slide.

        LibreOffice sometimes only emits the first slide when exporting
        directly to PNG, so this path forces a multi-page PDF first.
        """
        pdf_path = self._with_libreoffice_pdf(ppt_path, output_dir)
        return self._pdf_to_images(pdf_path, output_dir)

    def to_images(
        self, ppt_path: Path, job_id: str, expected_count: Optional[int] = None
    ) -> List[Path]:
        if not ppt_path.exists():
            raise PreviewGenerationError(f"PPT file not found: {ppt_path}")

        job_dir = sanitize_job_id(job_id)
        output_dir = self.base_dir / job_dir
        if output_dir.exists():
            shutil.rmtree(output_dir)

        # Only use LibreOffice / OpenOffice (PPTX -> PNG directly, no Poppler, no PowerPoint)
        try:
            images = self._with_libreoffice_png(ppt_path, output_dir)

            needs_pdf_fallback = False
            if expected_count is not None:
                # We know how many slides should exist; if LibreOffice only gave us
                # one (common on Windows), force PDF pipeline to get per-slide images.
                needs_pdf_fallback = expected_count > 1 and len(images) < expected_count
            else:
                # When count is unknown, a single image is suspicious for PPT decks.
                needs_pdf_fallback = len(images) <= 1

            if needs_pdf_fallback:
                shutil.rmtree(output_dir, ignore_errors=True)
                images = self._via_pdf_pipeline(ppt_path, output_dir)

            return images
        except Exception as second_error:
            raise PreviewGenerationError(
                f"Failed to export preview via LibreOffice/OpenOffice only. Underlying error: {second_error}"
            ) from second_error

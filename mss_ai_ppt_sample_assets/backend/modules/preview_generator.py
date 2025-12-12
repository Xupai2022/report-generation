from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path
from typing import List

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
    """Convert PPTX to slide images using LibreOffice and PyMuPDF.

    Pipeline: PPTX → LibreOffice → PDF → PyMuPDF → PNG images
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

    def _pptx_to_pdf(self, ppt_path: Path, output_dir: Path) -> Path:
        """
        Use LibreOffice in headless mode to convert PPTX to PDF.

        Requires LibreOffice with `soffice` CLI available.
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
                "LibreOffice (`soffice`) is required for PPTX to PDF conversion. "
                "Please install LibreOffice and, if needed, set environment variable "
                "LIBREOFFICE_PATH to the full path of soffice.exe."
            ) from exc

        pdf_files = sorted(output_dir.glob("*.pdf"))
        if not pdf_files:
            raise PreviewGenerationError("No PDF generated from LibreOffice export")
        return pdf_files[0]

    def _pdf_to_images(self, pdf_path: Path, output_dir: Path) -> List[Path]:
        """
        Convert PDF to PNG images using PyMuPDF (fitz).

        Uses a zoom factor of 2.0 for 144 DPI output quality.
        """
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            import fitz  # PyMuPDF
        except ImportError as exc:
            raise PreviewGenerationError(
                "PyMuPDF (pymupdf) is required for PDF to PNG conversion. "
                "Install it with: pip install pymupdf"
            ) from exc

        try:
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

            if not result:
                raise PreviewGenerationError("No images generated from PDF")

            return sorted(result)

        except Exception as exc:
            raise PreviewGenerationError(
                f"Failed to convert PDF to images using PyMuPDF: {exc}"
            ) from exc

    def to_images(self, ppt_path: Path, job_id: str) -> List[Path]:
        """
        Convert PPTX to PNG images.

        Pipeline: PPTX → LibreOffice → PDF → PyMuPDF → PNG
        """
        if not ppt_path.exists():
            raise PreviewGenerationError(f"PPT file not found: {ppt_path}")

        job_dir = sanitize_job_id(job_id)
        output_dir = self.base_dir / job_dir
        if output_dir.exists():
            shutil.rmtree(output_dir)

        # Step 1: PPTX → PDF (LibreOffice)
        pdf_path = self._pptx_to_pdf(ppt_path, output_dir)

        # Step 2: PDF → PNG (PyMuPDF)
        images = self._pdf_to_images(pdf_path, output_dir)

        return images

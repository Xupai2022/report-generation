from __future__ import annotations

import os
import shutil
import subprocess
import time
import logging
from pathlib import Path
from typing import Dict, List, Tuple

from mss_ai_ppt_sample_assets.backend import config


class PreviewGenerationError(Exception):
    pass


INVALID_FS_CHARS = [":", "*", "?", "\"", "<", ">", "|"]

logger = logging.getLogger(__name__)


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

    def __init__(self, base_dir: Path = config.PREVIEWS_DIR, cleanup_days: int = 7):
        """
        Initialize the preview generator.

        Args:
            base_dir: Base directory for storing previews
            cleanup_days: Number of days to keep previews (default: 7)
                         Set to 0 to disable automatic cleanup
        """
        self.base_dir = base_dir
        self.cleanup_days = cleanup_days
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

    def _cleanup_old_previews(self) -> None:
        """
        Clean up preview directories older than cleanup_days.

        This method is automatically called when generating new previews.
        It removes directories whose last modification time exceeds the threshold.
        """
        if self.cleanup_days <= 0:
            return  # Cleanup disabled

        cutoff_time = time.time() - (self.cleanup_days * 24 * 60 * 60)

        try:
            for item in self.base_dir.iterdir():
                if not item.is_dir():
                    continue
                # Preserve static template preview assets used by pre-config UI.
                if item.name.startswith("template_"):
                    continue

                # Check directory modification time (last access/creation)
                dir_mtime = item.stat().st_mtime

                if dir_mtime < cutoff_time:
                    try:
                        shutil.rmtree(item)
                        # Optional: log cleanup (you can enable logging if needed)
                        # import logging
                        # logging.info(f"Cleaned up old preview: {item.name}")
                    except Exception:
                        # Ignore errors for individual directories (might be in use)
                        pass
        except Exception:
            # Ignore cleanup errors to not block preview generation
            pass

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

    def _pptx_to_pdf_with_timings(
        self, ppt_path: Path, output_dir: Path
    ) -> Tuple[Path, Dict[str, float]]:
        start = time.perf_counter()
        pdf_path = self._pptx_to_pdf(ppt_path, output_dir)
        duration_ms = (time.perf_counter() - start) * 1000
        logger.info(f"PPTX->PDF in {duration_ms:.0f}ms ({ppt_path.name})")
        return pdf_path, {"pptx_to_pdf_ms": duration_ms}

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
            
            mat = fitz.Matrix(1.2, 1.2)

            for i, page in enumerate(doc):
                pix = page.get_pixmap(matrix=mat)
                target = output_dir / f"slide{i+1}.png"
                pix.save(str(target))
                result.append(target)

            doc.close()

            if not result:
                raise PreviewGenerationError("No images generated from PDF")

            # Sort by numeric slide number to ensure correct order
            # (slide1.png, slide2.png, ..., slide10.png instead of dictionary order)
            return sorted(result, key=lambda p: int(p.stem.replace('slide', '')))

        except Exception as exc:
            raise PreviewGenerationError(
                f"Failed to convert PDF to images using PyMuPDF: {exc}"
            ) from exc

    def _pdf_to_images_with_timings(
        self, pdf_path: Path, output_dir: Path
    ) -> Tuple[List[Path], Dict[str, float]]:
        start = time.perf_counter()
        images = self._pdf_to_images(pdf_path, output_dir)
        duration_ms = (time.perf_counter() - start) * 1000
        logger.info(f"PDF->images in {duration_ms:.0f}ms ({len(images)} pages)")
        return images, {"pdf_to_images_ms": duration_ms}

    def to_images(self, ppt_path: Path, job_id: str) -> List[Path]:
        """
        Convert PPTX to PNG images.

        Pipeline: PPTX → LibreOffice → PDF → PyMuPDF → PNG

        This method also triggers automatic cleanup of old previews
        if cleanup_days > 0.
        """
        # Cleanup old previews before generating new ones
        self._cleanup_old_previews()

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

    def to_images_with_timings(
        self, ppt_path: Path, job_id: str
    ) -> Tuple[List[Path], Dict[str, float]]:
        """Convert PPTX to PNG images and return timing breakdown.

        Returns:
            (images, timings_ms)
        """
        # Cleanup old previews before generating new ones
        self._cleanup_old_previews()

        if not ppt_path.exists():
            raise PreviewGenerationError(f"PPT file not found: {ppt_path}")

        job_dir = sanitize_job_id(job_id)
        output_dir = self.base_dir / job_dir
        if output_dir.exists():
            shutil.rmtree(output_dir)

        total_start = time.perf_counter()
        pdf_path, t1 = self._pptx_to_pdf_with_timings(ppt_path, output_dir)
        images, t2 = self._pdf_to_images_with_timings(pdf_path, output_dir)
        total_ms = (time.perf_counter() - total_start) * 1000

        timings: Dict[str, float] = {}
        timings.update(t1)
        timings.update(t2)
        timings["pptx_to_images_total_ms"] = total_ms
        return images, timings

    def get_pdf_path(self, ppt_path: Path, job_id: str) -> Path:
        """
        Get or generate PDF file for a PPT file.

        If PDF already exists in preview directory, return it.
        Otherwise, generate it using LibreOffice.

        Returns:
            Path to the PDF file
        """
        if not ppt_path.exists():
            raise PreviewGenerationError(f"PPT file not found: {ppt_path}")

        job_dir = sanitize_job_id(job_id)
        output_dir = self.base_dir / job_dir
        output_dir.mkdir(parents=True, exist_ok=True)

        # Check if PDF already exists
        pdf_files = sorted(output_dir.glob("*.pdf"))
        if pdf_files:
            return pdf_files[0]

        # Generate PDF if not exists
        pdf_path, _timings = self._pptx_to_pdf_with_timings(ppt_path, output_dir)
        return pdf_path

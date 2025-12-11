from __future__ import annotations

import os
import shutil
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
    """Convert PPTX to slide images. Prefers PowerPoint COM on Windows."""

    def __init__(self, base_dir: Path = config.PREVIEWS_DIR):
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def _with_powerpoint(self, ppt_path: Path, output_dir: Path) -> List[Path]:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise PreviewGenerationError("win32com.client and pythoncom are required for PowerPoint-based export") from exc

        # Ensure COM is initialized for this thread
        pythoncom.CoInitialize()
        powerpoint = None
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # Visible must be True on some systems; hiding triggers an error
            powerpoint.Visible = True
            presentation = powerpoint.Presentations.Open(str(ppt_path), WithWindow=True, ReadOnly=True)
            output_dir.mkdir(parents=True, exist_ok=True)
            width = int(presentation.PageSetup.SlideWidth)
            height = int(presentation.PageSetup.SlideHeight)

            for slide in presentation.Slides:
                idx = int(slide.SlideIndex)
                target = output_dir / f"slide{idx}.png"
                slide.Export(str(target), "PNG", width, height)

            presentation.Close()
        finally:
            if powerpoint:
                powerpoint.Quit()
            pythoncom.CoUninitialize()

        images = sorted(output_dir.glob("slide*.png"))
        if not images:
            raise PreviewGenerationError("No images generated from PowerPoint export")
        return images

    def to_images(self, ppt_path: Path, job_id: str) -> List[Path]:
        if not ppt_path.exists():
            raise PreviewGenerationError(f"PPT file not found: {ppt_path}")

        job_dir = sanitize_job_id(job_id)
        output_dir = self.base_dir / job_dir
        if output_dir.exists():
            shutil.rmtree(output_dir)

        # Try PowerPoint COM (Windows). If not available, raise with guidance.
        try:
            return self._with_powerpoint(ppt_path, output_dir)
        except PreviewGenerationError:
            raise
        except Exception as exc:
            raise PreviewGenerationError(f"Failed to export preview: {exc}") from exc

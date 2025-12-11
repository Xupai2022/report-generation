from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository


def _get_nested(data: Dict[str, Any], path: str) -> Any:
    current = data
    for part in path.split("."):
        if isinstance(current, dict):
            current = current.get(part)
        else:
            current = None
            break
    return current


class PPTGenerator:
    """Fill PPTX template placeholders with slidespec content."""

    def __init__(self, template_repo: TemplateRepository):
        self.template_repo = template_repo
        try:
            from pptx import Presentation  # type: ignore

            self._Presentation = Presentation
        except ImportError as exc:  # pragma: no cover - runtime dependency
            raise RuntimeError(
                "python-pptx is required for PPT rendering. Please install dependencies via requirements.txt."
            ) from exc

    def _replace_tokens_in_shape(self, shape, mapping: Dict[str, str]) -> None:
        if not shape.has_text_frame:
            return
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text = run.text
                for token, value in mapping.items():
                    placeholder = f"{{{{{token}}}}}"
                    if placeholder in text:
                        run.text = text.replace(placeholder, value)
                        text = run.text

    def render(self, slidespec: SlideSpec, output_path: Path) -> Path:
        template_path = self.template_repo.get_pptx_path(slidespec.template_id)
        prs = self._Presentation(template_path)
        descriptor = self.template_repo.get_descriptor(slidespec.template_id)
        slides_by_key = {s.slide_key: s for s in slidespec.slides}

        for slide_desc in descriptor.slides:
            spec_slide = slides_by_key.get(slide_desc.slide_key)
            if not spec_slide:
                continue
            mapping: Dict[str, str] = {}
            for token, data_path in slide_desc.render_map.items():
                value = _get_nested(spec_slide.data, data_path)
                if value is None:
                    value = ""
                if isinstance(value, list):
                    value = "\n".join(str(v) for v in value)
                mapping[token] = str(value)

            # pptx slides are 0-based
            slide = prs.slides[slide_desc.slide_no - 1]
            for shape in slide.shapes:
                self._replace_tokens_in_shape(shape, mapping)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                # If the file is locked (e.g., opened by PowerPoint), fall back to a temp name
                output_path = output_path.with_name(output_path.stem + "_new" + output_path.suffix)
        prs.save(output_path)
        return output_path

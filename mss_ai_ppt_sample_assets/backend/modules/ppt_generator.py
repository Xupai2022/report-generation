from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Union

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec, SlideSpecV2
from mss_ai_ppt_sample_assets.backend.models.templates import TemplateDescriptorV2
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository


def _get_nested(data: Dict[str, Any], path: str) -> Any:
    """Get nested value from dict using dot notation."""
    current = data
    for part in path.split("."):
        if isinstance(current, dict):
            current = current.get(part)
        else:
            current = None
            break
    return current


class PPTGeneratorV2:
    """Fill PPTX template placeholders with V2 slidespec content.

    V2 simplification: slidespec.placeholders directly maps token -> value,
    no need to traverse render_map paths.
    """

    def __init__(self, template_repo: TemplateRepository):
        self.template_repo = template_repo
        try:
            from pptx import Presentation
            self._Presentation = Presentation
        except ImportError as exc:
            raise RuntimeError(
                "python-pptx is required for PPT rendering. Please install via requirements.txt."
            ) from exc

    def _replace_tokens_in_shape(self, shape, mapping: Dict[str, str]) -> None:
        """Replace {{TOKEN}} placeholders in shape text."""
        if not shape.has_text_frame:
            return
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text = run.text
                for token, value in mapping.items():
                    placeholder = f"{{{{{token}}}}}"
                    if placeholder in text:
                        run.text = text.replace(placeholder, str(value) if value else "")
                        text = run.text

    def render(self, slidespec: SlideSpecV2, output_path: Path) -> Path:
        """Render V2 slidespec to PPTX.

        Args:
            slidespec: V2 slidespec with placeholders filled
            output_path: Where to save the output PPTX

        Returns:
            Path to the saved PPTX file
        """
        template_path = self.template_repo.get_pptx_path(slidespec.template_id)
        prs = self._Presentation(template_path)

        # Build mapping: slide_no -> placeholders dict
        slides_by_no = {s.slide_no: s for s in slidespec.slides}

        for slide_no, slide_content in slides_by_no.items():
            if slide_no > len(prs.slides):
                continue

            # pptx slides are 0-indexed
            pptx_slide = prs.slides[slide_no - 1]

            # V2: placeholders is directly {token: value}
            mapping: Dict[str, str] = {}
            for token, value in slide_content.placeholders.items():
                if value is None:
                    value = ""
                elif isinstance(value, list):
                    # Join list items with newlines
                    value = "\n".join(str(v) for v in value)
                mapping[token] = str(value)

            # Replace tokens in all shapes
            for shape in pptx_slide.shapes:
                self._replace_tokens_in_shape(shape, mapping)

        # Save output
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                output_path = output_path.with_name(output_path.stem + "_new" + output_path.suffix)

        prs.save(output_path)
        return output_path


class PPTGenerator:
    """Fill PPTX template placeholders with V1 slidespec content (legacy)."""

    def __init__(self, template_repo: TemplateRepository):
        self.template_repo = template_repo
        try:
            from pptx import Presentation
            self._Presentation = Presentation
        except ImportError as exc:
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
        """Render V1 slidespec to PPTX (legacy)."""
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
                output_path = output_path.with_name(output_path.stem + "_new" + output_path.suffix)
        prs.save(output_path)
        return output_path


def create_ppt_generator(template_repo: TemplateRepository, template_id: str) -> Union[PPTGenerator, PPTGeneratorV2]:
    """Factory function to create appropriate PPT generator based on template version.

    Args:
        template_repo: Template repository
        template_id: Template ID

    Returns:
        PPTGeneratorV2 for V2 templates, PPTGenerator for V1
    """
    if template_repo.is_v2(template_id):
        return PPTGeneratorV2(template_repo)
    return PPTGenerator(template_repo)

from __future__ import annotations

from pathlib import Path
from typing import Dict

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpecV2
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository


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
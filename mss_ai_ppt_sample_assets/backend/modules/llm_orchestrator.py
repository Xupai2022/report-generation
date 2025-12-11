from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Optional

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec, SlideSpecItem
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository
from mss_ai_ppt_sample_assets.backend.modules.data_prep import DataPrepResult


class MockOutputNotFound(Exception):
    pass


class LLMOrchestrator:
    """Orchestrates content generation. Defaults to deterministic/mock to avoid live LLM dependency."""

    def __init__(self, template_repo: Optional[TemplateRepository] = None):
        self.template_repo = template_repo or TemplateRepository()

    def _load_mock_slidespec(self, input_id: str, template_id: str) -> SlideSpec:
        audience = "management" if "management" in template_id else "technical"
        mock_file = f"{input_id}_{audience}_mock_slidespec.json"
        path = config.MOCK_OUTPUTS_DIR / mock_file
        if not path.exists():
            raise MockOutputNotFound(f"Mock slidespec {mock_file} not found")
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return SlideSpec.parse_obj(data)

    def generate_slidespec(
        self,
        input_id: str,
        template_id: str,
        prepared: DataPrepResult,
        use_mock: bool = True,
    ) -> SlideSpec:
        """Return slidespec either from mock, deterministic mapping, or (future) LLM."""
        if use_mock or not config.settings.enable_llm:
            try:
                return self._load_mock_slidespec(input_id, template_id)
            except MockOutputNotFound:
                # Fallback to deterministic build
                pass

        # Deterministic content: map prepared slide inputs directly into slidespec
        slides = []
        template = self.template_repo.get_descriptor(template_id)
        for slide in template.slides:
            data = prepared.slide_inputs.get(slide.slide_key, {})
            slides.append(
                SlideSpecItem(
                    slide_no=slide.slide_no,
                    slide_key=slide.slide_key,
                    data=data,
                )
            )
        return SlideSpec(template_id=template_id, slides=slides)

    def rewrite_slide(
        self,
        slide_spec: SlideSpec,
        slide_key: str,
        new_content: Dict[str, Any],
    ) -> SlideSpec:
        """Update a slide's data with new content (placeholder for real LLM rewrite)."""
        updated_slides = []
        for slide in slide_spec.slides:
            if slide.slide_key == slide_key:
                updated_slides.append(
                    SlideSpecItem(
                        slide_no=slide.slide_no,
                        slide_key=slide.slide_key,
                        data={**slide.data, **new_content},
                    )
                )
            else:
                updated_slides.append(slide)
        return SlideSpec(template_id=slide_spec.template_id, slides=updated_slides)

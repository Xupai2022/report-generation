from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any, Dict, Union

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpecV2
from mss_ai_ppt_sample_assets.backend.modules import (
    AuditLogger,
    TemplateRepository,
)
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.modules.ppt_generator import PPTGeneratorV2
from mss_ai_ppt_sample_assets.backend.modules.validator import ValidatorV2
from mss_ai_ppt_sample_assets.backend.modules.preview_generator import (
    PPTPreviewGenerator,
    sanitize_job_id,
)


class InputNotFoundError(Exception):
    pass


class SlideSpecNotFoundError(Exception):
    pass


class ReportService:
    """Orchestrates report generation for V2 (AI-driven) templates."""

    def __init__(self):
        self.template_repo = TemplateRepository()
        self.audit_logger = AuditLogger()
        self.preview_generator = PPTPreviewGenerator()
        self.inputs_catalog = self._load_inputs_catalog()

        # V2 (AI-driven) generators
        self.ppt_generator_v2 = PPTGeneratorV2(self.template_repo)
        self.llm_orchestrator_v2 = LLMOrchestratorV2(self.template_repo)

    def _load_inputs_catalog(self) -> Dict[str, Dict[str, Any]]:
        catalog_path = config.INPUTS_DIR / "catalog.json"
        with catalog_path.open("r", encoding="utf-8") as f:
            items = json.load(f).get("datasets", [])
        return {item["id"]: item for item in items}

    def _slidespec_path(self, input_id: str, template_id: str) -> Path:
        return config.SLIDESPECS_DIR / f"{input_id}_{template_id}.json"

    def _slideplan_path(self, input_id: str, template_id: str) -> Path:
        return config.OUTLINE_DIR / f"{input_id}_{template_id}_slideplan.json"

    def list_inputs(self):
        return list(self.inputs_catalog.values())

    def list_templates(self, include_deprecated: bool = False):
        """List available templates."""
        return self.template_repo.list_templates(include_deprecated=include_deprecated)

    def get_input_meta(self, input_id: str) -> Dict[str, Any]:
        entry = self.inputs_catalog.get(input_id)
        if not entry:
            raise InputNotFoundError(f"Input {input_id} not found")
        return entry

    def get_template_meta(self, template_id: str) -> Dict[str, Any]:
        return self.template_repo.get_catalog_entry(template_id)

    def _get_input_path(self, input_id: str) -> Path:
        entry = self.inputs_catalog.get(input_id)
        if not entry:
            raise InputNotFoundError(f"Input {input_id} not found")
        return config.INPUTS_DIR / entry["file"]

    def load_input(self, input_id: str) -> TenantInput:
        path = self._get_input_path(input_id)
        return TenantInput.load_from_file(path)

    def generate(
        self, input_id: str, template_id: str, use_mock: bool = False
    ) -> Dict[str, Any]:
        """Generate report using V2 template.

        For V2 templates:
        - Raw TenantInput goes directly to LLM
        - AI generates content based on placeholder descriptions
        - Only key numbers are validated
        """
        tenant_input = self.load_input(input_id)

        # Ensure we only handle V2
        if not self.template_repo.is_v2(template_id):
             raise ValueError(f"Template {template_id} is not a V2 template. Only V2 templates are supported.")

        return self._generate_v2(input_id, template_id, tenant_input, use_mock=use_mock)

    def _generate_v2(
        self, input_id: str, template_id: str, tenant_input: TenantInput, use_mock: bool = False
    ) -> Dict[str, Any]:
        """Generate report using V2 AI-driven flow."""

        # V2: Direct to LLM with raw data
        slidespec: SlideSpecV2 = self.llm_orchestrator_v2.generate_slidespec_v2(
            tenant_input=tenant_input,
            template_id=template_id,
            use_mock=use_mock,
            report_id=f"{input_id}_{template_id}",
        )

        # V2: Only validate key numbers
        validator = ValidatorV2(tenant_input)
        validation_result = validator.validate_key_numbers(slidespec)
        warnings = validation_result.warnings

        if warnings:
            self.audit_logger.log(
                event="validation_warning_v2",
                details={"warnings": warnings},
                job_id=f"{input_id}:{template_id}",
                severity="warning",
            )

        # Render and save
        report_path = config.REPORTS_DIR / f"{input_id}_{template_id}.pptx"
        self.ppt_generator_v2.render(slidespec, report_path)

        slidespec_path = self._slidespec_path(input_id, template_id)
        slidespec.save(slidespec_path)
        slideplan_path = self._slideplan_path(input_id, template_id)

        self.audit_logger.log(
            event="generate_v2",
            details={"template_id": template_id, "slides_count": len(slidespec.slides)},
            job_id=f"{input_id}:{template_id}",
        )

        return {
            "job_id": f"{input_id}:{template_id}",
            "report_path": str(report_path),
            "warnings": warnings,
            "slidespec": slidespec.model_dump(),
            "slidespec_path": str(slidespec_path),
            "slideplan_path": str(slideplan_path),
            "version": "v2",
        }

    def _load_slidespec(self, input_id: str, template_id: str) -> SlideSpecV2:
        """Load slidespec, ensuring it is V2 format."""
        path = self._slidespec_path(input_id, template_id)
        if not path.exists():
            raise SlideSpecNotFoundError(f"Slidespec for {input_id}/{template_id} not found, please generate first.")

        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)

        return SlideSpecV2.model_validate(data)

    def rewrite(
        self, job_id: str, slide_key: str, new_content: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Rewrite a slide with new content."""
        try:
            input_id, template_id = job_id.split(":", 1)
        except ValueError as e:
            raise ValueError("job_id must be formatted as input_id:template_id") from e

        if not self.template_repo.is_v2(template_id):
            raise ValueError(f"Template {template_id} is not V2. Rewrite only supported for V2.")

        # For V2, just update the placeholder directly
        slidespec = self._load_slidespec(input_id, template_id)
        
        slide = slidespec.get_slide(slide_key)
        if slide:
            slide.placeholders.update(new_content)

        report_path = config.REPORTS_DIR / f"{input_id}_{template_id}.pptx"
        slidespec.save(self._slidespec_path(input_id, template_id))
        self.ppt_generator_v2.render(slidespec, report_path)

        return {
            "job_id": job_id,
            "slide_key": slide_key,
            "report_path": str(report_path),
            "warnings": [],
            "slidespec": slidespec.model_dump(),
            "version": "v2",
        }

    def read_logs(self, limit: int = 100) -> str:
        path = self.audit_logger.log_path
        if not path.exists():
            return ""
        lines = path.read_text(encoding="utf-8").splitlines()
        if limit <= 0:
            return "\n".join(lines)
        return "\n".join(lines[-limit:])

    def preview(self, job_id: str, regenerate_if_missing: bool = True) -> Dict[str, Any]:
        try:
            input_id, template_id = job_id.split(":", 1)
        except ValueError as e:
            raise ValueError("job_id must be formatted as input_id:template_id") from e

        report_path = self.get_report_path(job_id, regenerate_if_missing=regenerate_if_missing)

        # Work on a temp copy to avoid locks on the report file
        tmp_dir = config.PREVIEWS_DIR / "tmp"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        job_dir = sanitize_job_id(job_id)
        tmp_copy = tmp_dir / f"{job_dir}.pptx"
        shutil.copyfile(report_path, tmp_copy)
        base_url_prefix = f"/static/previews/{job_dir}"

        slides_count = None
        try:
            slidespec = self._load_slidespec(input_id, template_id)
            slides_count = len(slidespec.slides)
        except Exception:
            slides_count = None

        # Generate physical image files for the PPTX
        images = self.preview_generator.to_images(tmp_copy, job_id)

        if slides_count is None:
            slides_count = len(images)

        urls: list[str] = []
        if slides_count <= 0:
            urls = [f"{base_url_prefix}/{img_path.name}" for img_path in images]
        else:
            for idx in range(slides_count):
                img_idx = idx if idx < len(images) else len(images) - 1
                img_path = images[img_idx]
                urls.append(f"{base_url_prefix}/{img_path.name}")

        return {"job_id": job_id, "images": urls}

    def get_report_path(self, job_id: str, regenerate_if_missing: bool = True) -> Path:
        """Return the generated PPTX path for a job, optionally regenerating it from the saved SlideSpec."""
        try:
            input_id, template_id = job_id.split(":", 1)
        except ValueError as e:
            raise ValueError("job_id must be formatted as input_id:template_id") from e

        report_path = config.REPORTS_DIR / f"{input_id}_{template_id}.pptx"
        if not report_path.exists() and regenerate_if_missing:
            slidespec = self._load_slidespec(input_id, template_id)
            self.ppt_generator_v2.render(slidespec, report_path)

        if not report_path.exists():
            raise SlideSpecNotFoundError(f"PPT not found for {job_id}, generate first.")

        return report_path

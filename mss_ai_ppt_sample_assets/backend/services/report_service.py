from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any, Dict, Union

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec, SlideSpecV2
from mss_ai_ppt_sample_assets.backend.modules import (
    AuditLogger,
    LLMOrchestrator,
    PPTGenerator,
    TemplateRepository,
    Validator,
)
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.modules.ppt_generator import PPTGeneratorV2
from mss_ai_ppt_sample_assets.backend.modules.validator import ValidatorV2
from mss_ai_ppt_sample_assets.backend.modules.preview_generator import (
    PPTPreviewGenerator,
    sanitize_job_id,
)
from mss_ai_ppt_sample_assets.backend.modules.data_prep import DataPrepResult, prepare_facts_for_template


class InputNotFoundError(Exception):
    pass


class SlideSpecNotFoundError(Exception):
    pass


class ReportService:
    """Orchestrates report generation for both V1 (legacy) and V2 (AI-driven) templates."""

    def __init__(self):
        self.template_repo = TemplateRepository()
        self.audit_logger = AuditLogger()
        self.preview_generator = PPTPreviewGenerator()
        self.inputs_catalog = self._load_inputs_catalog()

        # V1 (legacy) generators
        self.ppt_generator = PPTGenerator(self.template_repo)
        self.llm_orchestrator = LLMOrchestrator(self.template_repo)

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
        self, input_id: str, template_id: str
    ) -> Dict[str, Any]:
        """Generate report, auto-detecting V1 or V2 template.

        For V2 templates:
        - Raw TenantInput goes directly to LLM
        - AI generates content based on placeholder descriptions
        - Only key numbers are validated

        For V1 templates (legacy):
        - data_prep prepares deterministic content
        - LLM does minimal processing
        - Full schema validation
        """
        tenant_input = self.load_input(input_id)

        # Check if this is a V2 template
        if self.template_repo.is_v2(template_id):
            return self._generate_v2(input_id, template_id, tenant_input)
        else:
            return self._generate_v1(input_id, template_id, tenant_input)

    def _generate_v2(
        self, input_id: str, template_id: str, tenant_input: TenantInput
    ) -> Dict[str, Any]:
        """Generate report using V2 AI-driven flow."""

        # V2: Direct to LLM with raw data
        slidespec: SlideSpecV2 = self.llm_orchestrator_v2.generate_slidespec_v2(
            tenant_input=tenant_input,
            template_id=template_id,
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
            "version": "v2",
        }

    def _generate_v1(
        self, input_id: str, template_id: str, tenant_input: TenantInput
    ) -> Dict[str, Any]:
        """Generate report using V1 legacy flow."""
        template = self.template_repo.get_descriptor(template_id)
        prepared: DataPrepResult = prepare_facts_for_template(tenant_input, template)

        slidespec: SlideSpec = self.llm_orchestrator.generate_slidespec(
            input_id=input_id,
            template_id=template_id,
            prepared=prepared,
            use_mock=False,
        )

        validator = Validator(template, facts=prepared.facts)
        schema_result = validator.validate_schema(slidespec)
        fact_result = validator.fact_check(slidespec)

        warnings = schema_result.issues + schema_result.warnings + fact_result.warnings
        if warnings:
            self.audit_logger.log(
                event="validation_warning",
                details={"warnings": warnings},
                job_id=f"{input_id}:{template_id}",
                severity="warning",
            )

        report_path = config.REPORTS_DIR / f"{input_id}_{template_id}.pptx"
        self.ppt_generator.render(slidespec, report_path)
        slidespec_path = self._slidespec_path(input_id, template_id)
        slidespec.save(slidespec_path)

        return {
            "job_id": f"{input_id}:{template_id}",
            "report_path": str(report_path),
            "warnings": warnings,
            "slidespec": slidespec.model_dump(),
            "slidespec_path": str(slidespec_path),
            "version": "v1",
        }

    def _load_slidespec(self, input_id: str, template_id: str) -> Union[SlideSpec, SlideSpecV2]:
        """Load slidespec, auto-detecting V1 or V2 format."""
        path = self._slidespec_path(input_id, template_id)
        if not path.exists():
            raise SlideSpecNotFoundError(f"Slidespec for {input_id}/{template_id} not found, please generate first.")

        # Try to load and detect format
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)

        # V2 slidespecs have 'placeholders' in slides, V1 has 'data'
        if data.get("slides") and len(data["slides"]) > 0:
            first_slide = data["slides"][0]
            if "placeholders" in first_slide:
                return SlideSpecV2.model_validate(data)

        return SlideSpec.model_validate(data)

    def rewrite(
        self, job_id: str, slide_key: str, new_content: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Rewrite a slide with new content (V1 only for now)."""
        try:
            input_id, template_id = job_id.split(":", 1)
        except ValueError as e:
            raise ValueError("job_id must be formatted as input_id:template_id") from e

        # V2 rewrite would require different handling
        if self.template_repo.is_v2(template_id):
            # For V2, just update the placeholder directly
            slidespec = self._load_slidespec(input_id, template_id)
            if isinstance(slidespec, SlideSpecV2):
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

        # V1 flow
        template = self.template_repo.get_descriptor(template_id)
        slidespec = self._load_slidespec(input_id, template_id)
        if not isinstance(slidespec, SlideSpec):
            raise ValueError(f"Expected V1 slidespec for template {template_id}")

        updated = self.llm_orchestrator.rewrite_slide(slidespec, slide_key, new_content)

        validator = Validator(template)
        validation = validator.validate_schema(updated)
        warnings = validation.issues + validation.warnings
        if warnings:
            self.audit_logger.log(
                event="validation_warning_rewrite",
                details={"warnings": warnings, "slide_key": slide_key},
                job_id=job_id,
                slide_key=slide_key,
                severity="warning",
            )

        report_path = config.REPORTS_DIR / f"{input_id}_{template_id}.pptx"
        updated.save(self._slidespec_path(input_id, template_id))
        self.ppt_generator.render(updated, report_path)

        self.audit_logger.log(
            event="rewrite",
            details={"slide_key": slide_key, "fields": list(new_content.keys())},
            job_id=job_id,
            slide_key=slide_key,
        )

        return {
            "job_id": job_id,
            "slide_key": slide_key,
            "report_path": str(report_path),
            "warnings": warnings,
            "slidespec": updated.model_dump(),
            "version": "v1",
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

        report_path = config.REPORTS_DIR / f"{input_id}_{template_id}.pptx"
        if not report_path.exists() and regenerate_if_missing:
            slidespec = self._load_slidespec(input_id, template_id)
            if isinstance(slidespec, SlideSpecV2):
                self.ppt_generator_v2.render(slidespec, report_path)
            else:
                self.ppt_generator.render(slidespec, report_path)

        if not report_path.exists():
            raise SlideSpecNotFoundError(f"PPT not found for {job_id}, generate first.")

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

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any, Dict

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules import (
    AuditLogger,
    LLMOrchestrator,
    PPTGenerator,
    TemplateRepository,
    Validator,
)
from mss_ai_ppt_sample_assets.backend.modules.preview_generator import (
    PPTPreviewGenerator,
    sanitize_job_id,
)
from mss_ai_ppt_sample_assets.backend.modules.data_prep import DataPrepResult, prepare_facts_for_template
from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec


class InputNotFoundError(Exception):
    pass


class SlideSpecNotFoundError(Exception):
    pass


class ReportService:
    def __init__(self):
        self.template_repo = TemplateRepository()
        self.audit_logger = AuditLogger()
        self.ppt_generator = PPTGenerator(self.template_repo)
        self.llm_orchestrator = LLMOrchestrator(self.template_repo)
        self.preview_generator = PPTPreviewGenerator()
        self.inputs_catalog = self._load_inputs_catalog()

    def _load_inputs_catalog(self) -> Dict[str, Dict[str, Any]]:
        catalog_path = config.INPUTS_DIR / "catalog.json"
        with catalog_path.open("r", encoding="utf-8") as f:
            items = json.load(f).get("datasets", [])
        return {item["id"]: item for item in items}

    def _slidespec_path(self, input_id: str, template_id: str) -> Path:
        return config.SLIDESPECS_DIR / f"{input_id}_{template_id}.json"

    def list_inputs(self):
        return list(self.inputs_catalog.values())

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
        tenant_input = self.load_input(input_id)
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
        }

    def _load_slidespec(self, input_id: str, template_id: str) -> SlideSpec:
        path = self._slidespec_path(input_id, template_id)
        if not path.exists():
            raise SlideSpecNotFoundError(f"Slidespec for {input_id}/{template_id} not found, please generate first.")
        return SlideSpec.load_from_file(path)

    def rewrite(
        self, job_id: str, slide_key: str, new_content: Dict[str, Any]
    ) -> Dict[str, Any]:
        try:
            input_id, template_id = job_id.split(":", 1)
        except ValueError as e:
            raise ValueError("job_id must be formatted as input_id:template_id") from e

        template = self.template_repo.get_descriptor(template_id)
        slidespec = self._load_slidespec(input_id, template_id)
        updated = self.llm_orchestrator.rewrite_slide(slidespec, slide_key, new_content)

        # Revalidate and render
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
            # If slidespec cannot be loaded, we will fall back to however many
            # images the preview pipeline produces.
            slides_count = None

        # Generate physical image files for the PPTX
        images = self.preview_generator.to_images(tmp_copy, job_id)

        # Determine how many slides the logical slidespec has,
        # so we can align preview URLs with slide_no.
        if slides_count is None:
            slides_count = len(images)

        urls: list[str] = []
        if slides_count <= 0:
            # No logical slides; just expose whatever was generated
            urls = [f"{base_url_prefix}/{img_path.name}" for img_path in images]
        else:
            # For each logical slide, map to a physical PNG.
            # If LibreOffice only produced a single PNG, reuse it for all slides
            # so that at least something is shown for every slide card.
            for idx in range(slides_count):
                img_idx = idx if idx < len(images) else len(images) - 1
                img_path = images[img_idx]
                urls.append(f"{base_url_prefix}/{img_path.name}")

        return {"job_id": job_id, "images": urls}

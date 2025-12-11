from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpec
from mss_ai_ppt_sample_assets.backend.models.templates import TemplateDescriptor

try:
    import jsonschema
except ImportError:  # pragma: no cover - optional dependency
    jsonschema = None


@dataclass
class ValidationResult:
    is_valid: bool
    issues: List[str]
    warnings: List[str]


class Validator:
    def __init__(self, template: TemplateDescriptor, facts: Optional[Dict[str, Any]] = None):
        self.template = template
        self.facts = facts or {}

    def validate_schema(self, slidespec: SlideSpec) -> ValidationResult:
        issues: List[str] = []
        warnings: List[str] = []

        slides_by_key = {s.slide_key: s for s in slidespec.slides}
        for slide_desc in self.template.slides:
            if slide_desc.slide_key not in slides_by_key:
                issues.append(f"slide {slide_desc.slide_key} missing")
                continue
            slide = slides_by_key[slide_desc.slide_key]
            schema = slide_desc.schema
            if schema and jsonschema:
                try:
                    jsonschema.validate(instance=slide.data, schema=schema)
                except jsonschema.ValidationError as ve:  # type: ignore
                    issues.append(f"{slide.slide_key}: {ve.message}")
            elif schema:
                # Minimal required field check when jsonschema is unavailable
                required = schema.get("required", [])
                for field in required:
                    if field not in slide.data:
                        issues.append(f"{slide.slide_key}: missing field {field}")

        return ValidationResult(is_valid=len(issues) == 0, issues=issues, warnings=warnings)

    def fact_check(self, slidespec: SlideSpec) -> ValidationResult:
        issues: List[str] = []
        warnings: List[str] = []
        if not self.facts:
            return ValidationResult(True, issues, warnings)

        # Simple numeric/string equality checks
        key_fields = {
            "alerts_total": ("executive_summary", "kpis", "alerts_total"),
            "incidents_high": ("executive_summary", "kpis", "incidents_high"),
            "false_positive_rate": ("executive_summary", "kpis", "false_positive_rate"),
        }
        slides_by_key = {s.slide_key: s for s in slidespec.slides}
        for fact_key, path in key_fields.items():
            slide_key, top_field, subfield = path
            expected = self.facts.get(fact_key)
            slide = slides_by_key.get(slide_key)
            if not slide:
                continue
            value = None
            if isinstance(slide.data.get(top_field), dict):
                value = slide.data[top_field].get(subfield)
            if expected is not None and value is not None and str(value) != str(expected):
                warnings.append(
                    f"{slide_key}.{top_field}.{subfield} = {value} differs from fact {expected}"
                )

        return ValidationResult(is_valid=len(issues) == 0, issues=issues, warnings=warnings)

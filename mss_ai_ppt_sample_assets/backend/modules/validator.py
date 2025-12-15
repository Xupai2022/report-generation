from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideSpecV2
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput


@dataclass
class ValidationResult:
    is_valid: bool
    issues: List[str]
    warnings: List[str]


# ============================================================================
# V2 Validator - Simplified, only validates key numbers
# ============================================================================

class ValidatorV2:
    """Simplified validator for V2 templates.

    Only validates that key numerical fields in AI-generated content
    match the input data. No schema validation needed for V2.
    """

    # Key fields to validate: (token, input_path, computed_key)
    KEY_FIELDS = [
        ("KPI_ALERTS_TOTAL", "alerts.total", None),
        ("KPI_ALERTS_HIGH", "alerts.by_severity.high", None),
        ("KPI_INCIDENTS_COUNT", None, "incidents_count"),
        ("KPI_INCIDENTS_HIGH", None, "incidents_high_count"),
        ("KPI_MTTR_HOURS", "mss_ops.mttr_hours_avg", None),
        ("KPI_VULN_CRITICAL", "vulnerabilities.counts.critical", None),
        ("KPI_VULN_HIGH", "vulnerabilities.counts.high", None),
    ]

    def __init__(self, tenant_input: TenantInput):
        self.tenant_input = tenant_input
        self._computed = self._compute_derived_values()

    def _compute_derived_values(self) -> Dict[str, Any]:
        """Compute derived values from input data."""
        incidents = self.tenant_input.get("incidents", []) or []
        return {
            "incidents_count": len(incidents),
            "incidents_high_count": len([i for i in incidents if i.get("severity") == "high"]),
        }

    def _get_nested(self, data: Dict[str, Any], path: str) -> Any:
        """Get nested value using dot notation."""
        current = data
        for part in path.split("."):
            if isinstance(current, dict):
                current = current.get(part)
            else:
                return None
        return current

    def _extract_number(self, value: Any) -> Optional[float]:
        """Extract number from a value (handles strings like '12414' or '2.8小时')."""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            numbers = re.findall(r'[\d.]+', value)
            if numbers:
                try:
                    return float(numbers[0])
                except ValueError:
                    pass
        return None

    def validate_key_numbers(self, slidespec: SlideSpecV2) -> ValidationResult:
        """Validate that key numbers in slidespec match input data.

        Args:
            slidespec: V2 slidespec to validate

        Returns:
            ValidationResult with any mismatches as warnings
        """
        issues: List[str] = []
        warnings: List[str] = []

        for token, input_path, computed_key in self.KEY_FIELDS:
            # Get expected value
            if computed_key:
                expected = self._computed.get(computed_key)
            elif input_path:
                expected = self._get_nested(self.tenant_input, input_path)
            else:
                continue

            if expected is None:
                continue

            expected_num = self._extract_number(expected)
            if expected_num is None:
                continue

            # Find actual value in slidespec
            for slide in slidespec.slides:
                if token in slide.placeholders:
                    actual = slide.placeholders[token]
                    actual_num = self._extract_number(actual)

                    if actual_num is not None:
                        # Allow small floating point differences
                        if abs(expected_num - actual_num) > 0.01:
                            warnings.append(
                                f"{token}: expected {expected_num}, got {actual_num}"
                            )
                    break

        return ValidationResult(
            is_valid=len(issues) == 0,
            issues=issues,
            warnings=warnings
        )
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Literal, Union

from pydantic import BaseModel, Field, validator
import json


# ============================================================================
# V2 Models - AI-driven content generation
# ============================================================================

class PlaceholderDefinition(BaseModel):
    """Definition for a placeholder in V2 templates."""
    token: str  # Placeholder token name, e.g., "HEADLINE"
    type: Literal["text", "paragraph", "bullet_list", "kpi", "kpi_group", "table", "chart_data", "incident_list", "incident_detail", "bar_chart", "pie_chart", "native_table"]
    ai_generate: bool = False  # Whether content should be AI-generated

    # For ai_generate=False: direct data source
    source: Optional[str] = None  # Data path, e.g., "alerts.total"
    default: Optional[str] = None  # Default value if source is None
    format: Optional[str] = None  # Format template, e.g., "{value}小时"
    transform: Optional[str] = None  # Transform: "uppercase", "lowercase", "percent"

    # For ai_generate=True: AI instructions
    ai_instruction: Optional[str] = None  # Detailed instruction for AI
    max_length: Optional[int] = None  # Max character length
    max_items: Optional[int] = None  # Max items for list types
    max_chars_per_item: Optional[int] = None  # Max chars per list item

    # Validation
    validation: Optional[str] = None  # Field to validate against input data

    # Table-specific
    columns: Optional[List[str]] = None  # Column names for table type

    # Chart-specific configuration (for bar_chart, pie_chart)
    chart_config: Optional[Dict[str, Any]] = None  # Chart configuration
    # Expected structure for chart_config:
    # {
    #   "data_source": "alerts.trend_weekly",  # Path to data in TenantInput
    #   "x_field": "labels",  # Field name for X-axis (or categories)
    #   "y_field": "values",  # Field name for Y-axis (or values)
    #   "chart_title_ai": true,  # Whether chart title is AI-generated
    #   "position": {"left": 1.0, "top": 2.0, "width": 8.0, "height": 4.0}  # Position in inches
    # }
    # For pie_chart: data_source should point to dict like {"high": 52, "medium": 473}

    # Native table configuration (for native_table)
    table_config: Optional[Dict[str, Any]] = None  # Table configuration
    # Expected structure for table_config:
    # {
    #   "data_source": "alerts.top_rules",  # Path to list of dicts
    #   "columns": [
    #     {"header": "规则名称", "field": "name", "width": 3.0},
    #     {"header": "触发次数", "field": "count", "width": 1.5}
    #   ],
    #   "position": {"left": 1.0, "top": 2.0, "width": 8.0, "height": 3.0}
    # }


class SlideDefinitionV2(BaseModel):
    """Slide definition for V2 templates."""
    slide_no: int
    slide_key: str
    title: str
    placeholders: List[PlaceholderDefinition]


class TemplateDescriptorV2(BaseModel):
    """Template descriptor for V2 AI-driven templates."""
    template_id: str
    name: str
    version: str
    pptx_file: str
    audience: Literal["management", "technical"]
    language: str = "zh-CN"
    style: Dict[str, Any] = Field(default_factory=dict)
    slides: List[SlideDefinitionV2]

    @classmethod
    def load_from_file(cls, path: Path) -> "TemplateDescriptorV2":
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.model_validate(data)

    @validator("slides", pre=True)
    def sort_slides(cls, slides: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Ensure slides are ordered by slide_no."""
        return sorted(slides, key=lambda s: s.get("slide_no", 0))

    def get_ai_placeholders(self) -> List[tuple[str, str, PlaceholderDefinition]]:
        """Get all placeholders that require AI generation.

        Returns:
            List of (slide_key, token, placeholder_definition) tuples
        """
        result = []
        for slide in self.slides:
            for ph in slide.placeholders:
                if ph.ai_generate:
                    result.append((slide.slide_key, ph.token, ph))
        return result

    def get_data_placeholders(self) -> List[tuple[str, str, PlaceholderDefinition]]:
        """Get all placeholders that extract data directly.

        Returns:
            List of (slide_key, token, placeholder_definition) tuples
        """
        result = []
        for slide in self.slides:
            for ph in slide.placeholders:
                if not ph.ai_generate:
                    result.append((slide.slide_key, ph.token, ph))
        return result

    def get_validation_fields(self) -> Dict[str, str]:
        """Get all fields that need validation.

        Returns:
            Dict mapping token to validation field path
        """
        result = {}
        for slide in self.slides:
            for ph in slide.placeholders:
                if ph.validation:
                    result[ph.token] = ph.validation
        return result


# ============================================================================
# Utility functions
# ============================================================================

def is_v2_template(template_id: str) -> bool:
    """Check if a template is V2 (AI-driven) based on its ID."""
    return "_v2" in template_id


def load_template_descriptor(path: Path) -> TemplateDescriptorV2:
    """Load a template descriptor, ensuring it is V2 format.

    Args:
        path: Path to the descriptor JSON file

    Returns:
        TemplateDescriptorV2
    """
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    # Simplified check or just direct validation
    return TemplateDescriptorV2.model_validate(data)
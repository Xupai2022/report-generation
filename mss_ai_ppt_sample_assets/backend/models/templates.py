from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field, validator
import json


class SlideDescriptor(BaseModel):
    slide_no: int
    slide_key: str
    description: Optional[str] = None
    render_map: Dict[str, str]
    schema: Optional[Dict[str, Any]] = None
    constraints: Optional[Dict[str, Any]] = None


class TemplateDescriptor(BaseModel):
    template_id: str
    name: str
    version: str
    pptx_file: str
    token_syntax: str = "{{TOKEN}}"
    style: Dict[str, Any] = Field(default_factory=dict)
    slides: List[SlideDescriptor]

    @classmethod
    def load_from_file(cls, path: Path) -> "TemplateDescriptor":
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.parse_obj(data)

    @validator("slides", pre=True)
    def sort_slides(cls, slides: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Ensure slides are ordered by slide_no for deterministic processing."""
        return sorted(slides, key=lambda s: s.get("slide_no", 0))

from __future__ import annotations

from pathlib import Path
from typing import List

import json
from pydantic import BaseModel, Field


class SlidePlanSlide(BaseModel):
    slide_no: int
    slide_key: str
    enabled: bool = True
    title: str = ""
    intent: str = ""
    fact_refs: List[str] = Field(default_factory=list)


class SlidePlan(BaseModel):
    template_id: str
    slides: List[SlidePlanSlide]

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as f:
            json.dump(self.model_dump(), f, ensure_ascii=False, indent=2)


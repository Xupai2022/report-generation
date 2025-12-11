from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

import json
from pydantic import BaseModel


class SlideSpecItem(BaseModel):
    slide_no: int
    slide_key: str
    data: Dict[str, Any]


class SlideSpec(BaseModel):
    template_id: str
    slides: List[SlideSpecItem]

    @classmethod
    def load_from_file(cls, path: Path) -> "SlideSpec":
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.parse_obj(data)

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as f:
            json.dump(self.dict(), f, ensure_ascii=False, indent=2)


class RenderedSlide(BaseModel):
    slide_no: int
    slide_key: str
    placeholders: Dict[str, str]

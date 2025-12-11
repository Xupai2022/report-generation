from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

from mss_ai_ppt_sample_assets.backend.config import TEMPLATES_DIR
from mss_ai_ppt_sample_assets.backend.models.templates import TemplateDescriptor


class TemplateNotFoundError(Exception):
    pass


class TemplateRepository:
    """Loads and caches template descriptors from the local filesystem."""

    def __init__(self, base_dir: Path = TEMPLATES_DIR):
        self.base_dir = base_dir
        self._catalog = self._load_catalog()
        self._descriptor_cache: Dict[str, TemplateDescriptor] = {}

    def _load_catalog(self) -> List[Dict]:
        catalog_path = self.base_dir / "catalog.json"
        with catalog_path.open("r", encoding="utf-8") as f:
            return json.load(f).get("templates", [])

    def list_templates(self) -> List[Dict]:
        return self._catalog

    def get_descriptor(self, template_id: str) -> TemplateDescriptor:
        if template_id in self._descriptor_cache:
            return self._descriptor_cache[template_id]
        entry = next((t for t in self._catalog if t["template_id"] == template_id), None)
        if not entry:
            raise TemplateNotFoundError(f"Template {template_id} not found in catalog")
        descriptor_path = self.base_dir / entry["descriptor_file"]
        descriptor = TemplateDescriptor.load_from_file(descriptor_path)
        self._descriptor_cache[template_id] = descriptor
        return descriptor

    def get_pptx_path(self, template_id: str) -> Path:
        entry = next((t for t in self._catalog if t["template_id"] == template_id), None)
        if not entry:
            raise TemplateNotFoundError(f"Template {template_id} not found in catalog")
        return self.base_dir / entry["pptx_file"]

    def get_catalog_entry(self, template_id: str) -> Dict:
        entry = next((t for t in self._catalog if t["template_id"] == template_id), None)
        if not entry:
            raise TemplateNotFoundError(f"Template {template_id} not found in catalog")
        return entry

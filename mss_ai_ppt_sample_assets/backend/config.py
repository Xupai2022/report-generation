import os
from pathlib import Path
from typing import Optional


ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT_DIR / "data"
TEMPLATES_DIR = DATA_DIR / "templates"
INPUTS_DIR = DATA_DIR / "inputs"
MOCK_OUTPUTS_DIR = DATA_DIR / "mock_outputs"

OUTPUTS_DIR = ROOT_DIR / "outputs"
REPORTS_DIR = OUTPUTS_DIR / "reports"
LOGS_DIR = OUTPUTS_DIR / "logs"
PREVIEWS_DIR = OUTPUTS_DIR / "previews"
SLIDESPECS_DIR = OUTPUTS_DIR / "slidespecs"


class Settings:
    """Centralised configuration for the backend service (minimal env reader)."""

    def __init__(self):
        # OpenAI-compatible endpoint configuration
        self.openai_api_key: Optional[str] = os.getenv("OPENAI_API_KEY")
        self.openai_base_url: Optional[str] = os.getenv("OPENAI_BASE_URL")
        self.openai_model: str = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

        # Feature flags
        self.enable_llm: bool = os.getenv("ENABLE_LLM", "false").lower() == "true"
        self.default_locale: str = os.getenv("DEFAULT_LOCALE", "zh-CN")


settings = Settings()

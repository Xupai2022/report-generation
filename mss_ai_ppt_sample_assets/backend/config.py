import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

# Load .env file from project root
env_path = Path(__file__).resolve().parent.parent.parent / ".env"
load_dotenv(dotenv_path=env_path)

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
SESSIONS_DIR = OUTPUTS_DIR / "sessions"  # Isolated session directories for concurrent requests
JOBS_DIR = OUTPUTS_DIR / "jobs"  # Job state storage directory


class Settings:
    """Centralised configuration for the backend service (minimal env reader)."""

    def __init__(self):
        # OpenAI-compatible endpoint configuration
        self.openai_api_key: Optional[str] = os.getenv("OPENAI_API_KEY")
        self.openai_base_url: Optional[str] = os.getenv("OPENAI_BASE_URL")
        self.openai_model: str = os.getenv("OPENAI_MODEL", "GLM4.7")

        # Feature flags
        self.enable_llm: bool = os.getenv("ENABLE_LLM", "false").lower() == "true"
        self.default_locale: str = os.getenv("DEFAULT_LOCALE", "zh-CN")

        # Preview cleanup configuration
        self.preview_cleanup_days: int = int(os.getenv("PREVIEW_CLEANUP_DAYS", "7"))

        # Logging configuration
        self.log_level: str = os.getenv("LOG_LEVEL", "INFO").upper()
        self.log_max_bytes: int = int(os.getenv("LOG_MAX_BYTES", str(50 * 1024 * 1024)))  # 50MB default
        self.log_backup_count: int = int(os.getenv("LOG_BACKUP_COUNT", "10"))

        # Job management configuration
        self.job_retention_days: int = int(os.getenv("JOB_RETENTION_DAYS", "7"))
        self.job_max_retries: int = int(os.getenv("JOB_MAX_RETRIES", "3"))

        # Validate OpenAI configuration when LLM is enabled
        if self.enable_llm and not self.openai_api_key:
            raise ValueError(
                "OPENAI_API_KEY is required when ENABLE_LLM=true. "
                "Please set it in your .env file."
            )


settings = Settings()

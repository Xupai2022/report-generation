from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware
from pathlib import Path
import logging
import asyncio
import mimetypes

from mss_ai_ppt_sample_assets.backend.services.report_service import ReportService
from mss_ai_ppt_sample_assets.backend.services.job_manager import (
    JobManager,
    RESTART_INTERRUPTED_ERROR_MESSAGE,
)
from mss_ai_ppt_sample_assets.backend.modules.job_store import JobStore
from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.websocket_support import WebSocketManager
from mss_ai_ppt_sample_assets.backend.routers import v1_router

# Import new modules
from mss_ai_ppt_sample_assets.backend.logging_config import setup_logging
from mss_ai_ppt_sample_assets.backend.middleware import RequestIdMiddleware, ErrorLoggingMiddleware

# Setup enhanced logging with file persistence and rotation
setup_logging(
    log_level=config.settings.log_level,
    log_dir=config.LOGS_DIR,
    max_bytes=config.settings.log_max_bytes,
    backup_count=config.settings.log_backup_count
)
logger = logging.getLogger(__name__)

# Ensure SVG static files are served with the standard MIME type.
mimetypes.add_type("image/svg+xml", ".svg")

app = FastAPI(
    title="MSS AI PPT API",
    version="1.0.0",
    description="Intelligent Report Generation Platform - RESTful API",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Add middleware for request tracking and logging
app.add_middleware(ErrorLoggingMiddleware)
app.add_middleware(RequestIdMiddleware)

# Add session middleware for admin authentication
app.add_middleware(
    SessionMiddleware,
    secret_key=config.settings.admin_session_secret,
    session_cookie="admin_session",
    max_age=config.settings.admin_session_max_age,
    same_site="lax",
    https_only=False  # Set to True in production with HTTPS
)

# Configuration
MAX_CONCURRENT_LLM_REQUESTS = 5
llm_semaphore = asyncio.Semaphore(MAX_CONCURRENT_LLM_REQUESTS)

# Compact startup log with key configuration
logger.info(
    f"MSS AI PPT Backend | "
    f"LLM: {config.settings.openai_model} ({config.settings.openai_base_url}) | "
    f"Concurrency: {MAX_CONCURRENT_LLM_REQUESTS} | "
    f"Log: {config.settings.log_level} "
    f"({config.settings.log_max_bytes / (1024*1024):.0f}MB x {config.settings.log_backup_count})"
)

# Services
service = ReportService()
ws_manager = WebSocketManager()

# Job management
job_store = JobStore(config.JOBS_DIR)
job_manager = JobManager(job_store, service)

logger.info(f"Services initialized | Job retention: {config.settings.job_retention_days}d | Max retries: {config.settings.job_max_retries}")

app.mount("/static/previews", StaticFiles(directory=config.PREVIEWS_DIR), name="previews")

# Expose outputs directory (sessions/reports/jobs, etc.) for development/admin use.
# In production, prefer downloading via authenticated endpoints.
app.mount(config.OUTPUTS_URL_PREFIX, StaticFiles(directory=config.OUTPUTS_DIR), name="outputs")

# Register v1 API router
app.include_router(v1_router)

# Initialize dependencies for reports router (WebSocket and semaphore)
from mss_ai_ppt_sample_assets.backend.routers.v1 import reports, jobs, admin, ratings
reports.init_dependencies(ws_manager, llm_semaphore, MAX_CONCURRENT_LLM_REQUESTS, job_manager)
jobs.init_dependencies(job_manager)
admin.init_dependencies(job_store)
ratings.init_dependencies(job_store)

logger.info("API routes registered | WebSocket, JobManager, AdminService, RatingService initialized")

# 简单的前端静态页面（无需 npm），挂载在 /ui
FRONTEND_DIR = Path(__file__).parent / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/ui", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
    ASSETS_DIR = FRONTEND_DIR / "assets"
    if ASSETS_DIR.exists():
        app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")

# Mount i18n directory for translation files
I18N_DIR = FRONTEND_DIR / "i18n"
if I18N_DIR.exists():
    app.mount("/i18n", StaticFiles(directory=I18N_DIR), name="i18n")


# ==================== WebSocket Endpoint ====================

@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    """WebSocket connection for real-time progress updates."""
    await ws_manager.connect(websocket, client_id)
    try:
        while True:
            data = await websocket.receive_text()
            if data == "ping":
                await websocket.send_text("pong")
    except WebSocketDisconnect:
        ws_manager.disconnect(client_id)


# ==================== Root Endpoint ====================

@app.get("/")
def root():
    """Redirect root to the frontend UI."""
    from fastapi.responses import RedirectResponse
    return RedirectResponse(url="/ui/index.html")


@app.get("/api")
def api_root():
    """API root endpoint with service information."""
    return {
        "message": "MSS AI PPT API v1.0",
        "description": "Intelligent Report Generation Platform",
        "documentation": "/docs",
        "redoc": "/redoc",
        "api_version": "v1",
        "api_prefix": "/api/v1",
        "endpoints": {
            "reports": "/api/v1/reports",
            "templates": "/api/v1/templates",
            "inputs": "/api/v1/inputs",
            "sessions": "/api/v1/sessions",
            "system": "/api/v1/system",
            "rag": "/api/v1/rag",
            "websocket": "/ws/{client_id}"
        }
    }


# ==================== Startup Event ====================

@app.on_event("startup")
async def startup_cleanup():
    """Clean up old sessions, stale locks, and old jobs when server starts."""
    try:
        # Clean up old sessions (older than configured retention period)
        cleaned_count = service.cleanup_old_sessions(
            max_age_hours=config.settings.session_retention_days * 24
        )
        if cleaned_count > 0:
            logger.info(f"Startup cleanup: removed {cleaned_count} old sessions")

        # Clean up stale lock files (older than 5 minutes)
        # These locks may be left behind by crashed processes
        from mss_ai_ppt_sample_assets.backend.modules.file_lock import cleanup_stale_locks
        cleanup_stale_locks(config.OUTPUTS_DIR, max_age_seconds=300)  # 5 minutes
        logger.info("Startup cleanup: stale locks cleaned")

        # Mark interrupted running jobs as failed after restart.
        interrupted_count = job_store.mark_running_jobs_failed(
            RESTART_INTERRUPTED_ERROR_MESSAGE
        )
        if interrupted_count > 0:
            logger.info(
                f"Startup recovery: marked {interrupted_count} interrupted running jobs as failed"
            )

        # Clean up old jobs (older than configured retention period)
        job_cleaned_count = job_store.cleanup_old_jobs(days=config.settings.job_retention_days)
        if job_cleaned_count > 0:
            logger.info(f"Startup cleanup: removed {job_cleaned_count} old jobs")

    except Exception as e:
        logger.warning(f"Startup cleanup failed: {e}")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from pydantic import BaseModel
import logging
from datetime import datetime

from mss_ai_ppt_sample_assets.backend.services.report_service import (
    ReportService,
    InputNotFoundError,
    SlideSpecNotFoundError,
)
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateNotFoundError
from mss_ai_ppt_sample_assets.backend import config

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # Console output
    ]
)
logger = logging.getLogger(__name__)

app = FastAPI(title="MSS AI PPT Backend", version="0.1.0")
logger.info("Initializing MSS AI PPT Backend...")
logger.info(f"LLM Enabled: {config.settings.enable_llm}")
logger.info(f"OpenAI Model: {config.settings.openai_model}")
logger.info(f"OpenAI Base URL: {config.settings.openai_base_url}")

service = ReportService()
app.mount("/static/previews", StaticFiles(directory=config.PREVIEWS_DIR), name="previews")

# 简单的前端静态页面（无需 npm），挂载在 /ui
FRONTEND_DIR = Path(__file__).parent / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/ui", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")


class GenerateRequest(BaseModel):
    input_id: str
    template_id: str
    use_mock: bool = False


class RewriteRequest(BaseModel):
    job_id: str
    slide_key: str
    new_content: dict


class PreviewRequest(BaseModel):
    job_id: str
    regenerate_if_missing: bool = True


@app.get("/")
def root():
    return {
        "message": "MSS AI PPT backend is running.",
        "endpoints": ["/health", "/templates", "/inputs", "/generate", "/preview", "/download", "/docs"],
    }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/templates")
def list_templates():
    return service.template_repo.list_templates()


@app.get("/inputs")
def list_inputs():
    return service.list_inputs()


@app.post("/generate")
def generate(req: GenerateRequest):
    logger.info(f"=== Generate Request: input_id={req.input_id}, template_id={req.template_id}, use_mock={req.use_mock} ===")
    try:
        result = service.generate(req.input_id, req.template_id, use_mock=req.use_mock)
        logger.info(f"✓ Generation successful: {result.get('job_id')}")
        logger.info(f"  - Report path: {result.get('report_path')}")
        logger.info(f"  - Warnings: {len(result.get('warnings', []))}")
        return JSONResponse(result)
    except InputNotFoundError as e:
        logger.error(f"✗ Input not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except TemplateNotFoundError as e:
        logger.error(f"✗ Template not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.exception(f"✗ Generation failed with exception: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/rewrite")
def rewrite(req: RewriteRequest):
    try:
        result = service.rewrite(req.job_id, req.slide_key, req.new_content)
        return JSONResponse(result)
    except SlideSpecNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/preview")
def preview(job_id: str, regenerate_if_missing: bool = True):
    try:
        result = service.preview(job_id, regenerate_if_missing=regenerate_if_missing)
        return JSONResponse(result)
    except SlideSpecNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download")
def download(job_id: str, regenerate_if_missing: bool = True):
    try:
        report_path = service.get_report_path(job_id, regenerate_if_missing=regenerate_if_missing)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        download_name = f"{ts}_{report_path.name}"
        return FileResponse(
            path=report_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=download_name,
        )
    except SlideSpecNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/logs")
def logs(limit: int = 100):
    try:
        content = service.read_logs(limit=limit)
        return JSONResponse({"lines": content.splitlines() if content else []})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from mss_ai_ppt_sample_assets.backend.services.report_service import (
    ReportService,
    InputNotFoundError,
    SlideSpecNotFoundError,
)
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateNotFoundError
from mss_ai_ppt_sample_assets.backend import config

app = FastAPI(title="MSS AI PPT Backend", version="0.1.0")
service = ReportService()
app.mount("/static/previews", StaticFiles(directory=config.PREVIEWS_DIR), name="previews")


class GenerateRequest(BaseModel):
    input_id: str
    template_id: str
    use_mock: bool = True


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
        "endpoints": ["/health", "/templates", "/inputs", "/generate", "/docs"],
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
    try:
        result = service.generate(req.input_id, req.template_id, use_mock=req.use_mock)
        return JSONResponse(result)
    except InputNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except TemplateNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
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

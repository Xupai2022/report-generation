"""Reports API endpoints - RESTful report generation and management."""

from fastapi import APIRouter, HTTPException, status, Request
from fastapi.responses import FileResponse
from datetime import datetime
import asyncio
import logging
import time
from collections import defaultdict

from ...services.report_service import ReportService
from ...services.job_manager import JobManager
from ...models.job_state import JobStatus
from ...schemas.responses import SuccessResponse
from ...schemas.requests import CreateReportRequest, UpdateSlidesRequest
from ...exceptions import (
    InputNotFoundError,
    TemplateNotFoundError,
    SlideSpecNotFoundError,
    LLMGenerationError
)
from openai import RateLimitError

logger = logging.getLogger(__name__)

router = APIRouter()
service = ReportService()

# Simple in-memory rate limiting (IP-based)
# Format: {client_ip: last_request_timestamp}
_request_timestamps = defaultdict(float)
_RATE_LIMIT_SECONDS = 30  # 同一IP 30秒内只能提交一次

# Dependencies initialized by app.py
ws_manager = None
llm_semaphore = None
MAX_CONCURRENT_LLM_REQUESTS = 3
job_manager: JobManager = None


def init_dependencies(websocket_manager, semaphore, max_concurrent, manager):
    """Initialize dependencies from app module."""
    global ws_manager, llm_semaphore, MAX_CONCURRENT_LLM_REQUESTS, job_manager
    ws_manager = websocket_manager
    llm_semaphore = semaphore
    MAX_CONCURRENT_LLM_REQUESTS = max_concurrent
    job_manager = manager


async def _process_report_async(job_id: str, input_id: str, template_id: str, use_mock: bool):
    """Background task to process report generation.

    This runs asynchronously after returning response to client.
    Updates job state and sends WebSocket notifications.
    """
    job = job_manager.get_job(job_id)
    if not job:
        logger.error(f"Job not found in async processor: {job_id}")
        return

    try:
        # Acquire LLM semaphore if not using mock
        if not use_mock and llm_semaphore:
            async with llm_semaphore:
                await _do_generation(job, input_id, template_id, use_mock)
        else:
            await _do_generation(job, input_id, template_id, use_mock)

    except Exception as e:
        logger.exception(f"Job {job_id} failed: {e}")
        job_manager.fail_job(job_id, str(e))

        # Send WebSocket failure notification
        if ws_manager:
            await ws_manager.send_completion(job.session_id, {"error": str(e)}, success=False)

        # Check if should retry
        if job_manager.should_retry(job_id):
            retry_count = job.retry_count + 1
            wait_time = 2 ** retry_count  # Exponential backoff
            logger.info(f"🔄 Retrying job {job_id} in {wait_time}s (attempt {retry_count})")
            await asyncio.sleep(wait_time)
            await _process_report_async(job_id, input_id, template_id, use_mock)


async def _do_generation(job, input_id: str, template_id: str, use_mock: bool):
    """Execute the actual generation (called with semaphore acquired)."""
    job_id = job.job_id

    # Send WebSocket progress updates
    if ws_manager:
        await ws_manager.send_progress_update(
            job.session_id, 10, "正在加载模板和数据..."
        )

    # Get the current event loop to pass to synchronous code
    loop = asyncio.get_running_loop()

    # Run synchronous generation in thread pool, passing the loop
    result = await loop.run_in_executor(
        None,
        lambda: service.generate(
            input_id,
            template_id,
            use_mock=use_mock,
            session_id=job.session_id,
            ws_manager=ws_manager,
            event_loop=loop  # Pass the main event loop
        )
    )

    logger.info(f"Generation successful: {job_id}")

    # Mark job as completed
    job_manager.complete_job(job_id, result)

    # Send WebSocket completion
    if ws_manager:
        await ws_manager.send_completion(job.session_id, result, success=True)


@router.post(
    "",
    status_code=status.HTTP_202_ACCEPTED,  # Changed to 202 for async processing
    response_model=SuccessResponse,
    summary="Create Report",
    description="""
    Create a new PowerPoint report (async processing with idempotency support).

    ## New Features
    - **Idempotency**: Use `idempotency_key` to prevent duplicate processing
    - **Async Processing**: Returns immediately with job_id, check status via GET /api/v1/jobs/{job_id}/status
    - **Progress Tracking**: Monitor progress via WebSocket or status API

    ## Processing Flow
    1. Create or find existing job (idempotency check)
    2. Return job_id immediately (202 Accepted)
    3. Process generation in background
    4. Client queries job status to get results

    ## Idempotency
    If you provide an `idempotency_key`, subsequent requests with the same key will:
    - Return existing job if already completed
    - Return current progress if still running
    - Not create duplicate jobs

    ## Response Codes
    - `202`: Job created/found, processing started/continuing
    - `404`: Input data or template not found
    - `429`: Too many requests (rate limit)
    """,
    responses={
        202: {
            "description": "Job created and processing started",
            "content": {
                "application/json": {
                    "example": {
                        "data": {
                            "job_id": "abc123_20260202110000:mss_executive_v2",
                            "status": "running",
                            "message": "正在生成报告,请稍候..."
                        }
                    }
                }
            }
        },
        404: {"description": "Input or template not found"},
        429: {"description": "Too many requests - rate limit exceeded"}
    }
)
async def create_report(request: Request, req: CreateReportRequest):
    """Create a new report with async processing and idempotency support."""
    # Simple IP-based rate limiting
    client_ip = request.client.host if request.client else "unknown"
    current_time = time.time()
    last_request_time = _request_timestamps.get(client_ip, 0)

    if current_time - last_request_time < _RATE_LIMIT_SECONDS:
        remaining = int(_RATE_LIMIT_SECONDS - (current_time - last_request_time))
        logger.warning(f"Rate limit exceeded for IP {client_ip}, retry in {remaining}s")
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail=f"请求过于频繁，请 {remaining} 秒后再试"
        )

    # Update request timestamp
    _request_timestamps[client_ip] = current_time

    # Clean up old entries (older than 1 hour)
    cutoff_time = current_time - 3600
    expired_ips = [ip for ip, ts in _request_timestamps.items() if ts < cutoff_time]
    for ip in expired_ips:
        del _request_timestamps[ip]

    logger.info(
        f"=== POST /api/v1/reports: client_ip={client_ip}, input_id={req.input_id}, "
        f"template_id={req.template_id}, use_mock={req.use_mock}, idempotency_key={req.idempotency_key}, "
        f"client_id={req.client_id}, session_id={req.session_id} ==="
    )

    if not job_manager:
        raise HTTPException(status_code=500, detail="Job manager not initialized")

    try:
        # Create or find existing job (idempotency)
        job = job_manager.create_job(
            input_id=req.input_id,
            template_id=req.template_id,
            idempotency_key=req.idempotency_key,
            session_id=req.session_id
        )

        # Register WebSocket if provided
        if ws_manager and req.client_id:
            ws_manager.register_session(job.session_id, req.client_id)
            logger.info(f"Registered WebSocket: session={job.session_id}, client={req.client_id}")
        else:
            logger.warning(f"WebSocket NOT registered: ws_manager={ws_manager is not None}, client_id={req.client_id}")

        # If job already completed, return cached result
        if job.status == JobStatus.COMPLETED:
            logger.info(f"Job already completed (idempotency): {job.job_id}")
            return SuccessResponse(data={
                "job_id": job.job_id,
                "status": "completed",
                "report_path": job.report_path,
                "slidespec_path": job.slidespec_path,
                "from_cache": True,
                "message": "作业已完成（使用缓存结果）"
            })

        # If job is running, return current progress
        if job.status == JobStatus.RUNNING:
            logger.info(f"Job already running: {job.job_id}")
            # Format message based on progress
            if job.progress > 0:
                message = f"作业正在处理中 ({job.progress}%)"
            else:
                message = "作业正在处理中，请稍候..."

            return SuccessResponse(data={
                "job_id": job.job_id,
                "status": "running",
                "progress": job.progress,
                "message": message,
                "check_status_url": f"/api/v1/jobs/{job.job_id}/status"
            })

        # Start new job
        job_manager.start_job(job.job_id)

        # Send initial WebSocket notification
        if ws_manager and req.client_id:
            await ws_manager.send_progress_update(
                job.session_id, 0, "请求已接收，正在排队...",
                {"template": req.template_id}
            )

        # Launch background task
        asyncio.create_task(
            _process_report_async(job.job_id, req.input_id, req.template_id, req.use_mock)
        )

        logger.info(f"Job created and processing started: {job.job_id}")
        return SuccessResponse(data={
            "job_id": job.job_id,
            "session_id": job.session_id,
            "status": "running",
            "message": "正在生成报告,请稍候...",
            "check_status_url": f"/api/v1/jobs/{job.job_id}/status"
        })

    except InputNotFoundError as e:
        logger.error(f"Input not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except TemplateNotFoundError as e:
        logger.error(f"Template not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.exception(f"Failed to create job: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get(
    "/{report_id}/download",
    summary="Download Report",
    description="""
    Download the generated PowerPoint report file.

    The `report_id` is the job_id returned from POST /api/v1/reports.
    """,
    responses={
        200: {
            "description": "Report file downloaded successfully",
            "content": {"application/vnd.openxmlformats-officedocument.presentationml.presentation": {}}
        },
        404: {"description": "Report not found"}
    }
)
async def download_report(
    report_id: str,
    regenerate_if_missing: bool = True
):
    """Download report file."""
    try:
        report_path = service.get_report_path(report_id, regenerate_if_missing=regenerate_if_missing)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        download_name = f"{ts}_{report_path.name}"

        logger.info(f"Downloading report: {report_id}")
        return FileResponse(
            path=report_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=download_name,
        )
    except SlideSpecNotFoundError as e:
        logger.error(f"Report not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        logger.error(f"Invalid request: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.exception(f"Download failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get(
    "/{report_id}/download-pdf",
    summary="Download Report as PDF",
    description="""
    Download the generated report as a continuous PDF file.

    The system converts the PowerPoint report to PDF using LibreOffice.
    The PDF file is cached and reused if available.

    The `report_id` is the job_id returned from POST /api/v1/reports.
    """,
    responses={
        200: {
            "description": "PDF file downloaded successfully",
            "content": {"application/pdf": {}}
        },
        404: {"description": "Report not found"}
    }
)
async def download_report_pdf(
    report_id: str,
    regenerate_if_missing: bool = True
):
    """Download report as PDF file."""
    try:
        pdf_path = service.get_pdf_path(report_id, regenerate_if_missing=regenerate_if_missing)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        # Extract base name without extension from original report
        base_name = pdf_path.stem.replace(".pdf", "")
        download_name = f"{ts}_{base_name}.pdf"

        logger.info(f"Downloading PDF: {report_id}")
        return FileResponse(
            path=pdf_path,
            media_type="application/pdf",
            filename=download_name,
        )
    except SlideSpecNotFoundError as e:
        logger.error(f"Report not found for PDF: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        logger.error(f"Invalid request: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.exception(f"PDF download failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get(
    "/{report_id}/preview",
    response_model=SuccessResponse,
    summary="Get Report Preview",
    description="""
    Get preview images (PNG) for all slides in the report.

    The system converts PPT → PDF → PNG for preview generation.
    Preview images are cached and reused if available.
    """,
    responses={
        200: {
            "description": "Preview images generated successfully",
            "content": {
                "application/json": {
                    "example": {
                        "data": {
                            "preview_urls": [
                                "/static/previews/job123/slide1.png",
                                "/static/previews/job123/slide2.png"
                            ]
                        }
                    }
                }
            }
        },
        404: {"description": "Report not found"}
    }
)
async def preview_report(
    report_id: str,
    regenerate_if_missing: bool = True,
    force_regenerate: bool = False,
):
    """Get report preview images."""
    try:
        result = service.preview(
            report_id,
            regenerate_if_missing=regenerate_if_missing,
            force_regenerate=force_regenerate,
        )
        logger.info(f"Preview generated for: {report_id}")
        return SuccessResponse(data=result)
    except SlideSpecNotFoundError as e:
        logger.error(f"Report not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.exception(f"Preview generation failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.patch(
    "/{report_id}/slides",
    response_model=SuccessResponse,
    summary="Update Report Slides",
    description="""
    Batch update multiple slides in an existing report.

    This endpoint allows you to modify specific placeholders in one or more slides
    without regenerating the entire report.
    """,
    responses={
        200: {
            "description": "Slides updated successfully",
            "content": {
                "application/json": {
                    "example": {
                        "data": {
                            "job_id": "tenant_acme:mss_executive_v2",
                            "updated_slides": ["cover", "summary"],
                            "updated_count": 2,
                            "warnings": [],
                            "report_path": "outputs/sessions/abc123/report.pptx"
                        }
                    }
                }
            }
        },
        404: {"description": "Report not found"},
        400: {"description": "Invalid request"}
    }
)
async def update_slides(report_id: str, req: UpdateSlidesRequest):
    """Batch update report slides."""
    try:
        # Convert UpdateSlidesRequest to service format
        slides_data = [{"slide_key": s.slide_key, "new_content": s.new_content} for s in req.slides]

        result = service.rewrite(
            job_id=report_id,
            slides=slides_data
        )

        logger.info(f"Updated {len(req.slides)} slides in report: {report_id}")
        return SuccessResponse(data=result)
    except SlideSpecNotFoundError as e:
        logger.error(f"Report not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        logger.error(f"Invalid request: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.exception(f"Slide update failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

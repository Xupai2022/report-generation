"""Templates API endpoints - Template catalog and metadata retrieval."""

from fastapi import APIRouter, HTTPException, status
from typing import Optional
import logging

from ...services.report_service import ReportService
from ...schemas.responses import SuccessResponse
from ...exceptions import TemplateNotFoundError

logger = logging.getLogger(__name__)

router = APIRouter()
service = ReportService()


@router.get(
    "",
    response_model=SuccessResponse,
    summary="List Templates",
    description="""
    Retrieve a list of all available report templates.

    ## Filtering Options
    - `audience`: Filter by target audience (e.g., "management", "technical")
    - `language`: Filter by template language (e.g., "zh-CN", "en-US")

    ## Response
    Returns an array of template metadata including:
    - template_id: Unique identifier
    - name: Display name
    - version: Template version (e.g., "v2")
    - audience: Target audience
    - language: Template language
    - slides_count: Number of slides
    - description: Template description
    """,
    responses={
        200: {
            "description": "List of available templates",
            "content": {
                "application/json": {
                    "example": {
                        "data": [
                            {
                                "template_id": "mss_executive_v2",
                                "name": "MSS Executive Report V2",
                                "version": "v2",
                                "audience": "management",
                                "language": "zh-CN",
                                "slides_count": 8,
                                "description": "Management-oriented security report"
                            },
                            {
                                "template_id": "mss_technical_v2",
                                "name": "MSS Technical Report V2",
                                "version": "v2",
                                "audience": "technical",
                                "language": "zh-CN",
                                "slides_count": 10,
                                "description": "Technical security analysis report"
                            }
                        ]
                    }
                }
            }
        }
    }
)
async def list_templates(
    audience: Optional[str] = None,
    language: Optional[str] = None
):
    """List all available templates with optional filtering."""
    try:
        templates = service.template_repo.list_templates()

        # Apply filters if provided
        if audience:
            templates = [t for t in templates if t.get("audience") == audience]
        if language:
            templates = [t for t in templates if t.get("language") == language]

        logger.info(f"✓ Listed {len(templates)} templates (audience={audience}, language={language})")
        return SuccessResponse(data=templates)
    except Exception as e:
        logger.exception(f"✗ Failed to list templates: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get(
    "/{template_id}",
    response_model=SuccessResponse,
    summary="Get Template Details",
    description="""
    Retrieve detailed information about a specific template.

    ## Response
    Returns comprehensive template metadata including:
    - Basic info (id, name, version, audience, language)
    - Slide structure (slides array with slide_no, slide_key, placeholder definitions)
    - Placeholder definitions (token, type, ai_generate, source, ai_instruction)
    - Chart and table configurations

    ## Use Cases
    - Preview template structure before generation
    - Validate input data requirements
    - Understand placeholder types and expected data sources
    """,
    responses={
        200: {
            "description": "Template details retrieved successfully",
            "content": {
                "application/json": {
                    "example": {
                        "data": {
                            "template_id": "mss_executive_v2",
                            "name": "MSS Executive Report V2",
                            "version": "v2",
                            "audience": "management",
                            "language": "zh-CN",
                            "slides_count": 8,
                            "slides": [
                                {
                                    "slide_no": 1,
                                    "slide_key": "cover",
                                    "placeholders": [
                                        {
                                            "token": "TITLE",
                                            "type": "text",
                                            "ai_generate": True,
                                            "ai_instruction": "Generate report title"
                                        },
                                        {
                                            "token": "DATE",
                                            "type": "text",
                                            "ai_generate": False,
                                            "source": "period.start"
                                        }
                                    ]
                                }
                            ]
                        }
                    }
                }
            }
        },
        404: {"description": "Template not found"}
    }
)
async def get_template(template_id: str):
    """Get detailed information about a specific template."""
    try:
        # Get template descriptor
        template_descriptor = service.template_repo.load_descriptor(template_id)

        if not template_descriptor:
            raise TemplateNotFoundError(f"Template '{template_id}' not found")

        # Get template metadata from catalog
        templates = service.template_repo.list_templates()
        template_meta = next((t for t in templates if t.get("id") == template_id), None)

        # Build response with both metadata and descriptor
        response_data = {
            "template_id": template_id,
            "slides_count": len(template_descriptor.slides),
            "slides": [
                {
                    "slide_no": slide.slide_no,
                    "slide_key": slide.slide_key,
                    "placeholders": [
                        {
                            "token": ph.token,
                            "type": ph.type,
                            "ai_generate": ph.ai_generate,
                            "source": ph.source,
                            "ai_instruction": ph.ai_instruction,
                            "chart_config": ph.chart_config,
                            "table_config": ph.table_config
                        }
                        for ph in slide.placeholders
                    ]
                }
                for slide in template_descriptor.slides
            ]
        }

        # Merge catalog metadata if available
        if template_meta:
            response_data.update({
                "name": template_meta.get("name"),
                "version": template_meta.get("version"),
                "audience": template_meta.get("audience"),
                "language": template_meta.get("language"),
                "description": template_meta.get("description")
            })

        logger.info(f"✓ Retrieved template details: {template_id}")
        return SuccessResponse(data=response_data)

    except TemplateNotFoundError as e:
        logger.error(f"✗ Template not found: {e}")
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.exception(f"✗ Failed to get template details: {e}")
        raise HTTPException(status_code=500, detail=str(e))

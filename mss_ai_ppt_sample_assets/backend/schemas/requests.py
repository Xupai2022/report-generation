"""Unified request models for API endpoints."""

from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any


class CreateReportRequest(BaseModel):
    """Request model for creating a new report."""
    input_id: str = Field(..., description="Input data ID", example="tenant_acme_2025-11")
    template_id: str = Field(..., description="Template ID", example="mss_executive_v2")
    use_mock: bool = Field(False, description="Use mock mode (skip AI generation)")
    session_id: Optional[str] = Field(None, description="Optional session ID")
    client_id: Optional[str] = Field(None, description="Optional WebSocket client ID")
    idempotency_key: Optional[str] = Field(
        None,
        description="Idempotency key to prevent duplicate processing of the same request"
    )

    class Config:
        json_schema_extra = {
            "example": {
                "input_id": "tenant_acme_2025-11",
                "template_id": "mss_executive_v2",
                "use_mock": False,
                "idempotency_key": "user123-request456"
            }
        }


class SlideUpdate(BaseModel):
    """Single slide update information."""
    slide_key: str = Field(..., description="Slide identifier", example="cover")
    new_content: Dict[str, Any] = Field(..., description="New content for placeholders")

    class Config:
        json_schema_extra = {
            "example": {
                "slide_key": "cover",
                "new_content": {
                    "TITLE": "Updated Title",
                    "SUBTITLE": "New Subtitle"
                }
            }
        }


class UpdateSlidesRequest(BaseModel):
    """Request model for batch updating slides."""
    slides: List[SlideUpdate] = Field(..., description="List of slide updates")

    class Config:
        json_schema_extra = {
            "example": {
                "slides": [
                    {
                        "slide_key": "cover",
                        "new_content": {"TITLE": "Updated Title"}
                    },
                    {
                        "slide_key": "summary",
                        "new_content": {"SUMMARY_TEXT": "New summary"}
                    }
                ]
            }
        }

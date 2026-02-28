"""Unified request models for API endpoints."""

from pydantic import BaseModel, Field, validator
from typing import Optional, List, Dict, Any, Literal


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


class AISlideRewriteRequest(BaseModel):
    """Request model for AI rewrite of a single slide."""
    slide_key: str = Field(..., description="Slide identifier", example="summary")
    user_prompt: str = Field(
        ...,
        min_length=1,
        max_length=2000,
        description="User preference/instruction for AI rewrite of the selected slide"
    )
    target_tokens: Optional[List[str]] = Field(
        None,
        description=(
            "Optional AI placeholder tokens to rewrite. "
            "If omitted or empty, backend rewrites all ai_generate=true tokens on the slide."
        ),
    )

    @validator("user_prompt")
    def validate_user_prompt(cls, v: str):
        """Validate user prompt is not empty after trimming."""
        if not v or not v.strip():
            raise ValueError("user_prompt cannot be empty")
        return v.strip()

    @validator("target_tokens", pre=True)
    def validate_target_tokens(cls, v):
        """Normalize target token list (trim, dedupe, drop empty)."""
        if v is None:
            return None
        if not isinstance(v, list):
            raise ValueError("target_tokens must be a list of strings")

        normalized: List[str] = []
        seen = set()
        for item in v:
            if not isinstance(item, str):
                raise ValueError("target_tokens must contain only strings")
            token = item.strip()
            if not token or token in seen:
                continue
            seen.add(token)
            normalized.append(token)
        return normalized

    class Config:
        json_schema_extra = {
            "example": {
                "slide_key": "summary",
                "user_prompt": "请聚焦本月高风险告警的业务影响，语气偏管理层，并给出三条可执行建议。",
                "target_tokens": ["trust_assurance", "security_trust"]
            }
        }

class SubmitRatingRequest(BaseModel):
    """User rating submission request."""
    rating: Literal["liked", "disliked"] = Field(
        ...,
        description="Rating value: 'liked' (thumbs up) or 'disliked' (thumbs down)"
    )
    comment: Optional[str] = Field(
        None,
        max_length=500,
        description="Comment explaining the rating (required for disliked)"
    )

    @validator('comment')
    def validate_comment_for_disliked(cls, v, values):
        """Validate that disliked ratings have a comment."""
        rating = values.get('rating')
        if rating == 'disliked':
            if not v or len(v.strip()) < 10:
                raise ValueError('Comment is required for thumbs down (minimum 10 characters)')
        return v

    class Config:
        json_schema_extra = {
            "example": {
                "rating": "disliked",
                "comment": "The chart colors are hard to read and data formatting needs improvement."
            }
        }


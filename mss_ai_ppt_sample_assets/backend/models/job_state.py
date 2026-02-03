"""Job state model for tracking report generation jobs.

This module provides the data model for job status tracking, including
support for idempotency, retry logic, and progress monitoring.
"""

from __future__ import annotations

from enum import Enum
from datetime import datetime
from typing import Optional, Dict, Any, List
from pydantic import BaseModel, Field


class JobStatus(str, Enum):
    """Job execution status."""

    PENDING = "pending"      # Created, waiting to be processed
    RUNNING = "running"      # Currently being processed
    COMPLETED = "completed"  # Successfully completed
    FAILED = "failed"        # Failed with error
    CANCELLED = "cancelled"  # Cancelled by user


class JobState(BaseModel):
    """Complete state of a report generation job.

    This model tracks all aspects of a job's lifecycle, including:
    - Current status and progress
    - Timing information (created, started, completed)
    - Idempotency support
    - Retry tracking
    - Result file paths

    Example:
        job = JobState(
            job_id="session_123:mss_executive_v2",
            session_id="session_123",
            template_id="mss_executive_v2",
            input_id="tenant_acme",
            status=JobStatus.PENDING,
            created_at=datetime.utcnow(),
            updated_at=datetime.utcnow()
        )
    """

    # Identity
    job_id: str = Field(..., description="Unique job identifier (session_id:template_id)")
    session_id: str = Field(..., description="Session identifier for file isolation")
    template_id: str = Field(..., description="Template identifier")
    input_id: str = Field(..., description="Input data identifier")

    # Status
    status: JobStatus = Field(..., description="Current job status")
    progress: int = Field(default=0, ge=0, le=100, description="Progress percentage (0-100)")
    message: str = Field(default="", description="Current progress message")

    # Timing
    created_at: datetime = Field(..., description="When job was created")
    started_at: Optional[datetime] = Field(default=None, description="When job execution started")
    completed_at: Optional[datetime] = Field(default=None, description="When job completed/failed")
    updated_at: datetime = Field(..., description="Last update timestamp")

    # Idempotency support
    idempotency_key: Optional[str] = Field(
        default=None,
        description="Idempotency key to prevent duplicate processing"
    )

    # Retry control
    retry_count: int = Field(default=0, ge=0, description="Number of retry attempts")
    max_retries: int = Field(default=3, ge=0, description="Maximum retry attempts")
    last_error: Optional[str] = Field(default=None, description="Last error message")

    # Results
    report_path: Optional[str] = Field(default=None, description="Path to generated report file")
    slidespec_path: Optional[str] = Field(default=None, description="Path to slidespec JSON file")
    preview_urls: Optional[List[str]] = Field(default=None, description="Preview image URLs")

    # Metadata
    metadata: Dict[str, Any] = Field(
        default_factory=dict,
        description="Additional metadata (warnings, statistics, etc.)"
    )

    class Config:
        json_schema_extra = {
            "example": {
                "job_id": "abc123_20260202110000:mss_executive_v2",
                "session_id": "abc123_20260202110000",
                "template_id": "mss_executive_v2",
                "input_id": "tenant_acme_2025-11",
                "status": "completed",
                "progress": 100,
                "message": "完成",
                "created_at": "2026-02-02T11:00:00Z",
                "started_at": "2026-02-02T11:00:01Z",
                "completed_at": "2026-02-02T11:00:15Z",
                "updated_at": "2026-02-02T11:00:15Z",
                "idempotency_key": "user123-request456",
                "retry_count": 0,
                "max_retries": 3,
                "report_path": "outputs/sessions/abc123/report_mss_executive_v2.pptx",
                "slidespec_path": "outputs/sessions/abc123/slidespec_mss_executive_v2.json",
                "metadata": {"warnings": []}
            }
        }

    def is_terminal(self) -> bool:
        """Check if job is in a terminal state (completed, failed, or cancelled)."""
        return self.status in [JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELLED]

    def is_active(self) -> bool:
        """Check if job is actively running."""
        return self.status == JobStatus.RUNNING

    def can_retry(self) -> bool:
        """Check if job can be retried."""
        return self.status == JobStatus.FAILED and self.retry_count < self.max_retries

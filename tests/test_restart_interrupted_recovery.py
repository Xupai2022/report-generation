from datetime import datetime, timezone
from typing import Optional

from mss_ai_ppt_sample_assets.backend.models.job_state import JobState, JobStatus
from mss_ai_ppt_sample_assets.backend.modules.job_store import JobStore
from mss_ai_ppt_sample_assets.backend.services.job_manager import (
    JobManager,
    RESTART_INTERRUPTED_ERROR_CODE,
    RESTART_INTERRUPTED_ERROR_MESSAGE,
)


class _DummySessionManager:
    def __init__(self):
        self._counter = 0

    def generate_session_id(self) -> str:
        self._counter += 1
        return f"new_session_{self._counter}"


class _DummyReportService:
    def __init__(self):
        self.session_manager = _DummySessionManager()


def _make_job(
    session_id: str,
    template_id: str,
    input_id: str,
    status: JobStatus,
    idempotency_key: Optional[str] = None,
    last_error: Optional[str] = None,
    error_code: Optional[str] = None,
) -> JobState:
    now = datetime.now(timezone.utc)
    return JobState(
        job_id=f"{session_id}:{template_id}",
        session_id=session_id,
        template_id=template_id,
        input_id=input_id,
        status=status,
        idempotency_key=idempotency_key,
        created_at=now,
        updated_at=now,
        last_error=last_error,
        error_code=error_code,
    )


def test_mark_running_jobs_failed_on_recovery(tmp_path):
    store = JobStore(tmp_path / "jobs")
    running = _make_job("s1", "tpl", "input1", JobStatus.RUNNING)
    completed = _make_job("s2", "tpl", "input2", JobStatus.COMPLETED)

    store.create_job(running)
    store.create_job(completed)

    count = store.mark_running_jobs_failed(RESTART_INTERRUPTED_ERROR_MESSAGE)

    assert count == 1
    recovered_running = store.get_job(running.job_id)
    assert recovered_running is not None
    assert recovered_running.status == JobStatus.FAILED
    assert recovered_running.error_code == RESTART_INTERRUPTED_ERROR_CODE
    assert recovered_running.last_error == RESTART_INTERRUPTED_ERROR_MESSAGE
    assert recovered_running.completed_at is not None

    unchanged_completed = store.get_job(completed.job_id)
    assert unchanged_completed is not None
    assert unchanged_completed.status == JobStatus.COMPLETED


def test_create_job_recreates_when_idempotent_job_failed_by_restart(tmp_path):
    store = JobStore(tmp_path / "jobs")
    manager = JobManager(store, _DummyReportService())
    existing = _make_job(
        "old_session",
        "tpl",
        "input1",
        JobStatus.FAILED,
        idempotency_key="idem-key",
        last_error=RESTART_INTERRUPTED_ERROR_MESSAGE,
        error_code=RESTART_INTERRUPTED_ERROR_CODE,
    )
    store.create_job(existing)

    recreated = manager.create_job(
        input_id="input1",
        template_id="tpl",
        idempotency_key="idem-key",
        session_id="old_session",
    )

    assert recreated.job_id != existing.job_id
    assert recreated.status == JobStatus.PENDING
    assert recreated.session_id.startswith("new_session_")

    mapped = store.find_by_idempotency_key("idem-key")
    assert mapped is not None
    assert mapped.job_id == recreated.job_id

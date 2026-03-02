from types import SimpleNamespace
from pathlib import Path
import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.schemas.requests import CreateReportRequest


def _build_template_stub():
    ai_placeholder = SimpleNamespace(
        token="AI_TEXT",
        ai_generate=True,
        ai_instruction="生成摘要内容",
        max_length=None,
        max_items=None,
        max_chars_per_item=None,
    )
    slide = SimpleNamespace(
        slide_key="summary",
        title="Summary",
        placeholders=[ai_placeholder],
    )
    return SimpleNamespace(
        slides=[slide],
        get_ai_placeholders=lambda: [("summary", "AI_TEXT", ai_placeholder)],
    )


def _build_orchestrator():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    return orchestrator


def test_build_user_prompt_without_focus_options():
    orchestrator = _build_orchestrator()
    template = _build_template_stub()
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=None,
    )

    assert "报告重点偏向（用户选择）" not in prompt


def test_build_user_prompt_with_vulnerability_focus():
    orchestrator = _build_orchestrator()
    template = _build_template_stub()
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["vulnerability"],
    )

    assert "报告重点偏向（用户选择）" in prompt
    assert "报告重点偏向漏洞" in prompt


def test_build_user_prompt_with_alert_focus():
    orchestrator = _build_orchestrator()
    template = _build_template_stub()
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["alert"],
    )

    assert "报告重点偏向（用户选择）" in prompt
    assert "报告重点偏向告警" in prompt


def test_build_user_prompt_with_both_focus_options():
    orchestrator = _build_orchestrator()
    template = _build_template_stub()
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["vulnerability", "alert"],
    )

    assert "报告重点偏向漏洞" in prompt
    assert "报告重点偏向告警" in prompt
    assert prompt.index("报告重点偏向漏洞") < prompt.index("报告重点偏向告警")


def test_build_user_prompt_for_slides_includes_focus_section():
    orchestrator = _build_orchestrator()
    template = _build_template_stub()
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt_for_slides(
        tenant_input=tenant_input,
        template=template,
        slide_keys=["summary"],
        batch_index=0,
        total_batches=1,
        focus_options=["vulnerability", "alert"],
    )

    assert "报告重点偏向（用户选择）" in prompt
    assert "报告重点偏向漏洞" in prompt
    assert "报告重点偏向告警" in prompt


def test_create_report_request_focus_options_optional():
    req = CreateReportRequest(
        input_id="tenant_acme_2025-11",
        template_id="mss_executive_v2",
        use_mock=False,
    )
    assert req.focus_options is None


def test_create_report_request_focus_options_empty_list_allowed():
    req = CreateReportRequest(
        input_id="tenant_acme_2025-11",
        template_id="mss_executive_v2",
        use_mock=False,
        focus_options=[],
    )
    assert req.focus_options == []


def test_create_report_request_focus_options_invalid_value_rejected():
    with pytest.raises(ValidationError):
        CreateReportRequest(
            input_id="tenant_acme_2025-11",
            template_id="mss_executive_v2",
            use_mock=False,
            focus_options=["invalid_option"],
        )

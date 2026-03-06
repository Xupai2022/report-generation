from pathlib import Path
import sys
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.schemas.requests import CreateReportRequest


def _build_template_stub(ai_instruction: str):
    ai_placeholder = SimpleNamespace(
        token="AI_TEXT",
        ai_generate=True,
        ai_instruction=ai_instruction,
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
    return LLMOrchestratorV2.__new__(LLMOrchestratorV2)


def test_build_user_prompt_replaces_preference_placeholder_and_includes_selected_annotation():
    orchestrator = _build_orchestrator()
    template = _build_template_stub(
        "根据已选{preference}从注解中寻找基础知识，重点突出与用户偏好的关联性。"
    )
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["业务保护"],
    )

    assert "## 参考注解（按用户已选偏好唯一匹配）" in prompt
    assert "- 业务保护：" in prompt
    assert "{preference}" not in prompt
    assert "根据已选业务保护从注解中寻找基础知识" in prompt


def test_build_user_prompt_supports_internal_focus_key_and_legacy_user_preference_marker():
    orchestrator = _build_orchestrator()
    template = _build_template_stub("结合**用户偏好**，生成本页结论。")
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["business_protection"],
    )

    assert "结合**业务保护**，生成本页结论。" in prompt
    assert "用户偏好" not in prompt


def test_build_user_prompt_rejects_unsupported_focus_option():
    orchestrator = _build_orchestrator()
    template = _build_template_stub("test")
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    with pytest.raises(ValueError):
        orchestrator._build_user_prompt(
            tenant_input=tenant_input,
            template=template,
            focus_options=["not-exists"],
        )


def test_create_report_request_normalizes_focus_options_to_titles():
    req = CreateReportRequest(
        input_id="tenant_acme_2025-11",
        template_id="mss_executive_v2",
        use_mock=False,
        focus_options=["business_protection", "业务保护", "alert"],
    )

    assert req.focus_options == ["业务保护", "告警优先"]


def test_create_report_request_rejects_invalid_focus_option():
    with pytest.raises(ValidationError):
        CreateReportRequest(
            input_id="tenant_acme_2025-11",
            template_id="mss_executive_v2",
            use_mock=False,
            focus_options=["invalid_option"],
        )

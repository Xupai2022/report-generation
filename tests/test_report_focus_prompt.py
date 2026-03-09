from pathlib import Path
import sys
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.schemas.requests import CreateReportRequest


ZH_FOCUS_SECTION = "## \u504f\u597d\u91cd\u70b9\u6307\u5f15"
ZH_BUSINESS_PROTECTION = "\u4e1a\u52a1\u4fdd\u62a4"
ZH_USER_PREFERENCE = "\u7528\u6237\u504f\u597d"


def _build_template_stub(ai_instruction: str):
    ai_placeholder = SimpleNamespace(
        token="AI_TEXT",
        ai_generate=True,
        ai_instruction=ai_instruction,
        max_length=None,
        max_items=None,
        max_chars_per_item=None,
        source=None,
    )
    slide = SimpleNamespace(
        slide_key="summary",
        title="Summary",
        context_policy="local_only",
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
        "\u6839\u636e\u5df2\u9009{preference}\u4ece\u6ce8\u89e3\u4e2d\u5bfb\u627e\u57fa\u7840\u77e5\u8bc6\uff0c\u91cd\u70b9\u7a81\u51fa\u4e0e\u7528\u6237\u504f\u597d\u7684\u5173\u8054\u6027\u3002"
    )
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["business_protection"],
    )

    assert ZH_FOCUS_SECTION in prompt
    assert f"- {ZH_BUSINESS_PROTECTION}:" in prompt
    assert "{preference}" not in prompt
    assert f"\u6839\u636e\u5df2\u9009{ZH_BUSINESS_PROTECTION}\u4ece\u6ce8\u89e3\u4e2d\u5bfb\u627e\u57fa\u7840\u77e5\u8bc6" in prompt


def test_build_user_prompt_supports_internal_focus_key_and_legacy_user_preference_marker():
    orchestrator = _build_orchestrator()
    template = _build_template_stub(f"\u7ed3\u5408**{ZH_USER_PREFERENCE}**\uff0c\u751f\u6210\u672c\u9875\u7ed3\u8bba\u3002")
    tenant_input = TenantInput(raw={"period": {"start": "2026-01-01", "end": "2026-01-31"}})

    prompt = orchestrator._build_user_prompt(
        tenant_input=tenant_input,
        template=template,
        focus_options=["business_protection"],
    )

    assert f"\u7ed3\u5408**{ZH_BUSINESS_PROTECTION}**\uff0c\u751f\u6210\u672c\u9875\u7ed3\u8bba\u3002" in prompt
    assert ZH_USER_PREFERENCE not in prompt


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


def test_create_report_request_normalizes_focus_options_to_internal_keys():
    req = CreateReportRequest(
        input_id="tenant_acme_2025-11",
        template_id="mss_executive_v2",
        use_mock=False,
        focus_options=["business_protection", "business protection", "alert"],
    )

    assert req.focus_options == ["business_protection", "alert"]


def test_create_report_request_rejects_invalid_focus_option():
    with pytest.raises(ValidationError):
        CreateReportRequest(
            input_id="tenant_acme_2025-11",
            template_id="mss_executive_v2",
            use_mock=False,
            focus_options=["invalid_option"],
        )

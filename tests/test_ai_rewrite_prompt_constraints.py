from types import SimpleNamespace
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
import pytest


def test_rewrite_prompt_includes_structured_data_and_dynamic_target_tokens():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="incident_effectiveness",
                placeholders=[
                    SimpleNamespace(
                        token="incident_total",
                        ai_generate=False,
                        source="incident_effectiveness.incident_total",
                    ),
                    SimpleNamespace(
                        token="trust_assurance",
                        ai_generate=True,
                        source=None,
                    ),
                    SimpleNamespace(
                        token="security_trust",
                        ai_generate=True,
                        source=None,
                    ),
                ],
            )
        ]
    )

    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)
    orchestrator._build_system_prompt = lambda _template: "system"
    orchestrator._build_user_prompt_for_slides = lambda **_kwargs: "BASE_PROMPT"

    captured = {}

    def fake_call_and_parse_with_retry(system_prompt, user_prompt, _template):
        captured["system_prompt"] = system_prompt
        captured["user_prompt"] = user_prompt
        return {
            "incident_effectiveness": {
                "trust_assurance": "new trust",
                "security_trust": "new security",
                "incident_total": "should be filtered",
            }
        }

    orchestrator._call_and_parse_with_retry = fake_call_and_parse_with_retry

    tenant_input = TenantInput(
        raw={
            "incident_effectiveness": {
                "incident_total": "3366666666",
            }
        }
    )

    result = orchestrator.rewrite_single_slide_v2(
        tenant_input=tenant_input,
        template_id="mss_classic_ops",
        slide_key="incident_effectiveness",
        user_prompt="根据最新数据重写",
        current_slide_content={
            "incident_total": "3366666666",
            "trust_assurance": "old text with 33",
            "security_trust": "old text with 39",
        },
    )

    assert captured["system_prompt"] == "system"
    assert "## Current Slide Structured Data (Highest Priority)" in captured["user_prompt"]
    assert "3366666666" in captured["user_prompt"]
    assert "## Previous AI Copy (Style Reference Only)" in captured["user_prompt"]
    assert "old text with 33" in captured["user_prompt"]
    assert "Output only these dynamic target placeholders: trust_assurance, security_trust." in captured["user_prompt"]

    assert set(result["placeholders"].keys()) == {"trust_assurance", "security_trust"}
    assert "incident_total" not in result["placeholders"]


def test_rewrite_single_slide_honors_target_tokens_subset():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="incident_effectiveness",
                placeholders=[
                    SimpleNamespace(token="incident_total", ai_generate=False, source="incident_effectiveness.incident_total"),
                    SimpleNamespace(token="trust_assurance", ai_generate=True, source=None),
                    SimpleNamespace(token="security_trust", ai_generate=True, source=None),
                ],
            )
        ]
    )

    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)
    orchestrator._build_system_prompt = lambda _template: "system"
    orchestrator._build_user_prompt_for_slides = lambda **_kwargs: "BASE_PROMPT"

    captured = {}

    def fake_call_and_parse_with_retry(system_prompt, user_prompt, _template):
        captured["user_prompt"] = user_prompt
        return {
            "incident_effectiveness": {
                "trust_assurance": "new trust",
                "security_trust": "new security",
            }
        }

    orchestrator._call_and_parse_with_retry = fake_call_and_parse_with_retry

    tenant_input = TenantInput(raw={"incident_effectiveness": {"incident_total": 10}})
    result = orchestrator.rewrite_single_slide_v2(
        tenant_input=tenant_input,
        template_id="mss_classic_ops",
        slide_key="incident_effectiveness",
        user_prompt="只改第一个token",
        current_slide_content={
            "incident_total": 10,
            "trust_assurance": "old trust",
            "security_trust": "old security",
        },
        target_tokens=["trust_assurance"],
    )

    assert "Dynamic target placeholders: trust_assurance" in captured["user_prompt"]
    assert set(result["placeholders"].keys()) == {"trust_assurance"}
    assert result["updated_tokens"] == ["trust_assurance"]


def test_rewrite_single_slide_empty_target_tokens_falls_back_to_all():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="summary",
                placeholders=[
                    SimpleNamespace(token="metric", ai_generate=False, source="coverage.metric"),
                    SimpleNamespace(token="AI_A", ai_generate=True, source=None),
                    SimpleNamespace(token="AI_B", ai_generate=True, source=None),
                ],
            )
        ]
    )

    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)
    orchestrator._build_system_prompt = lambda _template: "system"
    orchestrator._build_user_prompt_for_slides = lambda **_kwargs: "BASE_PROMPT"
    orchestrator._call_and_parse_with_retry = lambda *_args, **_kwargs: {
        "summary": {
            "AI_A": "a",
            "AI_B": "b",
        }
    }

    result = orchestrator.rewrite_single_slide_v2(
        tenant_input=TenantInput(raw={"coverage": {"metric": 1}}),
        template_id="mss_classic_ops",
        slide_key="summary",
        user_prompt="rewrite",
        current_slide_content={"metric": 1},
        target_tokens=[],
    )

    assert set(result["placeholders"].keys()) == {"AI_A", "AI_B"}


def test_rewrite_single_slide_rejects_invalid_target_token():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="summary",
                placeholders=[
                    SimpleNamespace(token="metric", ai_generate=False, source="coverage.metric"),
                    SimpleNamespace(token="AI_A", ai_generate=True, source=None),
                ],
            )
        ]
    )
    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)

    with pytest.raises(ValueError, match="Invalid target_tokens"):
        orchestrator.rewrite_single_slide_v2(
            tenant_input=TenantInput(raw={"coverage": {"metric": 1}}),
            template_id="mss_classic_ops",
            slide_key="summary",
            user_prompt="rewrite",
            target_tokens=["NOT_EXISTS"],
        )

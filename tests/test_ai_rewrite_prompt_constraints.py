from types import SimpleNamespace
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
import pytest


ZH_REWRITE_TASK = "## \u6539\u5199\u4efb\u52a1"
ZH_STRUCTURED_DATA = "## \u5f53\u524d\u9875\u9762\u7ed3\u6784\u5316\u6570\u636e\uff08\u6700\u9ad8\u4f18\u5148\u7ea7\uff09"
ZH_RELEVANT_CONTEXT = "## \u5f53\u524d\u9875\u9762\u76f8\u5173\u6570\u636e\uff08\u4ec5\u4e0a\u4e0b\u6587\uff09"
ZH_HISTORY_COPY = "## \u5386\u53f2 AI \u6587\u6848\uff08\u4ec5\u98ce\u683c\u53c2\u8003\uff09"
ZH_HARD_CONSTRAINT = "## \u786c\u7ea6\u675f"
ZH_OUTPUT_FORMAT = "## \u8f93\u51fa\u683c\u5f0f"


def _ph(token: str, ai_generate: bool, source: str = None):
    return SimpleNamespace(
        token=token,
        ai_generate=ai_generate,
        source=source,
        ai_instruction=None,
    )


def test_rewrite_prompt_includes_structured_data_and_dynamic_target_tokens():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="incident_effectiveness",
                context_policy="local_only",
                placeholders=[
                    _ph("incident_total", False, "incident_effectiveness.incident_total"),
                    _ph("trust_assurance", True),
                    _ph("security_trust", True),
                ],
            )
        ]
    )

    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)
    orchestrator._build_system_prompt = lambda _template: "system"

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

    tenant_input = TenantInput(raw={"incident_effectiveness": {"incident_total": "3366666666"}})

    result = orchestrator.rewrite_single_slide_v2(
        tenant_input=tenant_input,
        template_id="mss_classic_ops",
        slide_key="incident_effectiveness",
        user_prompt="\u6839\u636e\u6700\u65b0\u6570\u636e\u91cd\u5199",
        current_slide_content={
            "incident_total": "3366666666",
            "trust_assurance": "old text with 33",
            "security_trust": "old text with 39",
        },
    )

    assert captured["system_prompt"] == "system"
    assert ZH_STRUCTURED_DATA in captured["user_prompt"]
    assert "3366666666" in captured["user_prompt"]
    assert ZH_HISTORY_COPY in captured["user_prompt"]
    assert "old text with 33" in captured["user_prompt"]
    assert "### \u9875\u9762\uff1a" not in captured["user_prompt"]
    assert "## \u6570\u636e\u4f18\u5148\u7ea7" in captured["user_prompt"]
    assert (
        "\u53ea\u8f93\u51fa\u4ee5\u4e0b\u76ee\u6807\u5360\u4f4d\u7b26\uff1atrust_assurance, security_trust\u3002"
        in captured["user_prompt"]
    )
    assert captured["user_prompt"].index(ZH_REWRITE_TASK) < captured["user_prompt"].index(ZH_STRUCTURED_DATA)
    assert captured["user_prompt"].index(ZH_RELEVANT_CONTEXT) < captured["user_prompt"].index(ZH_HISTORY_COPY)
    assert captured["user_prompt"].index(ZH_HISTORY_COPY) < captured["user_prompt"].index(ZH_HARD_CONSTRAINT)
    assert captured["user_prompt"].index(ZH_HARD_CONSTRAINT) < captured["user_prompt"].index(ZH_OUTPUT_FORMAT)

    assert set(result["placeholders"].keys()) == {"trust_assurance", "security_trust"}
    assert "incident_total" not in result["placeholders"]


def test_rewrite_single_slide_honors_target_tokens_subset():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="incident_effectiveness",
                context_policy="local_only",
                placeholders=[
                    _ph("incident_total", False, "incident_effectiveness.incident_total"),
                    _ph("trust_assurance", True),
                    _ph("security_trust", True),
                ],
            )
        ]
    )

    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)
    orchestrator._build_system_prompt = lambda _template: "system"

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
        user_prompt="\u53ea\u6539\u7b2c\u4e00\u4e2a token",
        current_slide_content={
            "incident_total": 10,
            "trust_assurance": "old trust",
            "security_trust": "old security",
        },
        target_tokens=["trust_assurance"],
    )

    assert "\u672c\u6b21\u76ee\u6807\u5360\u4f4d\u7b26\uff1atrust_assurance" in captured["user_prompt"]
    assert "old security" not in captured["user_prompt"]
    assert set(result["placeholders"].keys()) == {"trust_assurance"}
    assert result["updated_tokens"] == ["trust_assurance"]


def test_rewrite_single_slide_empty_target_tokens_falls_back_to_all():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            SimpleNamespace(
                slide_key="summary",
                context_policy="local_only",
                placeholders=[
                    _ph("metric", False, "coverage.metric"),
                    _ph("AI_A", True),
                    _ph("AI_B", True),
                ],
            )
        ]
    )

    orchestrator.template_repo = SimpleNamespace(get_descriptor_v2=lambda _template_id: template)
    orchestrator._build_system_prompt = lambda _template: "system"
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
                context_policy="local_only",
                placeholders=[
                    _ph("metric", False, "coverage.metric"),
                    _ph("AI_A", True),
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

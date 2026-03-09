from pathlib import Path
import sys
from types import SimpleNamespace
from typing import Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2


def _placeholder(token: str, source: Optional[str], ai_generate: bool, ai_instruction: Optional[str] = None):
    return SimpleNamespace(
        token=token,
        source=source,
        ai_generate=ai_generate,
        ai_instruction=ai_instruction,
        max_length=None,
        max_items=None,
        max_chars_per_item=None,
    )


def _slide(slide_no: int, key: str, policy: str, placeholders):
    return SimpleNamespace(
        slide_no=slide_no,
        slide_key=key,
        title=key,
        context_policy=policy,
        placeholders=placeholders,
    )


def test_build_context_payload_for_slides_local_only_excludes_irrelevant_roots():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            _slide(11, "slide_a", "local_only", [_placeholder("A", "alpha.value", True)]),
            _slide(12, "slide_b", "local_only", [_placeholder("B", "beta.value", True)]),
        ]
    )
    tenant_input = TenantInput(raw={"alpha": {"value": 1}, "beta": {"value": 2}, "gamma": {"value": 3}})

    payload = orchestrator._build_context_payload_for_slides(tenant_input, template, ["slide_a"])

    assert payload == {"alpha": {"value": 1}}


def test_build_context_payload_for_slides_full_data_policy_returns_full_payload():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)

    template = SimpleNamespace(
        slides=[
            _slide(19, "slide_full", "full_data", [_placeholder("X", "alpha.value", True)]),
        ]
    )
    raw = {"alpha": {"value": 1}, "beta": {"value": 2}}
    tenant_input = TenantInput(raw=raw)

    payload = orchestrator._build_context_payload_for_slides(tenant_input, template, ["slide_full"])

    assert payload == raw

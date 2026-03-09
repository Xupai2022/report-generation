from pathlib import Path
import sys
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2


def _ai_slide(slide_no: int, slide_key: str, policy: str = "local_only"):
    placeholder = SimpleNamespace(
        token=f"TOK_{slide_key}",
        ai_generate=True,
        ai_instruction="test",
        source=f"{slide_key}.value",
        max_length=None,
        max_items=None,
        max_chars_per_item=None,
    )
    return SimpleNamespace(
        slide_no=slide_no,
        slide_key=slide_key,
        title=slide_key,
        context_policy=policy,
        placeholders=[placeholder],
    )


def test_get_smart_slide_batches_supports_non_contiguous_local_grouping_and_full_data_single_batch():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    template = SimpleNamespace(
        slides=[
            _ai_slide(11, "s11", "local_only"),
            _ai_slide(13, "s13", "local_only"),
            _ai_slide(14, "s14", "local_only"),
            _ai_slide(15, "s15", "local_only"),
            _ai_slide(19, "s19", "full_data"),
        ]
    )

    # hard_cap = floor(2000 * 0.70) = 1400
    def fake_estimate_batch_prompt_tokens(*, slide_keys, **_kwargs):
        key = frozenset(slide_keys)
        if key == frozenset({"s11", "s14"}):
            return 850
        if key == frozenset({"s13", "s15"}):
            return 860
        if len(key) == 1:
            return 1000
        if len(key) == 2:
            return 1300
        return 2000

    orchestrator._estimate_batch_prompt_tokens = fake_estimate_batch_prompt_tokens

    batches = orchestrator._get_smart_slide_batches(
        tenant_input=TenantInput(raw={}),
        template=template,
        max_tokens_per_batch=2000,
        focus_options=["business_protection"],
    )

    normalized = [set(batch) for batch in batches]
    assert {"s11", "s14"} in normalized
    assert {"s13", "s15"} in normalized
    assert {"s19"} in normalized

    hard_cap = int(2000 * 0.70)
    for batch in batches:
        if batch == ["s19"]:
            continue
        assert fake_estimate_batch_prompt_tokens(slide_keys=batch) <= hard_cap

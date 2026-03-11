from __future__ import annotations

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.models.templates import TemplateDescriptorV2
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.services.rag_service import (
    GenerationRetrievalTask,
    QueryBuilderV2,
    RAGService,
    RetrievalResult,
)


def _template_with_ai() -> TemplateDescriptorV2:
    return TemplateDescriptorV2.model_validate(
        {
            "template_id": "mss_test_query_builder",
            "name": "RAG Query Builder Test",
            "version": "1.0.0",
            "pptx_file": "demo.pptx",
            "audience": "management",
            "slides": [
                {
                    "slide_no": 1,
                    "slide_key": "threat_effectiveness",
                    "title": "Threat",
                    "context_policy": "local_only",
                    "placeholders": [
                        {"token": "threat_total", "source": "threat_effectiveness.total"},
                        {
                            "token": "threat_summary",
                            "ai_generate": True,
                            "ai_instruction": "Summarize threat detection and response effectiveness.",
                        },
                    ],
                },
                {
                    "slide_no": 2,
                    "slide_key": "asset_management",
                    "title": "Asset",
                    "context_policy": "local_only",
                    "placeholders": [
                        {"token": "asset_total", "source": "asset_management.total"},
                        {
                            "token": "asset_summary",
                            "ai_generate": True,
                            "ai_instruction": "Summarize asset governance priorities.",
                        },
                    ],
                },
            ],
        }
    )


def test_query_builder_generation_payload_contains_task_signal():
    task = GenerationRetrievalTask(
        slide_key="threat_effectiveness",
        slide_title="Threat",
        ai_tokens=["threat_summary"],
        ai_instructions=["Summarize the business impact from the threat metrics."],
        facts={
            "threat_effectiveness.total": 12,
            "threat_effectiveness.levels": ["high", "mid"],
        },
        focus_options=["alert", "business_protection"],
        context_policy="local_only",
    )

    query = QueryBuilderV2.build_generation_query("mss_classic_ops", task)

    assert "scene: generate" in query
    assert "template: mss_classic_ops" in query
    assert "slide_key: threat_effectiveness" in query
    assert "ai_tokens: threat_summary" in query
    assert "focus: alert, business_protection" in query
    assert "facts:" in query
    assert "total=12" in query
    assert "levels=[list:2]" in query
    assert "{\"" not in query


def test_rag_service_generation_uses_slide_tasks():
    template = _template_with_ai()
    tenant_input = TenantInput(
        raw={
            "threat_effectiveness": {"total": 22, "levels": ["high", "mid", "low"]},
            "asset_management": {"total": 125},
        }
    )
    service = RAGService()
    captured = []

    def fake_retrieve_task_result(**kwargs):
        captured.append(kwargs)
        slide_key = kwargs["slide_key"]
        return RetrievalResult(
            rag_used=True,
            context="context",
            retrieval_trace=[
                {
                    "chunk_id": f"{slide_key}-chunk",
                    "score": 0.88,
                    "source_file": "kb.md",
                    "page_no": None,
                    "slide_key": slide_key,
                }
            ],
            hits=[
                {
                    "chunk_id": f"{slide_key}-chunk",
                    "score": 0.88,
                    "dense_score": 0.88,
                    "rerank_score": 1.2,
                    "source_file": "kb.md",
                    "page_no": None,
                    "text": f"evidence for {slide_key}",
                    "slide_key": slide_key,
                    "section_scope": "slide",
                    "section_title": "section",
                    "heading_path": "section",
                    "chunk_order": 0,
                }
            ],
            retrieval_stats={"kept_hits": 1, "local_hits": 1, "common_fallback_hits": 0},
        )

    service._retrieve_task_result = fake_retrieve_task_result

    result = service.retrieve_for_generation(
        tenant_input=tenant_input,
        input_id="tenant_demo",
        template_id="mss_classic_ops",
        focus_options=["alert"],
        use_rag=True,
        template_descriptor=template,
    )

    assert len(captured) == 2
    assert result.rag_used is True
    assert set(result.context_by_slide.keys()) == {"threat_effectiveness", "asset_management"}
    assert all(item.get("slide_key") for item in result.retrieval_trace)


def test_prompt_uses_slide_level_rag_context():
    template = _template_with_ai()
    tenant_input = TenantInput(
        raw={
            "threat_effectiveness": {"total": 22},
            "asset_management": {"total": 125},
        }
    )
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    prompt = orchestrator._build_user_prompt_for_slides(
        tenant_input=tenant_input,
        template=template,
        slide_keys=["threat_effectiveness", "asset_management"],
        rag_context="GLOBAL CONTEXT",
        rag_context_by_slide={
            "threat_effectiveness": "[1] source=kb1.md\nthreat evidence",
            "asset_management": "[1] source=kb2.md\nasset evidence",
        },
    )

    assert "[slide=threat_effectiveness]" in prompt
    assert "[slide=asset_management]" in prompt
    assert "GLOBAL CONTEXT" not in prompt


def test_batch_rag_budget_trims_long_context():
    template = _template_with_ai()
    tenant_input = TenantInput(
        raw={
            "threat_effectiveness": {"total": 22},
            "asset_management": {"total": 125},
        }
    )
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    by_slide = {
        "threat_effectiveness": "A" * 4000,
        "asset_management": "B" * 4000,
    }

    allocated = orchestrator._build_rag_context_by_slide_for_batch(
        tenant_input=tenant_input,
        template=template,
        slide_keys=["threat_effectiveness", "asset_management"],
        max_tokens_per_batch=800,
        focus_options=None,
        rag_context_by_slide=by_slide,
    )

    assert allocated
    total_chars = sum(len(item) for item in allocated.values())
    assert total_chars <= 800
    assert all(len(allocated[k]) < len(by_slide[k]) for k in allocated)


def test_local_only_common_fallback():
    service = RAGService()
    calls = []

    def fake_search_hits(**kwargs):
        calls.append(kwargs)
        if kwargs.get("slide_keys"):
            return RetrievalResult(
                rag_used=True,
                hits=[
                    {
                        "chunk_id": "local-1",
                        "score": 0.8,
                        "dense_score": 0.8,
                        "rerank_score": None,
                        "source_file": "kb.md",
                        "page_no": None,
                        "text": "local evidence",
                        "slide_key": "incident_effectiveness",
                        "section_scope": "slide",
                        "section_title": "local section",
                        "heading_path": "local section",
                        "chunk_order": 0,
                    }
                ],
                retrieval_stats={"candidate_count": 16, "raw_hits": 1, "dense_kept_hits": 1, "rerank_enabled": False},
            )
        return RetrievalResult(
            rag_used=True,
            hits=[
                {
                    "chunk_id": "common-1",
                    "score": 0.7,
                    "dense_score": 0.7,
                    "rerank_score": None,
                    "source_file": "kb.md",
                    "page_no": None,
                    "text": "common evidence",
                    "slide_key": "__common__",
                    "section_scope": "common",
                    "section_title": "shared section",
                    "heading_path": "shared section",
                    "chunk_order": 0,
                }
            ],
            retrieval_stats={"candidate_count": 16, "raw_hits": 1, "dense_kept_hits": 1, "rerank_enabled": False},
        )

    service._search_hits = fake_search_hits

    result = service._retrieve_task_result(
        query_text="scene: generate\nslide_key: incident_effectiveness",
        template_id="mss_classic_ops",
        slide_key="incident_effectiveness",
        scene="generate",
        context_policy="local_only",
    )

    assert [item["section_scope"] for item in result.hits] == ["slide", "common"]
    assert result.retrieval_stats["local_hits"] == 1
    assert result.retrieval_stats["common_fallback_hits"] == 1
    assert calls[0]["slide_keys"] == ["incident_effectiveness"]
    assert calls[1]["section_scopes"] == ["common"]


def test_rerank_degrades_without_model():
    service = RAGService()
    service._reranker_status = "unavailable"
    service._load_reranker = lambda: None
    hits = [
        {"chunk_id": "a", "text": "A", "dense_score": 0.4, "score": 0.4},
        {"chunk_id": "b", "text": "B", "dense_score": 0.8, "score": 0.8},
    ]

    ranked, stats = service._rerank_hits(scene="generate", normalized_query="query", hits=hits, final_top_k=2)

    assert [item["chunk_id"] for item in ranked] == ["b", "a"]
    assert stats["rerank_enabled"] is False
    assert stats["rerank_degraded_reason"] == "unavailable"


def test_rerank_reorders_hits():
    class DummyReranker:
        def predict(self, pairs):
            assert len(pairs) == 2
            return [0.1, 0.9]

    service = RAGService()
    service._load_reranker = lambda: DummyReranker()
    hits = [
        {"chunk_id": "a", "text": "A", "dense_score": 0.8, "score": 0.8},
        {"chunk_id": "b", "text": "B", "dense_score": 0.5, "score": 0.5},
    ]

    ranked, stats = service._rerank_hits(scene="generate", normalized_query="query", hits=hits, final_top_k=2)

    assert [item["chunk_id"] for item in ranked] == ["b", "a"]
    assert stats["rerank_enabled"] is True
    assert ranked[0]["rerank_score"] == 0.9


def test_rerank_can_be_disabled_for_scene(monkeypatch):
    service = RAGService()
    monkeypatch.setattr(config.settings, "rag_rerank_scenes", {"rewrite"})
    hits = [
        {"chunk_id": "a", "text": "A", "dense_score": 0.4, "score": 0.4},
        {"chunk_id": "b", "text": "B", "dense_score": 0.8, "score": 0.8},
    ]

    ranked, stats = service._rerank_hits(scene="generate", normalized_query="query", hits=hits, final_top_k=2)

    assert [item["chunk_id"] for item in ranked] == ["b", "a"]
    assert stats["rerank_enabled"] is False
    assert stats["rerank_degraded_reason"] == "disabled_for_scene:generate"


def test_local_reranker_is_complete_requires_weight_file(tmp_path):
    service = RAGService()
    model_dir = tmp_path / "bge-reranker-base"
    model_dir.mkdir()
    (model_dir / "config.json").write_text("{}", encoding="utf-8")
    (model_dir / "tokenizer.json").write_text("{}", encoding="utf-8")

    assert service._local_reranker_is_complete(str(model_dir)) is False

    (model_dir / "model.safetensors").write_text("stub", encoding="utf-8")

    assert service._local_reranker_is_complete(str(model_dir)) is True


def test_build_context_formats_header_and_skips_common_with_enough_local_hits():
    service = RAGService()
    hits = [
        {
            "chunk_id": "local-1",
            "text": "local one",
            "source_file": "kb.md",
            "slide_key": "incident_effectiveness",
            "section_scope": "slide",
            "section_title": "5.1 Framework",
            "heading_path": "5.1 Framework",
            "chunk_order": 0,
            "dense_score": 0.71,
            "rerank_score": 1.91,
            "page_no": None,
        },
        {
            "chunk_id": "local-2",
            "text": "local two",
            "source_file": "kb.md",
            "slide_key": "incident_effectiveness",
            "section_scope": "slide",
            "section_title": "5.2 Signals",
            "heading_path": "5.2 Signals",
            "chunk_order": 0,
            "dense_score": 0.69,
            "rerank_score": 1.87,
            "page_no": None,
        },
        {
            "chunk_id": "common-1",
            "text": "shared text",
            "source_file": "kb.md",
            "slide_key": "__common__",
            "section_scope": "common",
            "section_title": "Shared",
            "heading_path": "Shared",
            "chunk_order": 0,
            "dense_score": 0.6,
            "rerank_score": 1.0,
            "page_no": None,
        },
    ]

    context = service._build_context(hits, max_chars=1000, context_policy="local_only")

    assert "slide=incident_effectiveness" in context
    assert "section=5.1 Framework" in context
    assert "dense=0.7100" in context
    assert "rerank=1.9100" in context
    assert "scope=common" not in context

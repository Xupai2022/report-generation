from __future__ import annotations

import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

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
                            "ai_instruction": "基于threat_effectiveness数据总结威胁检测和处置成效。",
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
                            "ai_instruction": "基于asset_management数据总结资产治理重点。",
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
        ai_instructions=["请给出基于数据的业务影响总结"],
        facts={"threat_effectiveness.total": 12, "threat_effectiveness.levels": ["high", "mid"]},
        focus_options=["alert", "business_protection"],
        context_policy="local_only",
    )

    query = QueryBuilderV2.build_generation_query("mss_classic_ops", task)
    payload = json.loads(query)
    assert payload["scene"] == "generate"
    assert payload["template"] == "mss_classic_ops"
    assert payload["slide_key"] == "threat_effectiveness"
    assert payload["ai_tokens"] == ["threat_summary"]
    assert payload["focus"] == "alert,business_protection"
    assert "facts" in payload


def test_rag_service_generation_uses_slide_tasks():
    template = _template_with_ai()
    tenant_input = TenantInput(
        raw={
            "threat_effectiveness": {"total": 22, "levels": ["high", "mid", "low"]},
            "asset_management": {"total": 125},
        }
    )
    service = RAGService()
    captured_queries = []

    def fake_query(query_text, template_id=None, scene="diagnostic", top_k=None, min_score=None):
        captured_queries.append((scene, query_text, template_id))
        return RetrievalResult(
            rag_used=True,
            retrieval_trace=[{"chunk_id": "c1", "score": 0.88, "source_file": "kb.md", "page_no": 1}],
            hits=[{"chunk_id": "c1", "score": 0.88, "source_file": "kb.md", "page_no": 1, "text": "evidence"}],
            retrieval_stats={"kept_hits": 1},
        )

    service.query = fake_query

    result = service.retrieve_for_generation(
        tenant_input=tenant_input,
        input_id="tenant_demo",
        template_id="mss_classic_ops",
        focus_options=["alert"],
        use_rag=True,
        template_descriptor=template,
    )

    assert len(captured_queries) == 2
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

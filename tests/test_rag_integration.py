from pathlib import Path
import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.models.templates import TemplateDescriptorV2
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.schemas.requests import CreateReportRequest
from mss_ai_ppt_sample_assets.backend.services.rag_service import MarkdownSection, RAGService


def _minimal_template() -> TemplateDescriptorV2:
    return TemplateDescriptorV2.model_validate(
        {
            "template_id": "mss_test_rag",
            "name": "RAG Test Template",
            "version": "1.0.0",
            "pptx_file": "demo.pptx",
            "audience": "management",
            "language": "zh-CN",
            "slides": [
                {
                    "slide_no": 1,
                    "slide_key": "summary",
                    "title": "Summary",
                    "placeholders": [
                        {"token": "metric", "source": "metrics.total", "ai_generate": False},
                        {
                            "token": "insight",
                            "ai_generate": True,
                            "ai_instruction": "Generate an insight from the metrics.",
                            "max_length": 100,
                        },
                    ],
                }
            ],
        }
    )


def test_prompt_contains_rag_context_section():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    tenant_input = TenantInput(raw={"metrics": {"total": 10}})
    template = _minimal_template()

    prompt = orchestrator._build_user_prompt_for_slides(
        tenant_input=tenant_input,
        template=template,
        slide_keys=["summary"],
        rag_context="[1] source=kb.md, section=overview, dense=0.91\nkey evidence",
    )

    assert "检索证据（辅助上下文）" in prompt
    assert "key evidence" in prompt


def test_rewrite_base_prompt_contains_rag_context():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    prompt = orchestrator._build_rewrite_base_prompt(
        context_payload={"metric": 10},
        use_full_data=False,
        rag_context="source=kb.md\napply layered controls",
    )
    assert "检索证据（辅助上下文）" in prompt
    assert "apply layered controls" in prompt


def test_create_report_request_accepts_default_use_rag():
    req = CreateReportRequest(
        input_id="tenant_demo",
        template_id="mss_test_rag",
        focus_options=["alert"],
    )
    assert req.use_rag is True
    assert req.focus_options == ["alert"]


def test_create_report_request_rejects_empty_focus_options():
    with pytest.raises(ValidationError):
        CreateReportRequest(
            input_id="tenant_demo",
            template_id="mss_test_rag",
            focus_options=[],
        )


def test_rag_service_split_text_respects_overlap():
    service = RAGService()
    text = "A" * 5000
    chunks = service._split_text(text, page_no=1)

    assert len(chunks) >= 2
    assert all(chunk["page_no"] == 1 for chunk in chunks)
    assert all(chunk["text"] for chunk in chunks)
    assert chunks[0]["text"][-40:] == chunks[1]["text"][:40]


def test_extract_markdown_sections_parses_scope_and_heading(tmp_path):
    service = RAGService()
    md = tmp_path / "kb.md"
    md.write_text(
        (
            "# Title\n\n"
            "<!-- rag:scope=common -->\n"
            "## 1. Shared\n"
            + ("Shared body. " * 10)
            + "\n\n<!-- rag:scope=slide slide_key=incident_effectiveness -->\n"
            + "## 5. Incident\n"
            + ("Incident body. " * 10)
            + "\n\n### 5.1 Detail\n"
            + ("Detail body. " * 10)
        ),
        encoding="utf-8",
    )

    sections = service._extract_markdown_sections(md)

    assert len(sections) == 3
    assert sections[0].section_scope == "common"
    assert sections[0].slide_key == "__common__"
    assert sections[1].slide_key == "incident_effectiveness"
    assert sections[2].heading_path.endswith("5.1 Detail")


def test_split_markdown_section_keeps_section_metadata():
    service = RAGService()
    section = MarkdownSection(
        text=("Paragraph one.\n\n" + "Paragraph two. " * 120).strip(),
        section_scope="slide",
        slide_key="incident_effectiveness",
        section_title="5.1 Detail",
        heading_path="5 Incident / 5.1 Detail",
    )

    chunks = service._split_markdown_section(section)

    assert chunks
    assert all(chunk["slide_key"] == "incident_effectiveness" for chunk in chunks)
    assert all(chunk["section_scope"] == "slide" for chunk in chunks)
    assert all(chunk["heading_path"] == "5 Incident / 5.1 Detail" for chunk in chunks)

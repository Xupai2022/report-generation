from __future__ import annotations

import unittest
from types import SimpleNamespace

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository


class LLMJsonStabilityTestCase(unittest.TestCase):
    def test_stream_extraction_ignores_reasoning_content(self) -> None:
        orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
        delta = SimpleNamespace(
            content='{"slides":[]}',
            reasoning_content="{slides: not valid json}",
        )

        self.assertEqual(orchestrator._extract_stream_content_piece(delta), '{"slides":[]}')

    def test_stream_extraction_keeps_text_compatibility_field(self) -> None:
        orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
        delta = SimpleNamespace(
            text='{"slides":[]}',
            reasoning_content="{slides: not valid json}",
        )

        self.assertEqual(orchestrator._extract_stream_content_piece(delta), '{"slides":[]}')

    def test_vuln_summary_prompt_preserves_outer_json_contract(self) -> None:
        template = TemplateRepository().get_descriptor_v2("mss_classic_ops_2")
        orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
        prompt = orchestrator._build_user_prompt_for_slides(
            tenant_input=TenantInput(
                raw={
                    "risk_prevention_work_details": {
                        "high_risk_exploitable_vulnerability_count": 26,
                        "high_risk_exploitable_vulnerability_closure_rate": "25%",
                    },
                    "vulnerability_effectiveness": {
                        "vulnerability_distribution": {"高危": 2548},
                    },
                }
            ),
            template=template,
            slide_keys=["risk_prevention_work_details"],
            focus_options=["business_protection", "vulnerability", "alert"],
        )

        self.assertIn("输出必须是合法 JSON，且只能输出 JSON。", prompt)
        self.assertIn("最终响应必须仍然是外层指定的完整 JSON", prompt)
        self.assertIn("不要直接输出本字段值本身", prompt)


if __name__ == "__main__":
    unittest.main()

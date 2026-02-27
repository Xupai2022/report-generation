from types import SimpleNamespace
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2


def test_p16_combo_normalizes_string_and_percent_rates():
    orchestrator = LLMOrchestratorV2.__new__(LLMOrchestratorV2)
    tenant_input = SimpleNamespace(
        raw={
            "critical_assurance": {
                "P16_combo": {
                    "categories": ["A", "B", "C", "D", "E", "F"],
                    "attack_counts": [100, "200", "300.00", None, "bad", 12.5],
                    "defense_rates": ["0.50", 50, "50%", 0.98, "bad", None],
                }
            }
        }
    )

    result = orchestrator._extract_chart_data(
        tenant_input,
        {"data_source": "critical_assurance.P16_combo"},
        "P16_combo",
    )

    assert result["attack_counts"] == [100, 200, 300, 0, 0, 12.5]
    assert result["defense_rates"] == [0.5, 0.5, 0.5, 0.98, 0, 0]

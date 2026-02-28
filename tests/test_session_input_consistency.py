import glob
import json
from pathlib import Path
from types import SimpleNamespace
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mss_ai_ppt_sample_assets.backend.models.slidespec import SlideContentV2, SlideSpecV2
from mss_ai_ppt_sample_assets.backend.services.report_service import ReportService


def _build_test_service(tmp_path: Path) -> ReportService:
    session_dir = tmp_path / "session"
    session_dir.mkdir(parents=True, exist_ok=True)

    service = ReportService.__new__(ReportService)
    service.template_repo = SimpleNamespace(
        clear_cache=lambda: None,
        get_descriptor_v2=lambda _template_id: SimpleNamespace(
            slides=[
                SimpleNamespace(
                    slide_key="summary",
                    placeholders=[
                        SimpleNamespace(token="metric", type="text", source="coverage.metric"),
                        SimpleNamespace(token="AI_TEXT", type="text", source=None),
                    ],
                )
            ]
        ),
    )
    service.session_manager = SimpleNamespace(
        get_input_path=lambda _session_id: session_dir / "input.json",
        get_report_path=lambda _session_id, _template_id: session_dir / "report.pptx",
        get_slidespec_path=lambda _session_id, _template_id: session_dir / "slidespec_mss_classic_ops.json",
    )
    service.audit_logger = SimpleNamespace(log=lambda **_kwargs: None)
    service.ppt_generator_v2 = SimpleNamespace(
        render=lambda _slidespec, output_path: output_path.write_bytes(b"pptx")
    )
    return service


def test_manual_rewrite_syncs_session_input_json(tmp_path):
    service = _build_test_service(tmp_path)
    session_id = "abc_20260228120000000000"
    job_id = f"{session_id}:mss_classic_ops"

    input_path = service.session_manager.get_input_path(session_id)
    input_path.write_text(json.dumps({"coverage": {"metric": 10}}), encoding="utf-8")

    slidespec = SlideSpecV2(
        template_id="mss_classic_ops",
        slides=[
            SlideContentV2(
                slide_no=1,
                slide_key="summary",
                placeholders={"metric": 10, "AI_TEXT": "before"},
            )
        ],
    )
    service._load_slidespec = lambda _session_id, _template_id: slidespec
    service._get_input_id_for_job = lambda _job_id: "classic_ops_dataxlsx"

    result = service.rewrite(
        job_id=job_id,
        slide_key="summary",
        new_content={"metric": 25},
    )

    updated_input = json.loads(input_path.read_text(encoding="utf-8"))
    assert updated_input["coverage"]["metric"] == 25
    assert result["updated_count"] == 1
    assert slidespec.get_slide("summary").placeholders["metric"] == 25


def test_ai_rewrite_uses_session_input_json(tmp_path):
    service = _build_test_service(tmp_path)
    session_id = "abc_20260228120000000001"
    job_id = f"{session_id}:mss_classic_ops"

    input_path = service.session_manager.get_input_path(session_id)
    input_path.write_text(json.dumps({"coverage": {"metric": 88}}), encoding="utf-8")

    slidespec = SlideSpecV2(
        template_id="mss_classic_ops",
        slides=[
            SlideContentV2(
                slide_no=1,
                slide_key="summary",
                placeholders={"metric": 88, "AI_TEXT": "old"},
            )
        ],
    )
    service._load_slidespec = lambda _session_id, _template_id: slidespec
    service._get_input_id_for_job = lambda _job_id: "classic_ops_dataxlsx"

    # Guard against regression to static data/inputs JSON loading.
    service.load_input = lambda _input_id: (_ for _ in ()).throw(
        AssertionError("ai_rewrite should use session input.json, not static load_input")
    )

    captured = {}

    def _rewrite_single_slide_v2(**kwargs):
        captured["tenant_input_raw"] = kwargs["tenant_input"].raw
        captured["target_tokens"] = kwargs.get("target_tokens")
        return {
            "placeholders": {"AI_TEXT": "rewritten"},
            "warnings": [],
            "updated_tokens": ["AI_TEXT"],
        }

    service.llm_orchestrator_v2 = SimpleNamespace(
        rewrite_single_slide_v2=_rewrite_single_slide_v2
    )

    result = service.ai_rewrite_slide(
        job_id=job_id,
        slide_key="summary",
        user_prompt="rewrite",
        target_tokens=["AI_TEXT"],
    )

    assert captured["tenant_input_raw"]["coverage"]["metric"] == 88
    assert captured["target_tokens"] == ["AI_TEXT"]
    assert result["updated_count"] == 1
    assert slidespec.get_slide("summary").placeholders["AI_TEXT"] == "rewritten"


def test_classic_template_closure_rate2_source_mapping():
    descriptor_paths = glob.glob("mss_ai_ppt_sample_assets/backend/data/templates/*_descriptor.json")
    assert descriptor_paths, "classic descriptor file not found"

    with open(descriptor_paths[0], "r", encoding="utf-8") as f:
        descriptor = json.load(f)

    coverage_slide = next(s for s in descriptor["slides"] if s["slide_key"] == "coverage_summary")
    source_map = {ph["token"]: ph.get("source") for ph in coverage_slide["placeholders"]}
    assert source_map["closure_rate2"] == "coverage_summary.closure_rate2"

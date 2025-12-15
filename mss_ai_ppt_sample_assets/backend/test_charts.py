"""
Test script for chart and table rendering functionality.

This script generates a report with charts (bar chart, pie chart) and native tables
to verify the new visualization features.
"""

import sys
from pathlib import Path

# Add parent directory to path
backend_dir = Path(__file__).parent.parent.parent
sys.path.insert(0, str(backend_dir))

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.modules.template_loader import TemplateRepository
from mss_ai_ppt_sample_assets.backend.modules.llm_orchestrator import LLMOrchestratorV2
from mss_ai_ppt_sample_assets.backend.modules.ppt_generator import PPTGeneratorV2

def test_chart_generation():
    """Test chart and table generation"""
    print("=" * 80)
    print("Testing Chart and Table Generation")
    print("=" * 80)

    # 1. Load input data
    input_file = config.DATA_DIR / "inputs" / "tenant_acme_2025-12_mss_input.json"
    print(f"\n1. Loading input data: {input_file.name}")
    tenant_input = TenantInput.load_from_file(input_file)
    print(f"   ✓ Loaded tenant: {tenant_input.get('tenant', {}).get('name')}")

    # 2. Load template
    template_id = "mss_technical_v2"
    print(f"\n2. Loading template: {template_id}")
    template_repo = TemplateRepository()
    template_desc = template_repo.get_descriptor_v2(template_id)
    print(f"   ✓ Template loaded: {len(template_desc.slides)} slides")

    # Find charts_demo slide
    charts_slide = None
    for slide in template_desc.slides:
        if slide.slide_key == 'charts_demo':
            charts_slide = slide
            break

    if charts_slide:
        print(f"   ✓ Found charts_demo slide at position {charts_slide.slide_no}")
        print(f"   ✓ Placeholders: {[ph.token for ph in charts_slide.placeholders]}")
    else:
        print("   ✗ charts_demo slide not found!")
        return

    # 3. Generate SlideSpec
    print(f"\n3. Generating SlideSpec (mock mode)...")
    orchestrator = LLMOrchestratorV2(template_repo)
    slidespec = orchestrator.generate_slidespec_v2(
        tenant_input,
        template_id,
        use_mock=True  # Use mock/fallback mode
    )
    print(f"   ✓ SlideSpec generated: {len(slidespec.slides)} slides")

    # Check charts_demo slide data
    charts_demo_data = slidespec.get_slide('charts_demo')
    if charts_demo_data:
        print(f"\n4. Verifying charts_demo data:")
        for token, value in charts_demo_data.placeholders.items():
            if isinstance(value, dict):
                keys = list(value.keys())
                print(f"   - {token}: dict with keys {keys}")
            else:
                print(f"   - {token}: {type(value).__name__} = {str(value)[:50]}")
    else:
        print("   ✗ No data for charts_demo slide!")

    # 4. Generate PPTX
    print(f"\n5. Generating PPTX...")
    generator = PPTGeneratorV2(template_repo)
    output_file = config.OUTPUTS_DIR / "reports" / f"test_charts_{template_id}.pptx"
    output_path = generator.render(slidespec, output_file)
    print(f"   ✓ PPTX saved to: {output_path}")
    print(f"   ✓ File size: {output_path.stat().st_size / 1024:.1f} KB")

    print("\n" + "=" * 80)
    print("Test Complete!")
    print("=" * 80)
    print(f"\nOpen the generated file to verify:")
    print(f"  {output_path}")
    print("\nExpected content on slide 3 (charts_demo):")
    print("  - Bar chart: Alert trend over 4 weeks")
    print("  - Pie chart: Alert severity distribution")
    print("  - Native table: Top alert rules with counts and false positive rates")

if __name__ == "__main__":
    test_chart_generation()

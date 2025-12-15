"""
Script to generate V2 PPTX templates for MSS reports.
All content is placeholders - NO hardcoded text.

Run this script to create:
- mss_executive_v2.pptx (8 slides, management version)
- mss_technical_v2.pptx (10 slides, technical version)
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


# Colors
LIGHT_PRIMARY = RGBColor(30, 64, 175)      # #1E40AF - Dark blue
LIGHT_ACCENT = RGBColor(59, 130, 246)      # #3B82F6 - Blue

DARK_PRIMARY = RGBColor(15, 23, 42)        # #0F172A - Dark slate
DARK_ACCENT = RGBColor(34, 197, 94)        # #22C55E - Green


def add_placeholder_box(slide, left, top, width, height, token, font_size=12, bold=False, color=None):
    """Add a placeholder text box with {{TOKEN}} format."""
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"{{{{{token}}}}}"
    p.font.size = Pt(font_size)
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    return shape


def create_executive_template():
    """Create the management/executive version template (8 slides).
    ALL content is placeholders - no hardcoded text.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 2.5, 12, 1, "REPORT_TITLE", font_size=36, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 3.8, 6, 0.5, "CUSTOMER_LABEL", font_size=18)
    add_placeholder_box(slide, 0.5, 4.4, 6, 0.5, "PERIOD_LABEL", font_size=14)
    add_placeholder_box(slide, 0.5, 5.0, 6, 0.5, "CONFIDENTIALITY_LABEL", font_size=12)
    add_placeholder_box(slide, 0.5, 6.5, 6, 0.5, "GENERATED_AT_LABEL", font_size=10)

    # Slide 2: Executive Summary
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.0, 12, 0.8, "HEADLINE", font_size=18, bold=True)
    # KPI section - all placeholders
    add_placeholder_box(slide, 0.5, 1.9, 12.5, 1.2, "KPI_SECTION", font_size=11)
    # Summary paragraph
    add_placeholder_box(slide, 0.5, 3.3, 12, 1.5, "SUMMARY_PARAGRAPH", font_size=12)
    # Key insights
    add_placeholder_box(slide, 0.5, 5.0, 12, 0.4, "KEY_INSIGHTS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.4, 12, 1.8, "KEY_INSIGHTS", font_size=11)

    # Slide 3: Alert Trends
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.2, 6, 0.4, "TREND_SECTION_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 1.5, "TREND_ANALYSIS", font_size=11)
    add_placeholder_box(slide, 7, 1.2, 5.5, 0.4, "TOP_CATEGORIES_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 1.5, "TOP_CATEGORIES_LIST", font_size=11)
    add_placeholder_box(slide, 0.5, 3.5, 12, 0.4, "CATEGORIES_INSIGHT_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.9, 12, 1.5, "TOP_CATEGORIES_INSIGHT", font_size=11)
    add_placeholder_box(slide, 0.5, 5.6, 12, 0.4, "MOM_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 6.0, 12, 1.2, "MOM_COMPARISON", font_size=11)

    # Slide 4: Major Incidents
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.2, 12, 0.4, "INCIDENT_OVERVIEW_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 1.0, "INCIDENT_SUMMARY", font_size=11)
    add_placeholder_box(slide, 0.5, 2.8, 12, 0.4, "INCIDENT_DETAILS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.2, 12, 2.5, "INCIDENT_DETAILS", font_size=10)
    add_placeholder_box(slide, 0.5, 5.9, 12, 0.4, "INCIDENT_INSIGHT_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 6.3, 12, 1.0, "INCIDENT_INSIGHT", font_size=11)

    # Slide 5: Vulnerability & Exposure
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.0, 12, 0.4, "VULN_STATS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.4, 12, 0.5, "VULN_STATS", font_size=11)
    add_placeholder_box(slide, 0.5, 2.0, 6, 0.4, "VULN_OVERVIEW_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 2.4, 6, 1.0, "VULN_OVERVIEW", font_size=11)
    add_placeholder_box(slide, 7, 2.0, 5.5, 0.4, "EXPOSURE_STATS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 2.4, 5.5, 1.0, "EXPOSURE_STATS", font_size=11)
    add_placeholder_box(slide, 0.5, 3.6, 12, 0.4, "TOP_CVE_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.0, 12, 1.2, "TOP_CVE_LIST", font_size=10)
    add_placeholder_box(slide, 0.5, 5.4, 12, 0.4, "CVE_ANALYSIS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.8, 12, 1.0, "TOP_CVE_ANALYSIS", font_size=10)
    add_placeholder_box(slide, 0.5, 6.9, 12, 0.5, "EXPOSURE_SUMMARY", font_size=9)

    # Slide 6: Cloud Security
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.2, 3, 0.4, "CLOUD_ACCOUNTS_TITLE", font_size=12)
    add_placeholder_box(slide, 0.5, 1.6, 3, 0.5, "CLOUD_ACCOUNTS_COUNT", font_size=20)
    add_placeholder_box(slide, 0.5, 2.3, 6, 0.4, "CLOUD_RISK_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 2.7, 6, 1.5, "CLOUD_RISK_LIST", font_size=11)
    add_placeholder_box(slide, 7, 2.3, 5.5, 0.4, "CLOUD_ANALYSIS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 2.7, 5.5, 1.5, "CLOUD_RISK_SUMMARY", font_size=11)
    add_placeholder_box(slide, 0.5, 4.5, 12, 0.4, "CLOUD_REC_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.9, 12, 2.3, "CLOUD_RECOMMENDATIONS", font_size=11)

    # Slide 7: Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.2, 6, 0.4, "P0_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 2.5, "P0_ACTIONS", font_size=11)
    add_placeholder_box(slide, 7, 1.2, 5.5, 0.4, "P1_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 2.5, "P1_ACTIONS", font_size=11)
    add_placeholder_box(slide, 0.5, 4.5, 12, 0.4, "STRATEGIC_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.9, 12, 2.3, "STRATEGIC_RECOMMENDATIONS", font_size=11)

    # Slide 8: Appendix
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_placeholder_box(slide, 0.5, 1.2, 6, 0.4, "DATA_SCOPE_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 1.2, "DATA_SCOPE", font_size=10)
    add_placeholder_box(slide, 7, 1.2, 5.5, 0.4, "ASSET_COVERAGE_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 0.5, "ASSET_COVERAGE", font_size=10)
    add_placeholder_box(slide, 0.5, 3.0, 6, 0.4, "SLA_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.4, 6, 0.5, "SLA_NOTES", font_size=10)
    add_placeholder_box(slide, 0.5, 4.2, 12, 0.4, "TERMINOLOGY_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.6, 12, 2.0, "TERMINOLOGY", font_size=10)
    add_placeholder_box(slide, 0.5, 6.8, 6, 0.5, "CONTACT_INFO", font_size=10)

    return prs


def create_technical_template():
    """Create the technical version template (10 slides).
    ALL content is placeholders - no hardcoded text.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 2.5, 12, 1, "REPORT_TITLE", font_size=36, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 3.8, 6, 0.5, "CUSTOMER_LABEL", font_size=18)
    add_placeholder_box(slide, 0.5, 4.4, 6, 0.5, "PERIOD_LABEL", font_size=14)
    add_placeholder_box(slide, 0.5, 5.0, 6, 0.5, "VERSION_TAG", font_size=12, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 5.6, 6, 0.5, "CONFIDENTIALITY_LABEL", font_size=12)
    add_placeholder_box(slide, 0.5, 6.5, 6, 0.5, "GENERATED_AT_LABEL", font_size=10)

    # Slide 2: Security Dashboard
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    # KPI Dashboard - all in one section
    add_placeholder_box(slide, 0.5, 1.0, 12.5, 3.5, "KPI_DASHBOARD", font_size=11)
    # Assessment
    add_placeholder_box(slide, 0.5, 4.8, 12, 0.4, "ASSESSMENT_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.2, 12, 2.0, "TECHNICAL_ASSESSMENT", font_size=11)

    # Slide 3: Alert Deep Analysis
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.0, 12, 0.4, "SEVERITY_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.4, 12, 1.0, "SEVERITY_BREAKDOWN", font_size=11)
    add_placeholder_box(slide, 0.5, 2.6, 6, 0.4, "TOP_RULES_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.0, 6, 1.5, "TOP_RULES_TABLE", font_size=9)
    add_placeholder_box(slide, 7, 2.6, 5.5, 0.4, "RULES_ANALYSIS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 3.0, 5.5, 1.5, "TOP_RULES_ANALYSIS", font_size=10)
    add_placeholder_box(slide, 0.5, 4.8, 12, 0.4, "FP_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.2, 12, 2.0, "FALSE_POSITIVE_INSIGHT", font_size=11)

    # Slide 4: Incident Timeline
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.2, 12, 0.4, "TIMELINE_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.2, "TIMELINE_NARRATIVE", font_size=11)
    add_placeholder_box(slide, 0.5, 4.0, 6, 0.4, "RESPONSE_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.4, 6, 1.5, "RESPONSE_METRICS", font_size=11)
    add_placeholder_box(slide, 7, 4.0, 5.5, 0.4, "LESSONS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.4, 5.5, 2.8, "LESSONS_LEARNED", font_size=11)

    # Slide 5: Incident Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.2, 12, 0.4, "INCIDENT1_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.2, "INCIDENT_DETAIL_1", font_size=10)
    add_placeholder_box(slide, 0.5, 4.0, 12, 0.4, "INCIDENT2_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.4, 12, 2.2, "INCIDENT_DETAIL_2", font_size=10)
    add_placeholder_box(slide, 0.5, 6.8, 12, 0.5, "ATTACK_PATTERN_SUMMARY", font_size=9)

    # Slide 6: Vulnerability Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.2, 12, 0.4, "VULN_DIST_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 1.5, "VULN_DISTRIBUTION", font_size=11)
    add_placeholder_box(slide, 0.5, 3.3, 12, 0.4, "CVE_ANALYSIS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.7, 12, 2.0, "CVE_TECHNICAL_ANALYSIS", font_size=10)
    add_placeholder_box(slide, 0.5, 5.9, 12, 0.4, "PATCH_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 6.3, 12, 1.0, "PATCH_STATUS", font_size=10)

    # Slide 7: Attack Surface
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.0, 6, 0.4, "EXPOSED_SERVICES_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.4, 6, 1.2, "EXPOSED_SERVICES_TABLE", font_size=9)
    add_placeholder_box(slide, 7, 1.0, 5.5, 0.4, "EXPOSURE_ANALYSIS_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.4, 5.5, 1.2, "EXTERNAL_EXPOSURE", font_size=10)
    add_placeholder_box(slide, 0.5, 2.8, 6, 0.4, "SURFACE_TREND_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.2, 6, 1.5, "ATTACK_SURFACE_TREND", font_size=11)
    add_placeholder_box(slide, 7, 2.8, 5.5, 0.4, "REMEDIATION_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 3.2, 5.5, 4.0, "EXPOSURE_REMEDIATION", font_size=10)

    # Slide 8: Cloud Security Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.2, 12, 0.4, "CLOUD_MISCONFIG_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.0, "CLOUD_MISCONFIG_DETAILS", font_size=10)
    add_placeholder_box(slide, 0.5, 3.8, 6, 0.4, "IAM_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.2, 6, 1.5, "IAM_RISK_ANALYSIS", font_size=10)
    add_placeholder_box(slide, 7, 3.8, 5.5, 0.4, "COMPLIANCE_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.2, 5.5, 3.0, "CLOUD_COMPLIANCE", font_size=10)

    # Slide 9: Technical Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.2, 12, 0.4, "TECH_REC_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.0, "TECHNICAL_RECOMMENDATIONS", font_size=10)
    add_placeholder_box(slide, 0.5, 3.8, 6, 0.4, "DETECTION_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.2, 6, 1.8, "DETECTION_IMPROVEMENTS", font_size=10)
    add_placeholder_box(slide, 7, 3.8, 5.5, 0.4, "HARDENING_TITLE", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.2, 5.5, 3.0, "HARDENING_CHECKLIST", font_size=10)

    # Slide 10: Appendix
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_placeholder_box(slide, 0.5, 0.3, 12, 0.6, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    add_placeholder_box(slide, 0.5, 1.0, 6, 0.4, "EVIDENCE_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 1.3, 6, 0.8, "EVIDENCE_INDEX", font_size=9)
    add_placeholder_box(slide, 7, 1.0, 5.5, 0.4, "FRESHNESS_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 7, 1.3, 5.5, 0.4, "DATA_FRESHNESS", font_size=10)
    add_placeholder_box(slide, 0.5, 2.2, 12, 0.4, "DATA_SCOPE_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 2.6, 12, 0.6, "DATA_SCOPE_NOTES", font_size=10)
    add_placeholder_box(slide, 0.5, 3.4, 6, 0.4, "VULN_NOTES_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 3.7, 6, 0.4, "VULN_NOTES", font_size=10)
    add_placeholder_box(slide, 7, 3.4, 5.5, 0.4, "EVIDENCE_NOTES_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 7, 3.7, 5.5, 0.4, "EVIDENCE_NOTES", font_size=10)
    add_placeholder_box(slide, 0.5, 4.3, 6, 0.4, "ASSETS_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 4.6, 6, 0.8, "ASSET_GROUPS_LIST", font_size=10)
    add_placeholder_box(slide, 7, 4.3, 5.5, 0.4, "SERVICES_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 7, 4.6, 5.5, 0.5, "KEY_SERVICES_LIST", font_size=10)
    add_placeholder_box(slide, 0.5, 5.6, 12, 0.4, "TERMINOLOGY_TITLE", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 6.0, 12, 1.2, "TERMINOLOGY_TECHNICAL", font_size=9)

    return prs


def main():
    """Generate both V2 templates."""
    output_dir = Path(__file__).parent / "data" / "templates"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Generate executive template
    print("Generating mss_executive_v2.pptx...")
    exec_prs = create_executive_template()
    exec_path = output_dir / "mss_executive_v2.pptx"
    exec_prs.save(exec_path)
    print(f"  Saved to: {exec_path}")

    # Generate technical template
    print("Generating mss_technical_v2.pptx...")
    tech_prs = create_technical_template()
    tech_path = output_dir / "mss_technical_v2.pptx"
    tech_prs.save(tech_path)
    print(f"  Saved to: {tech_path}")

    print("\nDone! V2 templates generated successfully.")


if __name__ == "__main__":
    main()

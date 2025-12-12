"""
Script to generate V2 PPTX templates for MSS reports.

Run this script to create:
- mss_executive_v2.pptx (8 slides, management version)
- mss_technical_v2.pptx (10 slides, technical version)
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


# Colors
LIGHT_PRIMARY = RGBColor(30, 64, 175)      # #1E40AF - Dark blue
LIGHT_ACCENT = RGBColor(59, 130, 246)      # #3B82F6 - Blue
LIGHT_BG = RGBColor(248, 250, 252)         # #F8FAFC - Light gray

DARK_PRIMARY = RGBColor(15, 23, 42)        # #0F172A - Dark slate
DARK_ACCENT = RGBColor(34, 197, 94)        # #22C55E - Green
DARK_BG = RGBColor(30, 41, 59)             # #1E293B - Slate


def add_text_box(slide, left, top, width, height, text, font_size=12, bold=False, color=None):
    """Add a text box to slide."""
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    return shape


def add_placeholder_box(slide, left, top, width, height, token, label=None, font_size=11):
    """Add a placeholder text box with {{TOKEN}} format."""
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"{{{{{token}}}}}"
    p.font.size = Pt(font_size)
    return shape


def create_executive_template():
    """Create the management/executive version template (8 slides)."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    add_text_box(slide, 0.5, 2.5, 12, 1, "{{REPORT_TITLE}}", font_size=36, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 3.8, 6, 0.5, "客户：{{CUSTOMER_NAME}}", font_size=18)
    add_text_box(slide, 0.5, 4.4, 6, 0.5, "周期：{{PERIOD}}", font_size=14)
    add_text_box(slide, 0.5, 5.0, 6, 0.5, "保密等级：{{CONFIDENTIALITY}}", font_size=12)
    add_text_box(slide, 0.5, 6.5, 6, 0.5, "生成时间：{{GENERATED_AT}}", font_size=10)

    # Slide 2: Executive Summary
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.0, 12, 0.8, "{{HEADLINE}}", font_size=18, bold=True)
    # KPI cards row
    add_text_box(slide, 0.5, 2.0, 2.5, 0.3, "告警总数", font_size=10)
    add_placeholder_box(slide, 0.5, 2.3, 2.5, 0.5, "KPI_ALERTS_TOTAL", font_size=24)
    add_text_box(slide, 3.2, 2.0, 2.5, 0.3, "高危告警", font_size=10)
    add_placeholder_box(slide, 3.2, 2.3, 2.5, 0.5, "KPI_ALERTS_HIGH", font_size=24)
    add_text_box(slide, 5.9, 2.0, 2.5, 0.3, "高危事件", font_size=10)
    add_placeholder_box(slide, 5.9, 2.3, 2.5, 0.5, "KPI_INCIDENTS_HIGH", font_size=24)
    add_text_box(slide, 8.6, 2.0, 2.5, 0.3, "平均MTTR", font_size=10)
    add_placeholder_box(slide, 8.6, 2.3, 2.5, 0.5, "KPI_MTTR_HOURS", font_size=24)
    # Summary paragraph
    add_placeholder_box(slide, 0.5, 3.2, 12, 1.5, "SUMMARY_PARAGRAPH", font_size=12)
    # Key insights
    add_text_box(slide, 0.5, 5.0, 12, 0.4, "关键洞察", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.4, 12, 1.8, "KEY_INSIGHTS", font_size=11)

    # Slide 3: Alert Trends
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.2, 6, 0.4, "趋势分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 1.5, "TREND_ANALYSIS", font_size=11)
    add_text_box(slide, 7, 1.2, 5.5, 0.4, "Top告警类别", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 1.5, "TOP_CATEGORIES_LIST", font_size=11)
    add_text_box(slide, 0.5, 3.5, 12, 0.4, "类别分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.9, 12, 1.5, "TOP_CATEGORIES_INSIGHT", font_size=11)
    add_text_box(slide, 0.5, 5.6, 12, 0.4, "环比对比", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 6.0, 12, 1.2, "MOM_COMPARISON", font_size=11)

    # Slide 4: Major Incidents
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "事件概述", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 1.0, "INCIDENT_SUMMARY", font_size=11)
    add_text_box(slide, 0.5, 2.8, 12, 0.4, "重点事件详情", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.2, 12, 2.5, "INCIDENT_DETAILS", font_size=10)
    add_text_box(slide, 0.5, 5.9, 12, 0.4, "安全洞察", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 6.3, 12, 1.0, "INCIDENT_INSIGHT", font_size=11)

    # Slide 5: Vulnerability & Exposure
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.2, 6, 0.4, "漏洞态势", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 1.2, "VULN_OVERVIEW", font_size=11)
    add_text_box(slide, 7, 1.2, 5.5, 0.4, "暴露面统计", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 0.5, "EXPOSURE_STATS", font_size=11)
    add_text_box(slide, 0.5, 3.0, 12, 0.4, "Top CVE分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.4, 12, 1.8, "TOP_CVE_ANALYSIS", font_size=10)
    add_text_box(slide, 0.5, 5.4, 12, 0.4, "暴露面分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.8, 12, 1.5, "EXPOSURE_SUMMARY", font_size=11)

    # Slide 6: Cloud Security
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.2, 3, 0.4, "云账号数", font_size=12)
    add_placeholder_box(slide, 0.5, 1.6, 3, 0.5, "CLOUD_ACCOUNTS_COUNT", font_size=20)
    add_text_box(slide, 0.5, 2.3, 6, 0.4, "云风险清单", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 2.7, 6, 1.5, "CLOUD_RISK_LIST", font_size=11)
    add_text_box(slide, 7, 2.3, 5.5, 0.4, "风险分析", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 2.7, 5.5, 1.5, "CLOUD_RISK_SUMMARY", font_size=11)
    add_text_box(slide, 0.5, 4.5, 12, 0.4, "云安全建议", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.9, 12, 2.3, "CLOUD_RECOMMENDATIONS", font_size=11)

    # Slide 7: Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.2, 6, 0.4, "P0 紧急整改（7天内）", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 2.5, "P0_ACTIONS", font_size=11)
    add_text_box(slide, 7, 1.2, 5.5, 0.4, "P1 重要整改（30天内）", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 2.5, "P1_ACTIONS", font_size=11)
    add_text_box(slide, 0.5, 4.5, 12, 0.4, "中长期安全建设建议", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.9, 12, 2.3, "STRATEGIC_RECOMMENDATIONS", font_size=11)

    # Slide 8: Appendix
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=LIGHT_PRIMARY)
    add_text_box(slide, 0.5, 1.2, 6, 0.4, "数据口径说明", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 6, 1.2, "DATA_SCOPE", font_size=10)
    add_text_box(slide, 7, 1.2, 5.5, 0.4, "资产覆盖", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 1.6, 5.5, 0.5, "ASSET_COVERAGE", font_size=10)
    add_text_box(slide, 0.5, 3.0, 6, 0.4, "SLA说明", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.4, 6, 0.5, "SLA_NOTES", font_size=10)
    add_text_box(slide, 0.5, 4.2, 12, 0.4, "术语解释", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.6, 12, 2.0, "TERMINOLOGY", font_size=10)
    add_placeholder_box(slide, 0.5, 6.8, 6, 0.5, "CONTACT_INFO", font_size=10)

    return prs


def create_technical_template():
    """Create the technical version template (10 slides)."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 2.5, 12, 1, "{{REPORT_TITLE}}", font_size=36, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 3.8, 6, 0.5, "客户：{{CUSTOMER_NAME}}", font_size=18)
    add_text_box(slide, 0.5, 4.4, 6, 0.5, "周期：{{PERIOD}}", font_size=14)
    add_text_box(slide, 0.5, 5.0, 6, 0.5, "{{VERSION_TAG}}", font_size=12, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 5.6, 6, 0.5, "保密等级：{{CONFIDENTIALITY}}", font_size=12)
    add_text_box(slide, 0.5, 6.5, 6, 0.5, "生成时间：{{GENERATED_AT}}", font_size=10)

    # Slide 2: Security Dashboard
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    # KPI Row 1
    y = 1.2
    for i, (label, token) in enumerate([
        ("告警总数", "KPI_ALERTS_TOTAL"), ("高危告警", "KPI_ALERTS_HIGH"),
        ("中危告警", "KPI_ALERTS_MEDIUM"), ("低危告警", "KPI_ALERTS_LOW")
    ]):
        x = 0.5 + i * 3.1
        add_text_box(slide, x, y, 2.8, 0.3, label, font_size=10)
        add_placeholder_box(slide, x, y + 0.3, 2.8, 0.5, token, font_size=20)
    # KPI Row 2
    y = 2.3
    for i, (label, token) in enumerate([
        ("事件总数", "KPI_INCIDENTS_TOTAL"), ("高危事件", "KPI_INCIDENTS_HIGH"),
        ("平均MTTD", "KPI_MTTD_MINUTES"), ("平均MTTR", "KPI_MTTR_HOURS")
    ]):
        x = 0.5 + i * 3.1
        add_text_box(slide, x, y, 2.8, 0.3, label, font_size=10)
        add_placeholder_box(slide, x, y + 0.3, 2.8, 0.5, token, font_size=20)
    # KPI Row 3
    y = 3.4
    for i, (label, token) in enumerate([
        ("严重漏洞", "KPI_VULN_CRITICAL"), ("高危漏洞", "KPI_VULN_HIGH"),
        ("误报率", "KPI_FALSE_POSITIVE"), ("EDR覆盖", "KPI_EDR_COVERAGE")
    ]):
        x = 0.5 + i * 3.1
        add_text_box(slide, x, y, 2.8, 0.3, label, font_size=10)
        add_placeholder_box(slide, x, y + 0.3, 2.8, 0.5, token, font_size=20)
    # Assessment
    add_text_box(slide, 0.5, 4.8, 12, 0.4, "技术评估", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.2, 12, 2.0, "TECHNICAL_ASSESSMENT", font_size=11)

    # Slide 3: Alert Deep Analysis
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "严重程度分布分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 1.3, "SEVERITY_BREAKDOWN", font_size=11)
    add_text_box(slide, 0.5, 3.1, 12, 0.4, "Top规则分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.5, 12, 1.8, "TOP_RULES_ANALYSIS", font_size=11)
    add_text_box(slide, 0.5, 5.5, 12, 0.4, "误报治理建议", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 5.9, 12, 1.3, "FALSE_POSITIVE_INSIGHT", font_size=11)

    # Slide 4: Incident Timeline
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "事件时间线叙述", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.2, "TIMELINE_NARRATIVE", font_size=11)
    add_text_box(slide, 0.5, 4.0, 6, 0.4, "响应效率分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.4, 6, 1.5, "RESPONSE_METRICS", font_size=11)
    add_text_box(slide, 7, 4.0, 5.5, 0.4, "经验教训", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.4, 5.5, 2.8, "LESSONS_LEARNED", font_size=11)

    # Slide 5: Incident Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "高危事件1 - 详细分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.2, "INCIDENT_DETAIL_1", font_size=10)
    add_text_box(slide, 0.5, 4.0, 12, 0.4, "高危事件2 - 详细分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.4, 12, 2.2, "INCIDENT_DETAIL_2", font_size=10)
    add_text_box(slide, 0.5, 6.8, 12, 0.4, "攻击模式总结", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 7.0, 12, 0.4, "ATTACK_PATTERN_SUMMARY", font_size=9)

    # Slide 6: Vulnerability Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "漏洞分布分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 1.5, "VULN_DISTRIBUTION", font_size=11)
    add_text_box(slide, 0.5, 3.3, 12, 0.4, "CVE技术分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 3.7, 12, 2.0, "CVE_TECHNICAL_ANALYSIS", font_size=10)
    add_text_box(slide, 0.5, 5.9, 12, 0.4, "补丁状态追踪", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 6.3, 12, 1.0, "PATCH_STATUS", font_size=10)

    # Slide 7: Attack Surface
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "对外暴露服务分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 1.8, "EXTERNAL_EXPOSURE", font_size=11)
    add_text_box(slide, 0.5, 3.6, 6, 0.4, "攻击面趋势", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.0, 6, 1.5, "ATTACK_SURFACE_TREND", font_size=11)
    add_text_box(slide, 7, 3.6, 5.5, 0.4, "收敛技术方案", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.0, 5.5, 3.2, "EXPOSURE_REMEDIATION", font_size=10)

    # Slide 8: Cloud Security Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "云配置问题详情", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.0, "CLOUD_MISCONFIG_DETAILS", font_size=10)
    add_text_box(slide, 0.5, 3.8, 6, 0.4, "IAM权限风险分析", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.2, 6, 1.5, "IAM_RISK_ANALYSIS", font_size=10)
    add_text_box(slide, 7, 3.8, 5.5, 0.4, "云安全合规检查", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.2, 5.5, 3.0, "CLOUD_COMPLIANCE", font_size=10)

    # Slide 9: Technical Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 12, 0.4, "技术整改方案", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 1.6, 12, 2.0, "TECHNICAL_RECOMMENDATIONS", font_size=10)
    add_text_box(slide, 0.5, 3.8, 6, 0.4, "检测能力提升", font_size=14, bold=True)
    add_placeholder_box(slide, 0.5, 4.2, 6, 1.8, "DETECTION_IMPROVEMENTS", font_size=10)
    add_text_box(slide, 7, 3.8, 5.5, 0.4, "加固检查清单", font_size=14, bold=True)
    add_placeholder_box(slide, 7, 4.2, 5.5, 3.0, "HARDENING_CHECKLIST", font_size=10)

    # Slide 10: Appendix
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text_box(slide, 0.5, 0.3, 12, 0.6, "{{SLIDE_TITLE}}", font_size=28, bold=True, color=DARK_ACCENT)
    add_text_box(slide, 0.5, 1.2, 6, 0.4, "数据新鲜度", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 1.5, 6, 0.4, "DATA_FRESHNESS", font_size=10)
    add_text_box(slide, 0.5, 2.0, 12, 0.4, "数据口径说明", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 2.4, 12, 0.8, "DATA_SCOPE_NOTES", font_size=10)
    add_text_box(slide, 0.5, 3.4, 6, 0.4, "漏洞统计说明", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 3.8, 6, 0.5, "VULN_NOTES", font_size=10)
    add_text_box(slide, 7, 3.4, 5.5, 0.4, "证据索引说明", font_size=12, bold=True)
    add_placeholder_box(slide, 7, 3.8, 5.5, 0.5, "EVIDENCE_NOTES", font_size=10)
    add_text_box(slide, 0.5, 4.5, 6, 0.4, "资产组", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 4.9, 6, 0.8, "ASSET_GROUPS_LIST", font_size=10)
    add_text_box(slide, 7, 4.5, 5.5, 0.4, "关键服务", font_size=12, bold=True)
    add_placeholder_box(slide, 7, 4.9, 5.5, 0.5, "KEY_SERVICES_LIST", font_size=10)
    add_text_box(slide, 0.5, 5.9, 12, 0.4, "技术术语表", font_size=12, bold=True)
    add_placeholder_box(slide, 0.5, 6.3, 12, 1.0, "TERMINOLOGY_TECHNICAL", font_size=9)

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

"""
Script to generate V2 PPTX templates for MSS reports.
All content is placeholders - NO hardcoded text.

Run this script to create:
- mss_executive_v2.pptx (8 slides, management version)
- mss_technical_v2.pptx (10 slides, technical version)
"""

from pathlib import Path
from typing import Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN


# Colors
LIGHT_PRIMARY = RGBColor(30, 64, 175)      # #1E40AF - Dark blue
LIGHT_ACCENT = RGBColor(59, 130, 246)      # #3B82F6 - Blue
LIGHT_BG = RGBColor(248, 250, 252)         # #F8FAFC
LIGHT_PANEL = RGBColor(255, 255, 255)      # #FFFFFF
LIGHT_BORDER = RGBColor(226, 232, 240)     # #E2E8F0
LIGHT_MUTED = RGBColor(100, 116, 139)      # #64748B

DARK_PRIMARY = RGBColor(15, 23, 42)        # #0F172A - Dark slate
DARK_ACCENT = RGBColor(34, 197, 94)        # #22C55E - Green
DARK_BG = RGBColor(11, 18, 32)             # #0B1220
DARK_PANEL = RGBColor(17, 24, 39)          # #111827
DARK_BORDER = RGBColor(31, 41, 55)         # #1F2937
DARK_TEXT = RGBColor(226, 232, 240)        # #E2E8F0
DARK_MUTED = RGBColor(148, 163, 184)       # #94A3B8


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
    p.font.name = "Microsoft YaHei"
    return shape


def _set_slide_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_rect(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: RGBColor,
    line_color: Optional[RGBColor] = None,
    transparency: Optional[float] = None,
):
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if transparency is not None:
        shape.fill.transparency = max(0.0, min(1.0, transparency))

    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    return shape


def _add_rounded(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: RGBColor,
    line_color: Optional[RGBColor] = None,
    transparency: Optional[float] = None,
):
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if transparency is not None:
        shape.fill.transparency = max(0.0, min(1.0, transparency))

    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    return shape


def _add_accent_blob(slide, left: float, top: float, width: float, height: float, color: RGBColor, transparency: float = 0.85):
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.fill.transparency = max(0.0, min(1.0, transparency))
    shape.line.fill.background()
    return shape


def _style_textbox(shape, font_color: RGBColor, font_size: Optional[int] = None, bold: Optional[bool] = None, align: Optional[PP_ALIGN] = None):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(10)
    tf.margin_right = Pt(10)
    tf.margin_top = Pt(6)
    tf.margin_bottom = Pt(6)
    for p in tf.paragraphs:
        if font_size is not None:
            p.font.size = Pt(font_size)
        if bold is not None:
            p.font.bold = bold
        p.font.color.rgb = font_color
        p.font.name = "Microsoft YaHei"
        if align is not None:
            p.alignment = align


def _add_card(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill: RGBColor,
    border: RGBColor,
    shadow: bool = True,
):
    if shadow:
        sh = _add_rounded(slide, left + 0.06, top + 0.06, width, height, RGBColor(0, 0, 0), line_color=None, transparency=0.88)
        sh.fill.transparency = 0.90
    return _add_rounded(slide, left, top, width, height, fill, line_color=border, transparency=0.0)


def _light_base(slide):
    _set_slide_bg(slide, LIGHT_BG)
    _add_rect(slide, 0, 0, 13.333, 1.05, RGBColor(239, 246, 255), transparency=0.0)
    _add_rect(slide, 0, 0.98, 13.333, 0.06, LIGHT_ACCENT, transparency=0.0)
    _add_accent_blob(slide, 11.4, -0.9, 3.2, 3.2, LIGHT_ACCENT, transparency=0.88)
    _add_accent_blob(slide, -1.2, 6.2, 2.6, 2.6, LIGHT_PRIMARY, transparency=0.90)


def _dark_base(slide):
    _set_slide_bg(slide, DARK_BG)
    _add_rect(slide, 0, 0, 13.333, 1.05, DARK_PRIMARY, transparency=0.0)
    _add_rect(slide, 0, 0.98, 13.333, 0.06, DARK_ACCENT, transparency=0.0)
    _add_accent_blob(slide, 11.2, -0.8, 3.4, 3.4, DARK_ACCENT, transparency=0.92)
    _add_accent_blob(slide, -1.4, 6.1, 3.0, 3.0, RGBColor(59, 130, 246), transparency=0.94)
    _add_rect(slide, 0.3, 1.25, 12.75, 6.0, RGBColor(15, 23, 42), transparency=0.35, line_color=None)

def create_executive_template():
    """Create the management/executive version template (8 slides).
    ALL content is placeholders - no hardcoded text.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    _add_rect(slide, 0, 0, 0.18, 7.5, LIGHT_PRIMARY, transparency=0.0)
    _add_card(slide, 0.65, 2.15, 12.05, 2.15, LIGHT_PANEL, LIGHT_BORDER)
    title = add_placeholder_box(slide, 0.85, 2.35, 11.65, 0.9, "REPORT_TITLE", font_size=36, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(title, LIGHT_PRIMARY, font_size=36, bold=True, align=PP_ALIGN.LEFT)
    _add_card(slide, 0.65, 4.55, 7.2, 2.25, LIGHT_PANEL, LIGHT_BORDER, shadow=False)
    customer = add_placeholder_box(slide, 0.85, 4.72, 6.8, 0.55, "CUSTOMER_LABEL", font_size=18, bold=True, color=DARK_PRIMARY)
    _style_textbox(customer, DARK_PRIMARY, font_size=18, bold=True)
    period = add_placeholder_box(slide, 0.85, 5.33, 6.8, 0.45, "PERIOD_LABEL", font_size=14, color=LIGHT_MUTED)
    _style_textbox(period, LIGHT_MUTED, font_size=14)
    conf = add_placeholder_box(slide, 0.85, 5.85, 6.8, 0.45, "CONFIDENTIALITY_LABEL", font_size=12, color=LIGHT_MUTED)
    _style_textbox(conf, LIGHT_MUTED, font_size=12)
    gen = add_placeholder_box(slide, 0.85, 6.35, 6.8, 0.45, "GENERATED_AT_LABEL", font_size=10, color=LIGHT_MUTED)
    _style_textbox(gen, LIGHT_MUTED, font_size=10)

    # Slide 2: Executive Summary
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    _add_rect(slide, 0, 1.05, 0.18, 6.45, LIGHT_ACCENT, transparency=0.15)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)
    _add_card(slide, 0.65, 1.05, 12.3, 0.95, LIGHT_PANEL, LIGHT_BORDER)
    hl = add_placeholder_box(slide, 0.85, 1.18, 11.9, 0.7, "HEADLINE", font_size=18, bold=True, color=DARK_PRIMARY)
    _style_textbox(hl, DARK_PRIMARY, font_size=18, bold=True)
    _add_card(slide, 0.65, 2.15, 12.3, 1.1, LIGHT_PANEL, LIGHT_BORDER)
    kpi = add_placeholder_box(slide, 0.85, 2.28, 11.9, 0.85, "KPI_SECTION", font_size=11, color=DARK_PRIMARY)
    _style_textbox(kpi, DARK_PRIMARY, font_size=11)
    _add_card(slide, 0.65, 3.38, 12.3, 1.55, LIGHT_PANEL, LIGHT_BORDER)
    sp = add_placeholder_box(slide, 0.85, 3.52, 11.9, 1.25, "SUMMARY_PARAGRAPH", font_size=12, color=DARK_PRIMARY)
    _style_textbox(sp, DARK_PRIMARY, font_size=12)
    _add_card(slide, 0.65, 5.05, 12.3, 2.25, LIGHT_PANEL, LIGHT_BORDER)
    kit = add_placeholder_box(slide, 0.85, 5.15, 11.9, 0.4, "KEY_INSIGHTS_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(kit, LIGHT_PRIMARY, font_size=14, bold=True)
    ki = add_placeholder_box(slide, 0.85, 5.55, 11.9, 1.65, "KEY_INSIGHTS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(ki, DARK_PRIMARY, font_size=11)

    # Slide 3: Alert Trends
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 6.05, 2.25, LIGHT_PANEL, LIGHT_BORDER)
    tst = add_placeholder_box(slide, 0.85, 1.15, 5.65, 0.4, "TREND_SECTION_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(tst, LIGHT_PRIMARY, font_size=14, bold=True)
    ta = add_placeholder_box(slide, 0.85, 1.55, 5.65, 1.65, "TREND_ANALYSIS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(ta, DARK_PRIMARY, font_size=11)

    _add_card(slide, 6.9, 1.05, 6.05, 2.25, LIGHT_PANEL, LIGHT_BORDER)
    tct = add_placeholder_box(slide, 7.1, 1.15, 5.65, 0.4, "TOP_CATEGORIES_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(tct, LIGHT_PRIMARY, font_size=14, bold=True)
    tcl = add_placeholder_box(slide, 7.1, 1.55, 5.65, 1.65, "TOP_CATEGORIES_LIST", font_size=11, color=DARK_PRIMARY)
    _style_textbox(tcl, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 3.45, 12.3, 1.95, LIGHT_PANEL, LIGHT_BORDER)
    cit = add_placeholder_box(slide, 0.85, 3.55, 11.9, 0.4, "CATEGORIES_INSIGHT_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(cit, LIGHT_PRIMARY, font_size=14, bold=True)
    tci = add_placeholder_box(slide, 0.85, 3.95, 11.9, 1.35, "TOP_CATEGORIES_INSIGHT", font_size=11, color=DARK_PRIMARY)
    _style_textbox(tci, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 5.55, 12.3, 1.7, LIGHT_PANEL, LIGHT_BORDER)
    mt = add_placeholder_box(slide, 0.85, 5.65, 11.9, 0.4, "MOM_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(mt, LIGHT_PRIMARY, font_size=14, bold=True)
    mc = add_placeholder_box(slide, 0.85, 6.05, 11.9, 1.1, "MOM_COMPARISON", font_size=11, color=DARK_PRIMARY)
    _style_textbox(mc, DARK_PRIMARY, font_size=11)

    # Slide 4: Major Incidents
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 1.25, LIGHT_PANEL, LIGHT_BORDER)
    iot = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "INCIDENT_OVERVIEW_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(iot, LIGHT_PRIMARY, font_size=14, bold=True)
    isum = add_placeholder_box(slide, 0.85, 1.55, 11.9, 0.68, "INCIDENT_SUMMARY", font_size=11, color=DARK_PRIMARY)
    _style_textbox(isum, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 2.45, 12.3, 3.25, LIGHT_PANEL, LIGHT_BORDER)
    idt = add_placeholder_box(slide, 0.85, 2.55, 11.9, 0.4, "INCIDENT_DETAILS_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(idt, LIGHT_PRIMARY, font_size=14, bold=True)
    idet = add_placeholder_box(slide, 0.85, 2.95, 11.9, 2.65, "INCIDENT_DETAILS", font_size=10, color=DARK_PRIMARY)
    _style_textbox(idet, DARK_PRIMARY, font_size=10)

    _add_card(slide, 0.65, 5.85, 12.3, 1.4, LIGHT_PANEL, LIGHT_BORDER)
    iit = add_placeholder_box(slide, 0.85, 5.95, 11.9, 0.4, "INCIDENT_INSIGHT_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(iit, LIGHT_PRIMARY, font_size=14, bold=True)
    ii = add_placeholder_box(slide, 0.85, 6.35, 11.9, 0.82, "INCIDENT_INSIGHT", font_size=11, color=DARK_PRIMARY)
    _style_textbox(ii, DARK_PRIMARY, font_size=11)

    # Slide 5: Vulnerability & Exposure
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 0.95, LIGHT_PANEL, LIGHT_BORDER)
    vst = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "VULN_STATS_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(vst, LIGHT_PRIMARY, font_size=14, bold=True)
    vs = add_placeholder_box(slide, 0.85, 1.52, 11.9, 0.43, "VULN_STATS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(vs, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 2.15, 6.05, 1.45, LIGHT_PANEL, LIGHT_BORDER)
    vot = add_placeholder_box(slide, 0.85, 2.25, 5.65, 0.4, "VULN_OVERVIEW_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(vot, LIGHT_PRIMARY, font_size=14, bold=True)
    vo = add_placeholder_box(slide, 0.85, 2.65, 5.65, 0.85, "VULN_OVERVIEW", font_size=11, color=DARK_PRIMARY)
    _style_textbox(vo, DARK_PRIMARY, font_size=11)

    _add_card(slide, 6.9, 2.15, 6.05, 1.45, LIGHT_PANEL, LIGHT_BORDER)
    est = add_placeholder_box(slide, 7.1, 2.25, 5.65, 0.4, "EXPOSURE_STATS_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(est, LIGHT_PRIMARY, font_size=14, bold=True)
    es = add_placeholder_box(slide, 7.1, 2.65, 5.65, 0.85, "EXPOSURE_STATS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(es, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 3.75, 12.3, 1.65, LIGHT_PANEL, LIGHT_BORDER)
    tct = add_placeholder_box(slide, 0.85, 3.85, 11.9, 0.4, "TOP_CVE_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(tct, LIGHT_PRIMARY, font_size=14, bold=True)
    tcl = add_placeholder_box(slide, 0.85, 4.25, 11.9, 1.05, "TOP_CVE_LIST", font_size=10, color=DARK_PRIMARY)
    _style_textbox(tcl, DARK_PRIMARY, font_size=10)

    _add_card(slide, 0.65, 5.55, 12.3, 1.35, LIGHT_PANEL, LIGHT_BORDER)
    cat = add_placeholder_box(slide, 0.85, 5.65, 11.9, 0.4, "CVE_ANALYSIS_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(cat, LIGHT_PRIMARY, font_size=14, bold=True)
    ana = add_placeholder_box(slide, 0.85, 6.05, 11.9, 0.75, "TOP_CVE_ANALYSIS", font_size=10, color=DARK_PRIMARY)
    _style_textbox(ana, DARK_PRIMARY, font_size=10)

    exs = add_placeholder_box(slide, 0.65, 6.95, 12.3, 0.45, "EXPOSURE_SUMMARY", font_size=9, color=LIGHT_MUTED)
    _style_textbox(exs, LIGHT_MUTED, font_size=9)

    # Slide 6: Cloud Security
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 3.65, 1.35, LIGHT_PANEL, LIGHT_BORDER)
    cat = add_placeholder_box(slide, 0.85, 1.15, 3.25, 0.35, "CLOUD_ACCOUNTS_TITLE", font_size=12, color=LIGHT_MUTED)
    _style_textbox(cat, LIGHT_MUTED, font_size=12)
    cac = add_placeholder_box(slide, 0.85, 1.45, 3.25, 0.85, "CLOUD_ACCOUNTS_COUNT", font_size=26, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(cac, LIGHT_PRIMARY, font_size=26, bold=True)

    _add_card(slide, 4.45, 1.05, 8.5, 3.55, LIGHT_PANEL, LIGHT_BORDER)
    crt = add_placeholder_box(slide, 4.65, 1.15, 8.1, 0.4, "CLOUD_RISK_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(crt, LIGHT_PRIMARY, font_size=14, bold=True)
    crl = add_placeholder_box(slide, 4.65, 1.55, 3.9, 2.95, "CLOUD_RISK_LIST", font_size=11, color=DARK_PRIMARY)
    _style_textbox(crl, DARK_PRIMARY, font_size=11)
    cat2 = add_placeholder_box(slide, 8.65, 1.15, 4.1, 0.4, "CLOUD_ANALYSIS_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(cat2, LIGHT_PRIMARY, font_size=14, bold=True)
    crs = add_placeholder_box(slide, 8.65, 1.55, 4.1, 2.95, "CLOUD_RISK_SUMMARY", font_size=11, color=DARK_PRIMARY)
    _style_textbox(crs, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 4.75, 12.3, 2.5, LIGHT_PANEL, LIGHT_BORDER)
    crt2 = add_placeholder_box(slide, 0.85, 4.85, 11.9, 0.4, "CLOUD_REC_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(crt2, LIGHT_PRIMARY, font_size=14, bold=True)
    rec = add_placeholder_box(slide, 0.85, 5.25, 11.9, 1.95, "CLOUD_RECOMMENDATIONS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(rec, DARK_PRIMARY, font_size=11)

    # Slide 7: Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 6.05, 3.25, LIGHT_PANEL, LIGHT_BORDER)
    p0t = add_placeholder_box(slide, 0.85, 1.15, 5.65, 0.4, "P0_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(p0t, LIGHT_PRIMARY, font_size=14, bold=True)
    p0a = add_placeholder_box(slide, 0.85, 1.55, 5.65, 2.65, "P0_ACTIONS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(p0a, DARK_PRIMARY, font_size=11)

    _add_card(slide, 6.9, 1.05, 6.05, 3.25, LIGHT_PANEL, LIGHT_BORDER)
    p1t = add_placeholder_box(slide, 7.1, 1.15, 5.65, 0.4, "P1_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(p1t, LIGHT_PRIMARY, font_size=14, bold=True)
    p1a = add_placeholder_box(slide, 7.1, 1.55, 5.65, 2.65, "P1_ACTIONS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(p1a, DARK_PRIMARY, font_size=11)

    _add_card(slide, 0.65, 4.45, 12.3, 2.8, LIGHT_PANEL, LIGHT_BORDER)
    srt = add_placeholder_box(slide, 0.85, 4.55, 11.9, 0.4, "STRATEGIC_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(srt, LIGHT_PRIMARY, font_size=14, bold=True)
    srr = add_placeholder_box(slide, 0.85, 4.95, 11.9, 2.2, "STRATEGIC_RECOMMENDATIONS", font_size=11, color=DARK_PRIMARY)
    _style_textbox(srr, DARK_PRIMARY, font_size=11)

    # Slide 8: Appendix
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _light_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(st, LIGHT_PRIMARY, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 6.05, 2.05, LIGHT_PANEL, LIGHT_BORDER)
    dst = add_placeholder_box(slide, 0.85, 1.15, 5.65, 0.4, "DATA_SCOPE_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(dst, LIGHT_PRIMARY, font_size=14, bold=True)
    ds = add_placeholder_box(slide, 0.85, 1.55, 5.65, 1.45, "DATA_SCOPE", font_size=10, color=DARK_PRIMARY)
    _style_textbox(ds, DARK_PRIMARY, font_size=10)

    _add_card(slide, 6.9, 1.05, 6.05, 2.05, LIGHT_PANEL, LIGHT_BORDER)
    act = add_placeholder_box(slide, 7.1, 1.15, 5.65, 0.4, "ASSET_COVERAGE_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(act, LIGHT_PRIMARY, font_size=14, bold=True)
    ac = add_placeholder_box(slide, 7.1, 1.55, 5.65, 1.45, "ASSET_COVERAGE", font_size=10, color=DARK_PRIMARY)
    _style_textbox(ac, DARK_PRIMARY, font_size=10)

    _add_card(slide, 0.65, 3.25, 12.3, 0.95, LIGHT_PANEL, LIGHT_BORDER)
    slat = add_placeholder_box(slide, 0.85, 3.35, 3.0, 0.35, "SLA_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(slat, LIGHT_PRIMARY, font_size=14, bold=True)
    slan = add_placeholder_box(slide, 3.85, 3.35, 9.0, 0.55, "SLA_NOTES", font_size=10, color=DARK_PRIMARY)
    _style_textbox(slan, DARK_PRIMARY, font_size=10)

    _add_card(slide, 0.65, 4.35, 12.3, 2.55, LIGHT_PANEL, LIGHT_BORDER)
    tt = add_placeholder_box(slide, 0.85, 4.45, 11.9, 0.4, "TERMINOLOGY_TITLE", font_size=14, bold=True, color=LIGHT_PRIMARY)
    _style_textbox(tt, LIGHT_PRIMARY, font_size=14, bold=True)
    term = add_placeholder_box(slide, 0.85, 4.85, 11.9, 1.95, "TERMINOLOGY", font_size=10, color=DARK_PRIMARY)
    _style_textbox(term, DARK_PRIMARY, font_size=10)

    ci = add_placeholder_box(slide, 0.65, 6.95, 12.3, 0.45, "CONTACT_INFO", font_size=10, color=LIGHT_MUTED)
    _style_textbox(ci, LIGHT_MUTED, font_size=10)

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
    _dark_base(slide)
    _add_rect(slide, 0, 0, 0.18, 7.5, DARK_ACCENT, transparency=0.0)
    _add_card(slide, 0.65, 2.0, 12.3, 2.35, DARK_PANEL, DARK_BORDER)
    rt = add_placeholder_box(slide, 0.85, 2.2, 11.9, 0.95, "REPORT_TITLE", font_size=34, bold=True, color=DARK_ACCENT)
    _style_textbox(rt, DARK_ACCENT, font_size=34, bold=True)

    _add_card(slide, 0.65, 4.55, 7.8, 2.25, DARK_PANEL, DARK_BORDER, shadow=False)
    cl = add_placeholder_box(slide, 0.85, 4.72, 7.4, 0.55, "CUSTOMER_LABEL", font_size=18, bold=True, color=DARK_TEXT)
    _style_textbox(cl, DARK_TEXT, font_size=18, bold=True)
    pl = add_placeholder_box(slide, 0.85, 5.33, 7.4, 0.45, "PERIOD_LABEL", font_size=14, color=DARK_MUTED)
    _style_textbox(pl, DARK_MUTED, font_size=14)
    vt = add_placeholder_box(slide, 0.85, 5.83, 7.4, 0.45, "VERSION_TAG", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(vt, DARK_ACCENT, font_size=12, bold=True)
    cf = add_placeholder_box(slide, 0.85, 6.28, 7.4, 0.45, "CONFIDENTIALITY_LABEL", font_size=12, color=DARK_MUTED)
    _style_textbox(cf, DARK_MUTED, font_size=12)
    ga = add_placeholder_box(slide, 0.85, 6.73, 7.4, 0.45, "GENERATED_AT_LABEL", font_size=10, color=DARK_MUTED)
    _style_textbox(ga, DARK_MUTED, font_size=10)

    # Slide 2: Security Dashboard
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 3.65, DARK_PANEL, DARK_BORDER)
    kpi = add_placeholder_box(slide, 0.85, 1.18, 11.9, 3.4, "KPI_DASHBOARD", font_size=11, color=DARK_TEXT)
    _style_textbox(kpi, DARK_TEXT, font_size=11)

    _add_card(slide, 0.65, 4.85, 12.3, 2.35, DARK_PANEL, DARK_BORDER)
    at = add_placeholder_box(slide, 0.85, 4.95, 11.9, 0.4, "ASSESSMENT_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(at, DARK_ACCENT, font_size=14, bold=True)
    ta = add_placeholder_box(slide, 0.85, 5.35, 11.9, 1.75, "TECHNICAL_ASSESSMENT", font_size=11, color=DARK_TEXT)
    _style_textbox(ta, DARK_TEXT, font_size=11)

    # Slide 3: Alert Deep Analysis
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 1.35, DARK_PANEL, DARK_BORDER)
    svt = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "SEVERITY_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(svt, DARK_ACCENT, font_size=14, bold=True)
    sv = add_placeholder_box(slide, 0.85, 1.55, 11.9, 0.75, "SEVERITY_BREAKDOWN", font_size=11, color=DARK_TEXT)
    _style_textbox(sv, DARK_TEXT, font_size=11)

    _add_card(slide, 0.65, 2.55, 6.05, 2.2, DARK_PANEL, DARK_BORDER)
    trt = add_placeholder_box(slide, 0.85, 2.65, 5.65, 0.4, "TOP_RULES_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(trt, DARK_ACCENT, font_size=14, bold=True)
    trtbl = add_placeholder_box(slide, 0.85, 3.05, 5.65, 1.6, "TOP_RULES_TABLE", font_size=9, color=DARK_TEXT)
    _style_textbox(trtbl, DARK_TEXT, font_size=9)

    _add_card(slide, 6.9, 2.55, 6.05, 2.2, DARK_PANEL, DARK_BORDER)
    rat = add_placeholder_box(slide, 7.1, 2.65, 5.65, 0.4, "RULES_ANALYSIS_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(rat, DARK_ACCENT, font_size=14, bold=True)
    raa = add_placeholder_box(slide, 7.1, 3.05, 5.65, 1.6, "TOP_RULES_ANALYSIS", font_size=10, color=DARK_TEXT)
    _style_textbox(raa, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 4.9, 12.3, 2.3, DARK_PANEL, DARK_BORDER)
    fpt = add_placeholder_box(slide, 0.85, 5.0, 11.9, 0.4, "FP_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(fpt, DARK_ACCENT, font_size=14, bold=True)
    fpi = add_placeholder_box(slide, 0.85, 5.4, 11.9, 1.7, "FALSE_POSITIVE_INSIGHT", font_size=11, color=DARK_TEXT)
    _style_textbox(fpi, DARK_TEXT, font_size=11)

    # Slide 4: Incident Timeline
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 2.95, DARK_PANEL, DARK_BORDER)
    tt = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "TIMELINE_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(tt, DARK_ACCENT, font_size=14, bold=True)
    tn = add_placeholder_box(slide, 0.85, 1.55, 11.9, 2.35, "TIMELINE_NARRATIVE", font_size=11, color=DARK_TEXT)
    _style_textbox(tn, DARK_TEXT, font_size=11)

    _add_card(slide, 0.65, 4.15, 6.05, 3.05, DARK_PANEL, DARK_BORDER)
    rt = add_placeholder_box(slide, 0.85, 4.25, 5.65, 0.4, "RESPONSE_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(rt, DARK_ACCENT, font_size=14, bold=True)
    rm = add_placeholder_box(slide, 0.85, 4.65, 5.65, 2.45, "RESPONSE_METRICS", font_size=11, color=DARK_TEXT)
    _style_textbox(rm, DARK_TEXT, font_size=11)

    _add_card(slide, 6.9, 4.15, 6.05, 3.05, DARK_PANEL, DARK_BORDER)
    lt = add_placeholder_box(slide, 7.1, 4.25, 5.65, 0.4, "LESSONS_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(lt, DARK_ACCENT, font_size=14, bold=True)
    ll = add_placeholder_box(slide, 7.1, 4.65, 5.65, 2.45, "LESSONS_LEARNED", font_size=11, color=DARK_TEXT)
    _style_textbox(ll, DARK_TEXT, font_size=11)

    # Slide 5: Incident Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 2.65, DARK_PANEL, DARK_BORDER)
    i1t = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "INCIDENT1_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(i1t, DARK_ACCENT, font_size=14, bold=True)
    i1d = add_placeholder_box(slide, 0.85, 1.55, 11.9, 2.05, "INCIDENT_DETAIL_1", font_size=10, color=DARK_TEXT)
    _style_textbox(i1d, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 3.85, 12.3, 2.65, DARK_PANEL, DARK_BORDER)
    i2t = add_placeholder_box(slide, 0.85, 3.95, 11.9, 0.4, "INCIDENT2_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(i2t, DARK_ACCENT, font_size=14, bold=True)
    i2d = add_placeholder_box(slide, 0.85, 4.35, 11.9, 2.05, "INCIDENT_DETAIL_2", font_size=10, color=DARK_TEXT)
    _style_textbox(i2d, DARK_TEXT, font_size=10)

    aps = add_placeholder_box(slide, 0.65, 6.95, 12.3, 0.45, "ATTACK_PATTERN_SUMMARY", font_size=9, color=DARK_MUTED)
    _style_textbox(aps, DARK_MUTED, font_size=9)

    # Slide 6: Vulnerability Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 2.15, DARK_PANEL, DARK_BORDER)
    vdt = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "VULN_DIST_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(vdt, DARK_ACCENT, font_size=14, bold=True)
    vd = add_placeholder_box(slide, 0.85, 1.55, 11.9, 1.55, "VULN_DISTRIBUTION", font_size=11, color=DARK_TEXT)
    _style_textbox(vd, DARK_TEXT, font_size=11)

    _add_card(slide, 0.65, 3.35, 12.3, 2.35, DARK_PANEL, DARK_BORDER)
    cat = add_placeholder_box(slide, 0.85, 3.45, 11.9, 0.4, "CVE_ANALYSIS_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(cat, DARK_ACCENT, font_size=14, bold=True)
    cta = add_placeholder_box(slide, 0.85, 3.85, 11.9, 1.75, "CVE_TECHNICAL_ANALYSIS", font_size=10, color=DARK_TEXT)
    _style_textbox(cta, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 5.85, 12.3, 1.35, DARK_PANEL, DARK_BORDER)
    pt = add_placeholder_box(slide, 0.85, 5.95, 11.9, 0.4, "PATCH_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(pt, DARK_ACCENT, font_size=14, bold=True)
    ps = add_placeholder_box(slide, 0.85, 6.35, 11.9, 0.75, "PATCH_STATUS", font_size=10, color=DARK_TEXT)
    _style_textbox(ps, DARK_TEXT, font_size=10)

    # Slide 7: Attack Surface
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 6.05, 2.1, DARK_PANEL, DARK_BORDER)
    est = add_placeholder_box(slide, 0.85, 1.15, 5.65, 0.4, "EXPOSED_SERVICES_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(est, DARK_ACCENT, font_size=14, bold=True)
    estbl = add_placeholder_box(slide, 0.85, 1.55, 5.65, 1.5, "EXPOSED_SERVICES_TABLE", font_size=9, color=DARK_TEXT)
    _style_textbox(estbl, DARK_TEXT, font_size=9)

    _add_card(slide, 6.9, 1.05, 6.05, 2.1, DARK_PANEL, DARK_BORDER)
    eat = add_placeholder_box(slide, 7.1, 1.15, 5.65, 0.4, "EXPOSURE_ANALYSIS_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(eat, DARK_ACCENT, font_size=14, bold=True)
    ee = add_placeholder_box(slide, 7.1, 1.55, 5.65, 1.5, "EXTERNAL_EXPOSURE", font_size=10, color=DARK_TEXT)
    _style_textbox(ee, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 3.3, 6.05, 3.9, DARK_PANEL, DARK_BORDER)
    stt = add_placeholder_box(slide, 0.85, 3.4, 5.65, 0.4, "SURFACE_TREND_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(stt, DARK_ACCENT, font_size=14, bold=True)
    ast = add_placeholder_box(slide, 0.85, 3.8, 5.65, 3.3, "ATTACK_SURFACE_TREND", font_size=11, color=DARK_TEXT)
    _style_textbox(ast, DARK_TEXT, font_size=11)

    _add_card(slide, 6.9, 3.3, 6.05, 3.9, DARK_PANEL, DARK_BORDER)
    rt = add_placeholder_box(slide, 7.1, 3.4, 5.65, 0.4, "REMEDIATION_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(rt, DARK_ACCENT, font_size=14, bold=True)
    er = add_placeholder_box(slide, 7.1, 3.8, 5.65, 3.3, "EXPOSURE_REMEDIATION", font_size=10, color=DARK_TEXT)
    _style_textbox(er, DARK_TEXT, font_size=10)

    # Slide 8: Cloud Security Details
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 2.7, DARK_PANEL, DARK_BORDER)
    cmt = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "CLOUD_MISCONFIG_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(cmt, DARK_ACCENT, font_size=14, bold=True)
    cmd = add_placeholder_box(slide, 0.85, 1.55, 11.9, 2.05, "CLOUD_MISCONFIG_DETAILS", font_size=10, color=DARK_TEXT)
    _style_textbox(cmd, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 3.95, 6.05, 3.25, DARK_PANEL, DARK_BORDER)
    it = add_placeholder_box(slide, 0.85, 4.05, 5.65, 0.4, "IAM_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(it, DARK_ACCENT, font_size=14, bold=True)
    ira = add_placeholder_box(slide, 0.85, 4.45, 5.65, 2.65, "IAM_RISK_ANALYSIS", font_size=10, color=DARK_TEXT)
    _style_textbox(ira, DARK_TEXT, font_size=10)

    _add_card(slide, 6.9, 3.95, 6.05, 3.25, DARK_PANEL, DARK_BORDER)
    ct = add_placeholder_box(slide, 7.1, 4.05, 5.65, 0.4, "COMPLIANCE_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(ct, DARK_ACCENT, font_size=14, bold=True)
    cc = add_placeholder_box(slide, 7.1, 4.45, 5.65, 2.65, "CLOUD_COMPLIANCE", font_size=10, color=DARK_TEXT)
    _style_textbox(cc, DARK_TEXT, font_size=10)

    # Slide 9: Technical Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 12.3, 2.45, DARK_PANEL, DARK_BORDER)
    trt = add_placeholder_box(slide, 0.85, 1.15, 11.9, 0.4, "TECH_REC_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(trt, DARK_ACCENT, font_size=14, bold=True)
    tr = add_placeholder_box(slide, 0.85, 1.55, 11.9, 1.85, "TECHNICAL_RECOMMENDATIONS", font_size=10, color=DARK_TEXT)
    _style_textbox(tr, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 3.65, 6.05, 3.55, DARK_PANEL, DARK_BORDER)
    dt = add_placeholder_box(slide, 0.85, 3.75, 5.65, 0.4, "DETECTION_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(dt, DARK_ACCENT, font_size=14, bold=True)
    di = add_placeholder_box(slide, 0.85, 4.15, 5.65, 2.95, "DETECTION_IMPROVEMENTS", font_size=10, color=DARK_TEXT)
    _style_textbox(di, DARK_TEXT, font_size=10)

    _add_card(slide, 6.9, 3.65, 6.05, 3.55, DARK_PANEL, DARK_BORDER)
    ht = add_placeholder_box(slide, 7.1, 3.75, 5.65, 0.4, "HARDENING_TITLE", font_size=14, bold=True, color=DARK_ACCENT)
    _style_textbox(ht, DARK_ACCENT, font_size=14, bold=True)
    hc = add_placeholder_box(slide, 7.1, 4.15, 5.65, 2.95, "HARDENING_CHECKLIST", font_size=10, color=DARK_TEXT)
    _style_textbox(hc, DARK_TEXT, font_size=10)

    # Slide 10: Appendix
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _dark_base(slide)
    st = add_placeholder_box(slide, 0.65, 0.22, 12.3, 0.68, "SLIDE_TITLE", font_size=28, bold=True, color=DARK_ACCENT)
    _style_textbox(st, DARK_TEXT, font_size=28, bold=True)

    _add_card(slide, 0.65, 1.05, 6.05, 1.2, DARK_PANEL, DARK_BORDER)
    et = add_placeholder_box(slide, 0.85, 1.15, 5.65, 0.35, "EVIDENCE_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(et, DARK_ACCENT, font_size=12, bold=True)
    ei = add_placeholder_box(slide, 0.85, 1.45, 5.65, 0.75, "EVIDENCE_INDEX", font_size=9, color=DARK_TEXT)
    _style_textbox(ei, DARK_TEXT, font_size=9)

    _add_card(slide, 6.9, 1.05, 6.05, 1.2, DARK_PANEL, DARK_BORDER)
    ft = add_placeholder_box(slide, 7.1, 1.15, 5.65, 0.35, "FRESHNESS_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(ft, DARK_ACCENT, font_size=12, bold=True)
    df = add_placeholder_box(slide, 7.1, 1.45, 5.65, 0.75, "DATA_FRESHNESS", font_size=10, color=DARK_TEXT)
    _style_textbox(df, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 2.35, 12.3, 0.85, DARK_PANEL, DARK_BORDER)
    dst = add_placeholder_box(slide, 0.85, 2.45, 3.4, 0.35, "DATA_SCOPE_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(dst, DARK_ACCENT, font_size=12, bold=True)
    dsn = add_placeholder_box(slide, 4.25, 2.45, 8.7, 0.55, "DATA_SCOPE_NOTES", font_size=10, color=DARK_TEXT)
    _style_textbox(dsn, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 3.3, 6.05, 1.35, DARK_PANEL, DARK_BORDER)
    vnt = add_placeholder_box(slide, 0.85, 3.4, 5.65, 0.35, "VULN_NOTES_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(vnt, DARK_ACCENT, font_size=12, bold=True)
    vn = add_placeholder_box(slide, 0.85, 3.7, 5.65, 0.85, "VULN_NOTES", font_size=10, color=DARK_TEXT)
    _style_textbox(vn, DARK_TEXT, font_size=10)

    _add_card(slide, 6.9, 3.3, 6.05, 1.35, DARK_PANEL, DARK_BORDER)
    ent = add_placeholder_box(slide, 7.1, 3.4, 5.65, 0.35, "EVIDENCE_NOTES_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(ent, DARK_ACCENT, font_size=12, bold=True)
    en = add_placeholder_box(slide, 7.1, 3.7, 5.65, 0.85, "EVIDENCE_NOTES", font_size=10, color=DARK_TEXT)
    _style_textbox(en, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 4.75, 6.05, 1.15, DARK_PANEL, DARK_BORDER)
    at = add_placeholder_box(slide, 0.85, 4.85, 5.65, 0.35, "ASSETS_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(at, DARK_ACCENT, font_size=12, bold=True)
    ag = add_placeholder_box(slide, 0.85, 5.15, 5.65, 0.65, "ASSET_GROUPS_LIST", font_size=10, color=DARK_TEXT)
    _style_textbox(ag, DARK_TEXT, font_size=10)

    _add_card(slide, 6.9, 4.75, 6.05, 1.15, DARK_PANEL, DARK_BORDER)
    svt = add_placeholder_box(slide, 7.1, 4.85, 5.65, 0.35, "SERVICES_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(svt, DARK_ACCENT, font_size=12, bold=True)
    ksl = add_placeholder_box(slide, 7.1, 5.15, 5.65, 0.65, "KEY_SERVICES_LIST", font_size=10, color=DARK_TEXT)
    _style_textbox(ksl, DARK_TEXT, font_size=10)

    _add_card(slide, 0.65, 6.0, 12.3, 1.45, DARK_PANEL, DARK_BORDER)
    termt = add_placeholder_box(slide, 0.85, 6.1, 11.9, 0.35, "TERMINOLOGY_TITLE", font_size=12, bold=True, color=DARK_ACCENT)
    _style_textbox(termt, DARK_ACCENT, font_size=12, bold=True)
    term = add_placeholder_box(slide, 0.85, 6.4, 11.9, 1.0, "TERMINOLOGY_TECHNICAL", font_size=9, color=DARK_TEXT)
    _style_textbox(term, DARK_TEXT, font_size=9)

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

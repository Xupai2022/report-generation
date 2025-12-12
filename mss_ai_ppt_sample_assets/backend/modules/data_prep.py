from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List

from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput
from mss_ai_ppt_sample_assets.backend.models.templates import TemplateDescriptor


@dataclass
class DataPrepResult:
    """Result of data preparation for V1 templates."""
    slide_inputs: Dict[str, Dict[str, Any]]
    facts: Dict[str, Any]


# ============================================================================
# V2 Data Preparation - Minimal, just passes raw data
# ============================================================================

def prepare_data_v2(tenant_input: TenantInput) -> TenantInput:
    """Prepare data for V2 templates.

    For V2, we pass the raw TenantInput directly to the LLM.
    No pre-processing of content is needed - the AI generates all text.

    Args:
        tenant_input: Raw tenant input data

    Returns:
        The same TenantInput (no transformation needed for V2)
    """
    # V2 doesn't need any data preparation - the raw input goes directly to LLM
    # The LLMOrchestratorV2 handles all data extraction and AI generation
    return tenant_input


# ============================================================================
# V1 Data Preparation - Legacy (kept for backward compatibility)
# ============================================================================

def _format_period(tenant_input: TenantInput) -> str:
    period = tenant_input.get("period", {})
    start = period.get("start", "")
    end = period.get("end", "")
    if start and end:
        return f"{start} ~ {end}"
    return period.get("label") or ""


def _safe_percent(value: float) -> str:
    try:
        return f"{round(value * 100)}%"
    except Exception:
        return str(value)


def _build_cover(tenant_input: TenantInput) -> Dict[str, Any]:
    tenant = tenant_input.get("tenant", {})
    report_meta = tenant_input.get("report_meta", {})
    return {
        "report_title": report_meta.get("report_title", "MSS 报告"),
        "customer_name": tenant.get("name", tenant.get("id", "")),
        "period": _format_period(tenant_input),
        "confidentiality": report_meta.get("secrecy_default", "INTERNAL").upper(),
        "generated_at": report_meta.get("generated_at", ""),
    }


def _build_management_summary(tenant_input: TenantInput) -> Dict[str, Any]:
    alerts = tenant_input.get("alerts", {})
    incidents = tenant_input.get("incidents", []) or []
    mss_ops = tenant_input.get("mss_ops", {})

    alerts_total = alerts.get("total", 0)
    incidents_high = len([i for i in incidents if i.get("severity") == "high"])
    mttr_hours = mss_ops.get("mttr_hours_avg", 0)
    false_positive_rate = alerts.get("false_positive_rate", 0)

    bullets: List[str] = [
        f"本期告警总量 {alerts_total}，高危告警 {alerts.get('by_severity', {}).get('high', 0)}。",
        f"高危事件 {incidents_high} 起，平均MTTR {mttr_hours} 小时。",
        "Top 类别："
        + "/".join([c.get("category") for c in (alerts.get("top_categories") or [])[:3]]),
    ]

    cloud_risks = tenant_input.get("cloud_risks", [])
    if cloud_risks:
        bullets.append(
            f"云侧暴露面：外部IP {tenant_input.get('exposure', {}).get('external_ips', 0)}，新暴露 {tenant_input.get('exposure', {}).get('new_exposed_since_last_period', 0)}。"
        )

    bullets_text = "\n".join(f"• {b}" for b in bullets)
    return {
        "title": "本期态势概览（管理层）",
        "kpis": {
            "alerts_total": str(alerts_total),
            "incidents_high": str(incidents_high),
            "mttr_hours": str(mttr_hours),
            "false_positive_rate": _safe_percent(false_positive_rate),
        },
        "bullets": bullets,
        "bullets_text": bullets_text,
        "footer": f"{_build_cover(tenant_input)['customer_name']} | MSS",
    }


def _build_alerts_overview(tenant_input: TenantInput) -> Dict[str, Any]:
    alerts = tenant_input.get("alerts", {})
    trend = alerts.get("trend_weekly", {})
    labels = trend.get("labels", [])
    values = trend.get("values", [])
    chart_placeholder_text = f"趋势占位：{labels} -> {values}"

    top_categories = alerts.get("top_categories", []) or []
    top_cats_lines = [
        f"• {item.get('category')}: {item.get('count')} (MoM {item.get('mom_change_pct')}%)"
        for item in top_categories[:3]
    ]
    top_categories_text = "Top 类别\n" + "\n".join(top_cats_lines) if top_cats_lines else ""
    footer = alerts.get("notes", "")

    return {
        "title": "告警趋势与Top类别",
        "chart_placeholder_text": chart_placeholder_text,
        "top_categories": top_cats_lines,
        "top_categories_text": top_categories_text,
        "footer": footer,
    }


def _build_major_incidents(tenant_input: TenantInput) -> Dict[str, Any]:
    incidents = tenant_input.get("incidents", []) or []
    incidents_sorted = sorted(
        incidents, key=lambda i: i.get("severity", ""), reverse=True
    )[:4]
    lines = []
    severity_order = {"critical": 4, "high": 3, "medium": 2, "low": 1}
    incidents_sorted = sorted(
        incidents_sorted, key=lambda i: severity_order.get(i.get("severity"), 0), reverse=True
    )
    for inc in incidents_sorted:
        sev = inc.get("severity", "").upper()
        lines.append(
            f"{sev}: {inc.get('title')} | MTTD {inc.get('mttd_minutes')}m | MTTR {inc.get('mttr_minutes')}m | 资产: {', '.join(inc.get('impacted_assets', []))}"
        )
    incident_list_text = "本期重点事件\n" + "\n".join(f"• {l}" for l in lines) if lines else ""

    return {
        "title": "重点事件摘要",
        "incidents": incidents_sorted,
        "incident_list_text": incident_list_text,
        "footer": "仅供内部使用",
    }


def _build_recommendations(tenant_input: TenantInput) -> Dict[str, Any]:
    cloud_risks = tenant_input.get("cloud_risks", [])

    p0_actions = [
        "对高权限账号强制MFA并定期轮换密钥",
        "云安全组0.0.0.0/0暴露自动告警与收敛",
    ]
    if cloud_risks:
        p0_actions.append("针对暴露服务启用WAF/速率限制与地理封禁")

    p1_actions = [
        "完善资产基线与补丁节奏，关键组件月度复盘",
        "建立高危告警误报治理与规则版本评审机制",
    ]

    p0_text = "\n".join(f"• {a}" for a in p0_actions)
    p1_text = "\n".join(f"• {a}" for a in p1_actions)

    return {
        "title": "整改建议（P0/P1）",
        "p0_actions": p0_actions,
        "p1_actions": p1_actions,
        "p0_text": p0_text,
        "p1_text": p1_text,
        "footer": "建议可按业务线拆解落地",
    }


def prepare_facts_for_template(
    tenant_input: TenantInput, template: TemplateDescriptor
) -> DataPrepResult:
    """Prepare deterministic facts from tenant input for V1 templates (legacy).

    Note: For V2 templates, use prepare_data_v2() instead - it simply returns
    the raw data for the AI to process.
    """
    slide_inputs: Dict[str, Dict[str, Any]] = {}

    # Common slide data for known templates
    if template.template_id == "mss_management_light_v1":
        slide_inputs["cover"] = _build_cover(tenant_input)
        slide_inputs["executive_summary"] = _build_management_summary(tenant_input)
        slide_inputs["alerts_overview"] = _build_alerts_overview(tenant_input)
        slide_inputs["major_incidents"] = _build_major_incidents(tenant_input)
        slide_inputs["recommendations"] = _build_recommendations(tenant_input)

    elif template.template_id == "mss_technical_dark_v1":
        # For now reuse some generic pieces; can be specialised later.
        slide_inputs["cover"] = _build_cover(tenant_input)
        slide_inputs["alerts_detail"] = {
            "title": "告警详情",
            "alerts_detail_text": "Top 规则与误报说明占位",
            "false_positive_notes": "误报率待模型生成",
            "footer": "技术版",
        }
        slide_inputs["incident_timeline"] = {
            "title": "事件时间线",
            "incident_timeline_text": "时间线占位",
            "footer": "技术版",
        }
        slide_inputs["vulnerability_exposure"] = {
            "title": "漏洞暴露",
            "vuln_summary_text": "漏洞摘要占位",
            "top_vulns_text": "Top CVE 占位",
            "footer": "技术版",
        }
        slide_inputs["cloud_attack_surface"] = {
            "title": "云攻击面",
            "cloud_risks_text": "云风险占位",
            "exposure_findings_text": "暴露项占位",
            "footer": "技术版",
        }
        slide_inputs["appendix_evidence"] = {
            "title": "证据索引",
            "evidence_index_text": "Query ID 占位",
            "footer": "技术版",
        }

    facts: Dict[str, Any] = {
        "alerts_total": tenant_input.get("alerts", {}).get("total"),
        "incidents_high": len(
            [i for i in tenant_input.get("incidents", []) if i.get("severity") == "high"]
        ),
        "mttr_hours_avg": tenant_input.get("mss_ops", {}).get("mttr_hours_avg"),
        "false_positive_rate": tenant_input.get("alerts", {}).get("false_positive_rate"),
        "top_categories": [c.get("category") for c in tenant_input.get("alerts", {}).get("top_categories", [])],
    }

    return DataPrepResult(slide_inputs=slide_inputs, facts=facts)

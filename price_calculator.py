import argparse
import re
from decimal import Decimal, ROUND_HALF_UP


INPUT_TEXT = """## System Prompt
text
你是一名专业的 MSS 安全报告写作专家，职责是根据客户的安全数据进行分析，根据指定的PPT模板撰写报告。

## 写作目标
- 基于真实数据为 PPT 页面产出基于证据、富有洞察的内容。
- 确保表述风格与管理层受众匹配。

## 输出约束
1. 所有表述必须基于输入数据，保持事实准确。
2. 禁止输出无依据的判断或结论。
3. 在合适场景下优先给出可执行、可落地的建议。
4. 语言清晰、专业，满足业务汇报语境。
5. 仅返回合法 JSON（不要使用 markdown 包裹）。
## User Prompt
text
## 任务
请为指定页面和占位符生成 AI 内容。
输出必须是合法 JSON，且只能输出 JSON。

## 写作硬约束
1) 严禁空话和套话（如“持续提升”“稳步推进”）单独成句。
2) 每条结论至少包含“数据依据 + 判断”，优先补充“业务影响或行动建议”。
3) 不得编造数据，不得输出与输入数据冲突的结论。

## 输出格式
json { "slides": [ {"slide_key": "incident_effectiveness", "placeholders": {"business_continuity_assurance_conclusion": "...", "user_trust_assurance_conclusion": "..."}}, {"slide_key": "asset_management", "placeholders": {"asset_improve1": "...", "asset_improve2": "...", "asset_improve3": "...", "asset_summary": "...", "asset_value1": "...", "asset_value2": "...", "asset_value3": "..."}}, {"slide_key": "vulnerability_effectiveness", "placeholders": {"vuln_summary": "..."}}, {"slide_key": "threat_effectiveness", "placeholders": {"threat_improve1": "...", "threat_improve2": "...", "threat_improve3": "...", "threat_summary": "..."}}, {"slide_key": "critical_assurance", "placeholders": {"duty_summary": "...", "security_value1": "...", "security_value2": "...", "security_value3": "..."}}, {"slide_key": "platform_effectiveness", "placeholders": {"platform_improve1": "...", "platform_improve2": "...", "platform_improve3": "..."}} ] }
## 偏好重点指引
请根据用户已选偏好控制内容重点与表达风格。
- 业务保护: 定义为：围绕关键资产，明确保护对象；消除脆弱性，减少被攻击面；抵御威胁，防止业务被破坏；快速处置安全事件，保障业务持续不中断。保护业务不中断、数据不失控、运行可持续。此用户选择偏好强调业务连续性、风险遏制和可执行防护结果。
### 页面：安全事件管理成效 (incident_effectiveness)

**business_continuity_assurance_conclusion**
请基于本页数据（incident_effectiveness）输出“安全事件管理对业务连续性保障成效”的结论，且必须与用户已选重点业务保护直接关联。写作要求：1）不要泛化空话，必须引用至少2个关键数据（如事件总量、平均响应时长、平均处置时长、闭环率、事件等级分布/趋势）；2）必须同时覆盖三层信息：风险影响程度（高危/中危占比或分布特征）、处置效率（响应与处置时效）、治理结果（闭环与风险遏制效果）；3）在给出结果后补一句“业务连续性价值”，明确说明对核心业务可用性/中断风险/恢复能力的实际影响；4）可做合理归纳，但严禁编造未提供的数据或结论；5）避免仅罗列指标，需形成“数据依据→管理成效→业务连续性价值”的完整因果表达。


**user_trust_assurance_conclusion**
基于以下内容及业务连续性结论、结合本页数据（incident_effectiveness），输出用户依赖度目标达成情况结论：1.资产维度：核心数据、业务流程、账号权限高度集中在某一资产 / 服务商，资产不可迁移、不可导出。2.脆弱性维度：因过度依赖形成单点脆弱性，架构、接口、数据格式被锁定，替代成本极高。3.威胁维度：厂商停服、断供、涨价、漏洞、合规风险、供应链风险等，均直接构成业务威胁。4.事件维度：一旦发生中断、入侵、违约等安全事件，业务将直接受冲击，且无快速替代方案。

### 页面：资产安全管理成效 (asset_management)

**asset_improve1** 
基于本页资产管理数据（asset_management）生成第1条改进措施，梳理在本季/年度工作中，因资产管理工作执行差距产生的安全隐患，并给出可落地建议。


**asset_improve2**
基于本页资产管理数据（asset_management）生成第2条改进措施，梳理在本季/年度工作中，因资产管理工作执行差距产生的安全隐患，并给出可落地建议。


**asset_improve3**
基于本页资产管理数据（asset_management）生成第3条改进措施，梳理在本季/年度工作中，因资产管理工作执行差距产生的安全隐患，并给出可落地建议。


**asset_summary**
基于围绕资产产生的事件、风险、存在的暴露端口及其他相关信息，给出本季/年度在资产安全管理工作中的重点结论，突出资产管理举措的有效性和安全价值


**asset_value1** 
基于本页资产管理数据（asset_management），输出第一个简短短句（6-8个字）展示本季/年度在资产管理工作中的重点价值。


**asset_value2**
基于本页资产管理数据（asset_management），输出第二个简短短句（6-8个字）展示本季/年度在资产管理工作中的重点价值。


**asset_value3**
基于本页资产管理数据（asset_management），输出第三个简短短句（6-8个字）展示本季/年度在资产管理工作中的重点价值。

### 页面：脆弱性管理成效 (vulnerability_effectiveness)

**vuln_summary** 
你要生成的是该句“后半句续写”，前置固定文案已在PPT中写好：`全年开展{{scanning}} 次漏洞扫描，`。请不要重复或改写这段前置文案，不要再次输出扫描次数，不要以“全年开展”开头。仅从逗号后继续写总结：根据已选业务保护，基于 vulnerability_effectiveness 数据，围绕已发现漏洞处置成效，详细说明与业务保护目标的差距与已达成部分，并全量点明对整体安全运营的影响。

### 页面：威胁运营成效 (threat_effectiveness)

**threat_improve1**
生成威胁运营改进措施1。数据从threat_effectiveness获取。


**threat_improve2** 
生成威胁运营改进措施2。数据从threat_effectiveness获取。


**threat_improve3**
生成威胁运营改进措施3。数据从threat_effectiveness获取。


**threat_summary**
基于威胁运营数据threat_effectiveness，生成威胁运营工作总结。包括威胁检测覆盖、告警处置效率、主要威胁类型、防护效果等。

### 页面：重要时期保障成效 (critical_assurance)

**duty_summary**
基于period_comparison的每个节日，生成差不多的话术，突出保障效果。模板参考‘中秋节值守保障
2024年9月15日-9月17日 7*24小时全程值守，保障网络安全0事故'，并且对于峰值特殊的的节日加上‘实时攻击达到峰值’等突出趋势的话术


**security_value1** 
基于重保值守数据critical_assurance，生成重保值守工作总结。包括值守时间、保障期间的安全态势、应急响应情况、重点防护措施等。


**security_value2** 
基于重保值守数据critical_assurance，生成重保值守工作总结。包括值守时间、保障期间的安全态势、应急响应情况、重点防护措施等。


**security_value3** 
基于重保值守数据critical_assurance，生成重保值守工作总结。包括值守时间、保障期间的安全态势、应急响应情况、重点防护措施等。

### 页面：安全设备与平台成效 (platform_effectiveness)

**platform_improve1**
生成平台运营改进措施1。从platform_effectiveness获取。


**platform_improve2** 
生成平台运营改进措施2。从platform_effectiveness获取。


**platform_improve3** 
生成平台运营改进措施3。从platform_effectiveness获取。


## 输入数据
json { "incident_effectiveness": { "incident_total": "33", "average_response_time": "25.47", "average_resolution_duration": "572.02", "event_closed_loop_rate": "100%", "response_timeliness": { "labels": [ "识别", "响应", "遏制", "处置", "闭环" ], "values": [ "6.42", "25.47", "10.20", "572.02", "625.38" ] }, "P11_bar": { "labels": [ "识别", "响应", "遏制", "处置", "闭环" ], "values": [ "6.42", "25.47", "10.20", "572.02", "625.38" ] }, "response_trend": { "months": [ "2025-03", "2025-04", "2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12" ], "avg_response_minutes": [ 26.6, 15.25, 0, 18, 10.67, 0, 0, 0, 0, 0 ] }, "P11_line": { "months": [ "2025-03", "2025-04", "2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12" ], "avg_response_minutes": [ 26.6, 15.25, 0, 18, 10.67, 0, 0, 0, 0, 0 ] }, "incident_distribution": { "categories": [ "挖矿", "僵尸网络", "木马", "账号爆破", "代理工具" ], "values": [ 76, 39, 31, 26, 11 ] }, "P11_pie": { "categories": [ "挖矿", "僵尸网络", "木马", "账号爆破", "代理工具" ], "values": [ 76, 39, 31, 26, 11 ] } }, "asset_management": { "internal_network_server_count": "100", "internal_network_network_device_count": "30", "internal_network_IoT_device_count": "10", "internal_network_MSS_service_asset_count": "100", "external_root_domain_asset_count": "11", "external_subdomain_asset_count": "11", "web_asset": "11", "nonweb_asset": "11", "login_endpoint_count": "11", "total_server_assets_count": "125", "PC_assets_count": "31", "internet_IP_domain_count": "61", "internet_exposed_ports_count": "161", "asset_identification_count": "12", "asset_distribution": { "categories": [ "服务器", "终端", "网络设备", "安全设备", "物联网设备" ], "values": [ 10, 10, 10, 10, 10 ] }, "P13_pie": { "categories": [ "服务器", "终端", "网络设备", "安全设备", "物联网设备" ], "values": [ 10, 10, 10, 10, 10 ] } }, "vulnerability_effectiveness": { "high_risk_exploitable_vulnerability_count": "31", "closed_loop_external_asset_vulnerability_count": "185", "admin_weak_password_count": "11", "high_risk_exploitable_vulnerability_closure_rate": "72%", "scanning": "111111111", "vulnerability_distribution": { "categories": [ "高危漏洞", "中危漏洞", "低危漏洞" ], "values": [ 2537, 3633, 3284 ] }, "P14_pie": { "categories": [ "高危漏洞", "中危漏洞", "低危漏洞" ], "values": [ 2537, 3633, 3284 ] }, "closed_loop_counts": { "categories": [ "高危漏洞", "中危漏洞", "低危漏洞" ], "values": [ 262, 262, 212 ] }, "closed_loop_rates": { "categories": [ "高危漏洞", "中危漏洞", "低危漏洞" ], "values": [ "10.33%", "7.21%", "6.46%" ] } }, "threat_effectiveness": { "threat_trend": { "internal_lateral_attack_counts": [ 303457, 222567, 241853, 195076, 104462, 325269, 228846, 226718, 287975, 165318, 101338, 235646 ], "alert_counts": [ 15, 18, 13, 42, 9654, 0, 0, 0, 0, 0, 0, 0 ], "valid_incident_counts": [ 22, 19, 9, 16, 16, 0, 0, 0, 0, 0, 0, 0 ], "risk_host_counts": [ 40, 70, 35, 10, 2, 1, 0, 1, 2, 0, 1, 4 ], "months": [ "2025-03", "2025-04", "2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12", "2026-01" ], "external_attacks": [ 64198, 35890, 12478, 5613, 704, 902, 2412, 1362, 5696, 9140, 4126 ], "malicious_outbound": [ 770729, 665038, 1739364, 616425, 394727, 333492, 194377, 82013, 141130, 197098, 186542 ] }, "P15_line": { "internal_lateral_attack_counts": [ 303457, 222567, 241853, 195076, 104462, 325269, 228846, 226718, 287975, 165318, 101338, 235646 ], "alert_counts": [ 15, 18, 13, 42, 9654, 0, 0, 0, 0, 0, 0, 0 ], "valid_incident_counts": [ 22, 19, 9, 16, 16, 0, 0, 0, 0, 0, 0, 0 ], "risk_host_counts": [ 40, 70, 35, 10, 2, 1, 0, 1, 2, 0, 1, 4 ], "months": [ "2025-03", "2025-04", "2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12", "2026-01" ], "external_attacks": [ 64198, 35890, 12478, 5613, 704, 902, 2412, 1362, 5696, 9140, 4126 ], "malicious_outbound": [ 770729, 665038, 1739364, 616425, 394727, 333492, 194377, 82013, 141130, 197098, 186542 ] }, "external_attack_log_count_XDR": "11", "real_time_threat_alert_count": "9630", "MSS_threat_ticket_count": "0", "threat_ticket_average_response_time": "28", "security_device_policy_check_count": "11", "optimized_policy_risk_count": "11", "latest_threat_intelligence_count": "46", "latest_threat_impacted_asset_count": "11" }, "critical_assurance": { "duty_critical": "7", "incident_critical": "0", "availability_assure": "100%", "posture_comparison": { "categories": [ "中秋（24/9/15-9/17）", "国庆（24/10/1-10/7）", "元旦（25/1/1）", "春节（25/1/23-2/5）", "清明（25/4/4-4/6）", "五一（25/5/1-5/5）", "端午（25/5/31-6/2）" ], "attack_counts": [ 96230, 28655, 3461, 16440, 2238, 7218, 5570 ], "defense_rates": [ 0.5, 0.1, 0.2, 0.8, 0.6, 1, 1 ] }, "P16_combo": { "categories": [ "中秋（24/9/15-9/17）", "国庆（24/10/1-10/7）", "元旦（25/1/1）", "春节（25/1/23-2/5）", "清明（25/4/4-4/6）", "五一（25/5/1-5/5）", "端午（25/5/31-6/2）" ], "attack_counts": [ 96230, 28655, 3461, 16440, 2238, 7218, 5570 ], "defense_rates": [ 0.5, 0.1, 0.2, 0.8, 0.6, 1, 1 ] } }, "platform_effectiveness": { "firewall_detected_attack_count": "21", "firewall_automatic_block_rate": "22", "firewall_protection_path_coverage_rate": "23", "AES_detected_endpoint_security_risk_count": "24", "host_anomaly_risk_handling_count": "25", "server_endpoint_coverage_rate": "26", "XDR_total_security_log_count": "27", "aggregated_security_alert_count": "28", "intelligent_security_incident_identification_count": "29", "component_network_connectivity_anomaly_count": "4", "component_log_synchronization_anomaly_count": "0", "component_policy_effectiveness_alert_count": "0", "component_anomaly_automatic_handling_count": "1", "AES_trusted_risk_event_count": "31", "agent_installation_count": "32", "total_asset_count": "33", "XDR_monthly_average_log_count": "34", "XDR_monthly_average_alert_count": "35", "XDR_monthly_average_incident_count": "36" } }"""

OUTPUT_TEXT = """ { "slides": [ { "slide_key": "incident_effectiveness", "placeholders": { "business_continuity_assurance_conclusion": "本年度共处置安全事件33起，事件闭环率达100%，平均响应与处置时长分别为25.47分钟和572.02分钟，有效遏制了挖矿（76起）与僵尸网络（39起）等主要威胁；通过高效的应急响应机制，核心业务在遭受外部攻击期间未发生中断，有力保障了业务可用性与连续性。", "user_trust_assurance_conclusion": "虽然事件处置闭环率达100%，但挖矿和僵尸网络事件占比超过50%，表明存在部分资产被恶意利用的风险。鉴于核心资产高度集中，一旦遭受类似供应链中断或厂商停服等不可抗力，当前处置能力虽能应对常规入侵，但仍需强化第三方依赖评估以降低单点故障风险。" } }, { "slide_key": "asset_management", "placeholders": { "asset_improve1": "针对互联网暴露端口161个的情况，建议立即开展端口最小化收敛工作，关闭非业务必需端口，降低网络攻击面。", "asset_improve2": "鉴于存在11个弱口令风险（admin_weak_password_count），建议强制实施定期密码轮换策略，并推广多因素认证以杜绝账号爆破风险。", "asset_improve3": "针对125台服务器及物联网设备，建议建立动态资产测绘机制，确保新增资产被及时纳管，消除资产盲区。", "asset_summary": "本年度全面纳管服务器、终端及网络设备等140+台资产，识别并治理11个弱口令风险。虽然资产管理覆盖面广，但暴露端口较多，需持续收敛攻击面。", "asset_value1": "资产全面纳管", "asset_value2": "风险精准收敛", "asset_value3": "暴露面有效降" } }, { "slide_key": "vulnerability_effectiveness", "placeholders": { "vuln_summary": "发现高危漏洞2537个，修复率达10.33%，有效遏制了31个高风险可利用漏洞。但中低危漏洞修复率不足10%，整体修复效能有待提升，建议加强漏洞全生命周期管理以消除潜在业务风险。" } }, { "slide_key": "threat_effectiveness", "placeholders": { "threat_improve1": "针对7月出现的9654次实时威胁告警峰值，建议优化告警聚合规则，减少误报对运维精力的消耗。", "threat_improve2": "鉴于横向移动攻击次数高达32万次，建议加强内网微隔离建设，防止内网扩散。", "threat_improve3": "针对恶意外联流量高企的情况，建议部署更严格的出站流量控制策略，阻断数据外泄通道。", "threat_summary": "本年度累计监测外部攻击11万次，处置实时威胁告警9630次。虽然成功抵御了高频次的横向移动攻击，但7月份告警激增显示防护策略在应对突发流量时仍需优化。" } }, { "slide_key": "critical_assurance", "placeholders": { "duty_summary": "中秋节值守保障：2024年9月15日-9月17日 7*24小时全程值守，实时攻击达到96230次峰值，保障网络安全0事故。国庆节值守保障：2024年10月1日-10月7日 7*24小时全程值守，保障网络安全0事故。春节值守保障：2025年1月23日-2月5日 7*24小时全程值守，保障网络安全0事故。五一节值守保障：2025年5月1日-5月5日 7*24小时全程值守，保障网络安全0事故。", "security_value1": "重保零事故", "security_value2": "全时段响应", "security_value3": "威胁有效防" } }, { "slide_key": "platform_effectiveness", "placeholders": { "platform_improve1": "防火墙自动阻断率仅22%，建议调整防护策略配置，提升自动化处置能力。", "platform_improve2": "服务器端点覆盖率为26%，建议扩大Agent部署范围，消除主机监控死角。", "platform_improve3": "针对组件网络连接异常4次的情况，建议排查底层网络链路，确保平台数据传输稳定性。" } } ] }"""


def count_tokens(text: str, model: str = "gpt-4o-mini") -> tuple[int, str]:
    try:
        import tiktoken  # type: ignore

        encoding = tiktoken.encoding_for_model(model)
        return len(encoding.encode(text)), "tiktoken"
    except Exception:
        cjk_chars = re.findall(r"[\u4e00-\u9fff]", text)
        others = re.findall(r"[A-Za-z0-9_]+|[^\sA-Za-z0-9_]", text)
        return len(cjk_chars) + len(others), "fallback_estimation"


def calc_cost(
    input_tokens: int,
    output_tokens: int,
    input_rate_per_million: Decimal,
    output_rate_per_million: Decimal,
) -> Decimal:
    input_cost = Decimal(input_tokens) / Decimal(1_000_000) * input_rate_per_million
    output_cost = Decimal(output_tokens) / Decimal(1_000_000) * output_rate_per_million
    return input_cost + output_cost


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Calculate token usage price from hardcoded input/output text."
    )
    parser.add_argument(
        "--input-rate",
        type=Decimal,
        default=Decimal("2"),
        help="Input price per 1,000,000 tokens (default: 2).",
    )
    parser.add_argument(
        "--output-rate",
        type=Decimal,
        default=Decimal("8"),
        help="Output price per 1,000,000 tokens (default: 8).",
    )
    parser.add_argument(
        "--model",
        default="4.7",
        help="Model name used by tiktoken (default: 4.7).",
    )
    args = parser.parse_args()

    input_tokens, input_method = count_tokens(INPUT_TEXT, args.model)
    output_tokens, output_method = count_tokens(OUTPUT_TEXT, args.model)
    total_cost = calc_cost(
        input_tokens, output_tokens, args.input_rate, args.output_rate
    ).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP)

    print(f"input_tokens={input_tokens}")
    print(f"input_count_method={input_method}")
    print(f"output_tokens={output_tokens}")
    print(f"output_count_method={output_method}")
    print(f"input_rate_per_million={args.input_rate}")
    print(f"output_rate_per_million={args.output_rate}")
    print(f"total_cost_yuan={total_cost}")


if __name__ == "__main__":
    main()

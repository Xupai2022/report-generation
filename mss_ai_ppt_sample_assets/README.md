# MSS AI PPT 实验平台 - 样例资产包

本目录用于支撑你要做的「本地文件系统 + 两套PPT模板 + 伪造数据」的最小可运行实验。

## 目录结构

```
backend/
  data/
    inputs/              # 伪造的原始输入数据（用于 Data Prep）
    templates/           # 2套模板：pptx + descriptor
    mock_outputs/        # 可选：不走LLM也能跑通的“假LLM输出”(SlideSpec)
  outputs/
    reports/             # 生成PPT输出目录（运行时写入）
    logs/                # 审计日志输出目录（运行时写入）
```

## 两套模板说明

1. **管理层模板（浅色）**
- pptx: `mss_management_light_v1.pptx`
- descriptor: `mss_management_light_v1_descriptor.json`
- slides: 5页
  - cover / executive_summary / alerts_overview / major_incidents / recommendations

2. **技术模板（深色）**
- pptx: `mss_technical_dark_v1.pptx`
- descriptor: `mss_technical_dark_v1_descriptor.json`
- slides: 6页
  - cover / alerts_detail / incident_timeline / vulnerability_exposure / cloud_attack_surface / appendix_evidence

## Token 规则

模板内所有可替换字段都用 `{TOKEN}` 表示，例如 `{REPORT_TITLE}`。渲染时做字符串替换即可。

## 输入数据（原始输入）

`data/inputs/` 下提供 3 份样例：
- `tenant_acme_2025-11_mss_input.json`
- `tenant_acme_2025-12_mss_input.json`
- `tenant_globex_2025-11_mss_input.json`

这些数据包含：告警汇总、事件列表、漏洞态势、云风险、暴露面、MSS运营指标以及 evidence(query_id) 追溯信息。

## mock 输出（可选）

`data/mock_outputs/` 下提供 2 份 SlideSpec（用于离线联调 UI / 渲染，不依赖LLM）：
- `tenant_acme_2025-11_management_mock_slidespec.json`
- `tenant_acme_2025-11_technical_mock_slidespec.json`

注意：这两份文件是“模拟LLM输出”，字段设计与 descriptor 的 render_map 一致。

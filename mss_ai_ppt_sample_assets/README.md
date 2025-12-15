# MSS AI PPT 实验平台 - 样例资产包

本目录用于支撑「本地文件系统 + PPT模板 + 伪造数据 + AI生成」的最小可运行实验。

## 目录结构

```
backend/
  data/
    inputs/              # 伪造的原始输入数据 (TenantInput)
    templates/           # V2 AI驱动模板：pptx + descriptor
  outputs/
    reports/             # 生成PPT输出目录（运行时写入）
    logs/                # 审计日志输出目录（运行时写入）
    previews/            # 预览图缓存
    slidespecs/          # 中间态数据
```

## 模板说明 (V2 AI 驱动)

1. **MSS月度安全报告（管理层版）**
- ID: `mss_executive_v2`
- pptx: `mss_executive_v2.pptx`
- descriptor: `mss_executive_v2_descriptor.json`
- slides: 8页
- 特点: AI生成洞察，适合非技术人员阅读

2. **MSS月度安全报告（技术版）**
- ID: `mss_technical_v2`
- pptx: `mss_technical_v2.pptx`
- descriptor: `mss_technical_v2_descriptor.json`
- slides: 10页
- 特点: 深度数据分析，包含技术细节

## 占位符规则 (V2)

PPTX 模板中使用 `{{token}}` 标记占位符。
Descriptor 定义每个 `token` 的生成规则（直接提取数据或通过 AI 生成）。

## 输入数据（原始输入）

`data/inputs/` 下提供 3 份样例：
- `tenant_acme_2025-11_mss_input.json`
- `tenant_acme_2025-12_mss_input.json`
- `tenant_globex_2025-11_mss_input.json`

这些数据包含：告警汇总、事件列表、漏洞态势、云风险、暴露面、MSS运营指标以及 evidence(query_id) 追溯信息。

## Mock 模式

当未配置 LLM (OpenAI API) 或请求参数 `use_mock=true` 时，系统会自动填充预设的占位文本，不调用外部 API。
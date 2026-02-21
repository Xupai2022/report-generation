# Excel数据驱动的PPT报告生成技术方案

## 文档概述

本文档描述了基于Excel数据源，结合AI智能生成技术，自动生成PowerPoint报告的完整技术方案。该方案充分利用现有的V2模板系统，支持数据直接填充和AI智能总结的混合模式。

---

## 系统架构

```
Excel文件 (原始数据源)
    ↓ (openpyxl提取)
JSON数据文件 (结构化数据)
    ↓
【并发安全层】会话隔离 + 文件锁
    ↓
报告生成系统
    ├── 数据占位符提取 (ai_generate=false)
    │   ├── 简单文本替换
    │   ├── 格式化文本
    │   ├── 图表渲染 (柱状图/饼图)
    │   └── 表格渲染
    │
    └── AI内容生成 (ai_generate=true)
        ├── 态势总结
        ├── 趋势分析
        └── 整改建议
    ↓
PPT报告文件 + 预览图 (会话隔离目录)
```

---

## 核心工作流程

### 阶段1：数据提取与转换

#### 1.1 Excel数据结构

Excel文件建议按“汇总指标 + 明细列表”来组织（以 `samples/data.xlsx` 为例）：

- 汇总指标：集中放在"数据统计"sheet，用于看板/图表/关键数字（如告警/事件、工单与时效、资产与暴露面、漏洞与弱口令、威胁趋势、重要时期值守等）
- 明细列表：按主题拆分到独立sheet（如"资产表"、"资产漏洞表"、"弱密码"、"告警表"、"事件表"），用于PPT表格页的明细呈现与抽样TOP列表

| 数据类型 | Excel位置示例 | 用途 |
|---------|-------------|------|
| 报告元信息 | "数据统计"sheet（固定区域或配置映射） | 客户/项目名称、统计周期、出具时间等 |
| 运营与响应时效 | "数据统计"sheet | 工单量、闭环率、识别/响应/处置/闭环时长与趋势 |
| 威胁与事件 | "数据统计" + "告警表"/"事件表" | 告警/事件总量、威胁类型TOP、来源/地域TOP、重点事件清单 |
| 资产与暴露面 | "数据统计" + "资产表"/"暴露面" | 资产规模与类型分布、互联网IP/域名/端口、暴露面风险趋势 |
| 漏洞与弱口令 | "数据统计" + "资产漏洞表"/"弱密码" | 漏洞分布与闭环率、弱口令风险与整改状态 |

#### 1.2 数据提取脚本

使用`openpyxl`库提取Excel数据并转换为JSON格式：

```python
import openpyxl
import json
from pathlib import Path

def extract_data_from_excel(excel_path: str) -> dict:
    """从Excel文件提取数据并转换为JSON格式

    Args:
        excel_path: Excel文件路径

    Returns:
        结构化的数据字典
    """
    wb = openpyxl.load_workbook(excel_path)
    # 示例：优先读取汇总指标sheet；若仅提供简化模板则读取"数据表"
    ws = wb['数据统计'] if '数据统计' in wb.sheetnames else wb['数据表']

    # 提取基础信息（下述单元格映射仅为“简化模板”示例，实际可通过字段映射表配置）
    data = {
        "customer_name": ws['A2'].value,
        "period": {
            "start": ws['B2'].value,
            "end": ws['C2'].value
        },

        # 提取统计数据
        "alerts": {
            "total": ws['D2'].value,
            "by_severity": {
                "high": ws['E2'].value,
                "medium": ws['F2'].value,
                "low": ws['G2'].value,
                "info": ws['H2'].value
            }
        },

        # 提取分类数据（用于柱状图）
        "top_alert_categories": [],

        # 提取明细数据（用于表格）
        "top_incidents": []
    }

    # 提取告警分类数据（从第5行开始）
    for row in range(5, 11):  # A5:B10
        category = ws[f'A{row}'].value
        count = ws[f'B{row}'].value
        if category and count:
            data["top_alert_categories"].append({
                "category": category,
                "count": count
            })

    # 提取事件明细（从第15行开始）
    for row in range(15, 21):  # A15:E20
        incident_id = ws[f'A{row}'].value
        if incident_id:
            data["top_incidents"].append({
                "id": incident_id,
                "type": ws[f'B{row}'].value,
                "severity": ws[f'C{row}'].value,
                "status": ws[f'D{row}'].value,
                "description": ws[f'E{row}'].value
            })

    return data


def save_to_json(data: dict, output_path: str):
    """保存数据到JSON文件

    Args:
        data: 数据字典
        output_path: 输出JSON文件路径
    """
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ 数据已保存到: {output_path}")


# 使用示例
if __name__ == "__main__":
    excel_file = "data/excel/security_report_2025-12.xlsx"
    json_file = "mss_ai_ppt_sample_assets/backend/data/inputs/excel_report_2025-12.json"

    data = extract_data_from_excel(excel_file)
    save_to_json(data, json_file)
```

#### 1.3 JSON数据格式规范

生成的JSON数据遵循以下结构：

```json
{
  "customer_name": "ACME电子商务",
  "period": {
    "start": "2025-12-01",
    "end": "2025-12-31"
  },
  "alerts": {
    "total": 12455,
    "by_severity": {
      "high": 52,
      "medium": 473,
      "low": 10540,
      "info": 1390
    }
  },
  "top_alert_categories": [
    {"category": "钓鱼攻击", "count": 1244},
    {"category": "暴力破解", "count": 1888},
    {"category": "云配置错误", "count": 1917}
  ],
  "top_incidents": [
    {
      "id": "INC-001",
      "type": "钓鱼攻击",
      "severity": "高危",
      "status": "已解决",
      "description": "发现针对财务部门的钓鱼邮件"
    }
  ]
}
```

---

### 阶段2：模板与占位符定义

#### 2.1 PPT模板设计

PPT模板文件（`report_template.pptx`）包含占位符标记：

```
第1页 - 安全概览
┌─────────────────────────────────────────┐
│  {{CUSTOMER_NAME}}  安全月报             │
│                                         │
│  {{PERIOD_LABEL}}                       │
│                                         │
│  [饼图占位符: {{SEVERITY_PIE}}]         │
│                                         │
│  {{EXECUTIVE_SUMMARY}}                  │
└─────────────────────────────────────────┘

第2页 - 详细分析
┌─────────────────────────────────────────┐
│  告警趋势分析                            │
│                                         │
│  [柱状图占位符: {{ALERT_TREND_BAR}}]    │
│                                         │
│  [表格占位符: {{INCIDENT_TABLE}}]       │
│                                         │
│  {{THREAT_ANALYSIS}}                    │
└─────────────────────────────────────────┘
```

#### 2.2 Descriptor定义文件

模板描述文件（`report_template_descriptor.json`）定义了占位符的处理规则：

```json
{
  "template_id": "excel_security_report_v1",
  "name": "基于Excel的安全报告模板",
  "version": "1.0.0",
  "pptx_file": "report_template.pptx",
  "audience": "management",
  "language": "zh-CN",

  "slides": [
    {
      "slide_no": 1,
      "slide_key": "overview",
      "title": "安全概览",
      "placeholders": [
        {
          "token": "CUSTOMER_NAME",
          "type": "text",
          "ai_generate": false,
          "source": "customer_name",
          "description": "客户名称，直接从数据中提取"
        },
        {
          "token": "PERIOD_LABEL",
          "type": "text",
          "ai_generate": false,
          "source": "period",
          "format": "报告周期：{start} ~ {end}",
          "description": "报告周期，使用格式化模板"
        },
        {
          "token": "SEVERITY_PIE",
          "type": "pie_chart",
          "ai_generate": false,
          "chart_config": {
            "data_source": "alerts.by_severity",
            "category_map": {
              "high": "高危",
              "medium": "中危",
              "low": "低危",
              "info": "信息"
            },
            "position": {
              "left": 1.0,
              "top": 2.5,
              "width": 4.5,
              "height": 3.5
            }
          },
          "description": "告警严重程度分布饼图"
        },
        {
          "token": "EXECUTIVE_SUMMARY",
          "type": "paragraph",
          "ai_generate": true,
          "ai_instruction": "生成面向管理层的安全态势总结。要求：1) 包含告警总数和高危数量；2) 说明整体安全态势（平稳/恶化/改善）；3) 提及主要威胁类型。从alerts和top_alert_categories数据中提取信息。语言简洁专业，适合非技术管理层阅读。",
          "max_length": 150,
          "description": "AI生成的执行摘要"
        }
      ]
    },

    {
      "slide_no": 2,
      "slide_key": "details",
      "title": "详细分析",
      "placeholders": [
        {
          "token": "ALERT_TREND_BAR",
          "type": "bar_chart",
          "ai_generate": false,
          "chart_config": {
            "data_source": "top_alert_categories",
            "x_field": "category",
            "y_field": "count",
            "series_name": "告警数量",
            "position": {
              "left": 1.0,
              "top": 1.5,
              "width": 8.0,
              "height": 3.0
            }
          },
          "description": "Top告警类型柱状图"
        },
        {
          "token": "INCIDENT_TABLE",
          "type": "native_table",
          "ai_generate": false,
          "table_config": {
            "data_source": "top_incidents",
            "max_rows": 5,
            "columns": [
              {"header": "事件ID", "field": "id", "width": 1.5},
              {"header": "类型", "field": "type", "width": 2.0},
              {"header": "严重程度", "field": "severity", "width": 1.5},
              {"header": "状态", "field": "status", "width": 1.5},
              {"header": "描述", "field": "description", "width": 3.0}
            ],
            "position": {
              "left": 1.0,
              "top": 5.0,
              "width": 9.5,
              "height": 2.5
            }
          },
          "description": "重点事件明细表"
        },
        {
          "token": "THREAT_ANALYSIS",
          "type": "bullet_list",
          "ai_generate": true,
          "ai_instruction": "基于告警分类数据和事件明细，生成威胁分析要点。要求：1) 分析主要威胁类型的特点；2) 说明潜在影响；3) 提供应对建议。每条要点需具体且可执行。",
          "max_items": 4,
          "max_chars_per_item": 60,
          "description": "AI生成的威胁分析"
        }
      ]
    }
  ]
}
```

---

### 阶段3：报告生成处理

#### 3.1 系统处理流程

系统按以下步骤处理报告生成请求：

**Step 1: 加载数据和模板**
```python
# 系统内部代码（无需修改）
tenant_input = TenantInput.from_file("data/inputs/excel_report_2025-12.json")
template = template_repo.get_descriptor_v2("excel_security_report_v1")
```

**Step 2: 提取数据占位符（ai_generate=false）**
```python
# 处理所有非AI占位符
data_placeholders = llm_orchestrator._extract_data_placeholders(tenant_input, template)

# 结果示例：
{
  "overview": {
    "CUSTOMER_NAME": "ACME电子商务",
    "PERIOD_LABEL": "报告周期：2025-12-01 ~ 2025-12-31",
    "SEVERITY_PIE": {
      "categories": ["高危", "中危", "低危", "信息"],
      "values": [52, 473, 10540, 1390],
      "position": {...}
    }
  },
  "details": {
    "ALERT_TREND_BAR": {
      "categories": ["钓鱼攻击", "暴力破解", "云配置错误"],
      "series": [{"name": "告警数量", "values": [1244, 1888, 1917]}],
      "position": {...}
    },
    "INCIDENT_TABLE": {
      "headers": ["事件ID", "类型", "严重程度", "状态", "描述"],
      "rows": [
        ["INC-001", "钓鱼攻击", "高危", "已解决", "发现针对财务部门的钓鱼邮件"],
        ["INC-002", "恶意软件", "中危", "处理中", "办公终端发现木马程序"]
      ],
      "position": {...}
    }
  }
}
```

**Step 3: AI生成内容（ai_generate=true）**
```python
# 构建AI提示词
system_prompt = "你是资深安全分析师，正在撰写安全报告..."
user_prompt = """
## 原始数据
{完整的JSON数据}

## 需要生成的内容
### Slide: 安全概览 (overview)

**EXECUTIVE_SUMMARY** (最多150字)
生成面向管理层的安全态势总结。要求：1) 包含告警总数和高危数量...

### Slide: 详细分析 (details)

**THREAT_ANALYSIS** (最多4条, 每条最多60字)
基于告警分类数据和事件明细，生成威胁分析要点...
"""

# AI返回结果
{
  "slides": [
    {
      "slide_key": "overview",
      "placeholders": {
        "EXECUTIVE_SUMMARY": "本月共处理12,455条告警，其中高危52条已全部解决。整体安全态势平稳可控，主要威胁集中在钓鱼攻击和云配置错误，环比上升30%，建议加强云安全管控。"
      }
    },
    {
      "slide_key": "details",
      "placeholders": {
        "THREAT_ANALYSIS": "• 云配置错误告警环比激增30%，需立即开展配置审计\n• 钓鱼攻击持续活跃，建议加强员工安全意识培训\n• 暴力破解攻击主要针对SSH服务，已启用IP白名单防护\n• 2起高危事件均已完成根因分析和整改闭环"
      }
    }
  ]
}
```

**Step 4: 合并数据并渲染PPT**
```python
# 合并数据和AI生成的内容
slidespec.placeholders = {
    "CUSTOMER_NAME": "ACME电子商务",
    "PERIOD_LABEL": "报告周期：2025-12-01 ~ 2025-12-31",
    "SEVERITY_PIE": {...},  # 饼图数据
    "EXECUTIVE_SUMMARY": "本月共处理12,455条告警...",  # AI生成
    "ALERT_TREND_BAR": {...},  # 柱状图数据
    "INCIDENT_TABLE": {...},  # 表格数据
    "THREAT_ANALYSIS": "• 云配置错误告警..."  # AI生成
}

# 渲染PPT（替换占位符、绘制图表、生成表格）
ppt_generator.render(slidespec, output_path)
```

#### 3.2 API调用示例

**方式A：直接使用JSON数据（开发/测试）**

```bash
# 1. 上传JSON数据到inputs目录
# 2. 调用生成API
curl -X POST http://localhost:8000/api/v1/generate \
  -H "Content-Type: application/json" \
  -d '{
    "input_id": "excel_report_2025-12",
    "template_id": "excel_security_report_v1"
  }'

# 响应示例
{
  "status": "success",
  "job_id": "excel_report_2025-12:excel_security_report_v1",
  "output_path": "outputs/reports/excel_report_2025-12_excel_security_report_v1.pptx",
  "slidespec_path": "outputs/slidespecs/excel_report_2025-12_excel_security_report_v1.json"
}

# 3. 获取预览图
curl "http://localhost:8000/api/v1/reports/excel_report_2025-12%3Aexcel_security_report_v1/preview?regenerate_if_missing=true"

# 响应示例
{
  "status": "success",
  "preview_urls": [
    "/static/previews/excel_report_2025-12:excel_security_report_v1/slide1.png",
    "/static/previews/excel_report_2025-12:excel_security_report_v1/slide2.png"
  ]
}
```

**方式B：Web上传Excel文件（生产环境推荐）** ⭐

系统提供 `/upload-excel` 端点，支持用户直接上传Excel文件，自动解析并生成报告。

```bash
# 1. 上传Excel文件
curl -X POST http://localhost:8000/upload-excel \
  -F "file=@security_report_2025-12.xlsx"

# 响应示例
{
  "status": "success",
  "session_id": "a3f2c5d8_20250129143025123456",
  "filename": "security_report_2025-12.xlsx",
  "file_size_mb": 0.45,
  "preview": {
    "customer_name": "ACME电子商务",
    "period": {"start": "2025-12-01", "end": "2025-12-31"},
    "alerts": {"total": 12455, "high": 52},
    "categories_count": 6,
    "incidents_count": 5
  },
  "next_steps": {
    "description": "使用此session_id调用 /generate 端点生成报告"
  }
}

# 2. 使用返回的session_id生成报告
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d '{
    "input_id": "custom",
    "template_id": "mss_executive_v2",
    "session_id": "a3f2c5d8_20250129143025123456"
  }'

# 3. 下载生成的报告
curl -O "http://localhost:8000/download?job_id=a3f2c5d8_...:mss_executive_v2"
```

**Excel上传安全机制**：
- ✅ 仅接受 `.xlsx` 格式（无宏Excel），拒绝 `.xlsm`、`.xls`、`.xlsb`
- ✅ 文件大小限制 10MB（可配置）
- ✅ MIME类型双重验证，防止文件伪装
- ✅ 会话隔离存储，多用户并发安全
- ✅ `openpyxl` 只读模式，不执行宏或公式

**实现代码**：参见 `mss_ai_ppt_sample_assets/backend/excel_upload_endpoint.py`

---

## 占位符类型与配置详解

### 1. 文本类占位符

#### 1.1 简单文本替换

```json
{
  "token": "CUSTOMER_NAME",
  "type": "text",
  "ai_generate": false,
  "source": "customer_name"
}
```

**数据要求：** 字符串
```json
{"customer_name": "ACME电子商务"}
```

**渲染结果：** `ACME电子商务`

---

#### 1.2 格式化文本

```json
{
  "token": "PERIOD_LABEL",
  "type": "text",
  "ai_generate": false,
  "source": "period",
  "format": "报告周期：{start} ~ {end}"
}
```

**数据要求：** 对象
```json
{
  "period": {
    "start": "2025-12-01",
    "end": "2025-12-31"
  }
}
```

**渲染结果：** `报告周期：2025-12-01 ~ 2025-12-31`

---

#### 1.3 转换函数

```json
{
  "token": "CUSTOMER_REGION",
  "type": "text",
  "ai_generate": false,
  "source": "region",
  "transform": "uppercase"
}
```

**支持的transform：**
- `uppercase`: 转大写
- `lowercase`: 转小写
- `percent`: 转百分比（0.85 → 85%）

---

#### 1.4 列表格式化

```json
{
  "token": "KEY_SERVICES",
  "type": "text",
  "ai_generate": false,
  "source": "key_services",
  "format": "join_comma"
}
```

**数据要求：** 数组
```json
{"key_services": ["Kubernetes", "MySQL", "Redis"]}
```

**渲染结果：** `Kubernetes, MySQL, Redis`

---

### 2. 图表类占位符

系统支持多种图表类型，每种图表类型都有专门的渲染函数和样式配置。

#### 2.1 柱状图类型

##### 2.1.1 通用柱状图（bar_chart）

通用柱状图适用于大多数场景，使用默认的专业配色方案。

**支持两种数据格式：**

**格式A：字典格式**
```json
{
  "token": "ALERT_BAR",
  "type": "bar_chart",
  "ai_generate": false,
  "chart_config": {
    "data_source": "alert_stats",
    "x_field": "labels",
    "y_field": "values",
    "series_name": "告警数",
    "position": {
      "left": 1.0,
      "top": 2.0,
      "width": 8.0,
      "height": 4.0
    }
  }
}
```

**数据示例：**
```json
{
  "alert_stats": {
    "labels": ["钓鱼", "暴力破解", "配置错误"],
    "values": [1244, 1888, 1917]
  }
}
```

#### 2.1.2 P11_bar 处置时间柱状图 ⭐

**专门用于展示处置时间数据**，使用**模板内置图表 + 占位符定位**的渲染模式。

**渲染模式特点：**
- 🎯 **占位符定位**：在模板中添加 `{{TOKEN}}` 占位符文本框，代码通过占位符位置查找最近的图表
- 📊 **模板内置图表**：在模板中手动添加柱状图，保留手动配置的样式（图例位置、字体大小等）
- 🔄 **数据更新**：代码只更新图表数据，不改变图表的布局和样式配置
- 🗑️ **自动清理**：渲染完成后自动删除占位符文本框

**样式特点：**
- **柱子颜色**：RGB(68, 114, 196) 专业蓝灰色
- **数据标签**：柱子顶部显示具体数值（保留两位小数，如 6.42、572.02）
- **图例位置**：顶部显示"平均处置时间（分钟）"
- **Y轴刻度**：显示网格线，数值格式为 `0.00`（如 100.00, 200.00, 300.00...）
- **边框**：无边框，简洁清爽的设计
- **可编辑性**：✅ 生成的是**原生PowerPoint图表对象**，用户可以在PPT中：
  - 右键点击图表 → 编辑数据
  - 修改图表样式、颜色
  - 添加或删除数据系列
  - 完全自定义图表属性

**模板准备步骤：**
1. 在PPT模板中手动添加柱状图（通过"插入→图表"）
2. 手动配置图表样式：图例位置、字体、颜色等
3. 在图表附近添加文本框，内容为 `{{TOKEN_NAME}}`（如 `{{RESPONSE_TIME_BAR}}`）
4. 占位符位置建议：图表正上方或正下方

**占位符配置：**
```json
{
  "token": "RESPONSE_TIME_BAR",
  "type": "P11_bar",
  "ai_generate": false,
  "chart_config": {
    "data_source": "response_time_by_severity",
    "x_field": "severity_levels",
    "y_field": "avg_minutes",
    "series_name": "平均处置时间（分钟）"
  }
}
```

**注意：**
- ⚠️ `chart_config` 中的 `position` 字段**不再使用**（由模板中的图表位置决定）
- ✅ 占位符可以使用任何名称（如 `{{abc}}`、`{{哈哈哈}}`），代码会自动匹配并删除
- ✅ 如果一个幻灯片有多个同类型图表，代码会自动查找距离占位符最近的图表

**数据要求示例：**
```json
{
  "response_time_by_severity": {
    "severity_levels": ["识别", "响应", "处置", "闭环"],
    "avg_minutes": [6.42, 27.40, 572.02, 625.38]
  }
}
```

**工作原理：**
```
1. 查找占位符 {{RESPONSE_TIME_BAR}}
   → 记录占位符位置 (left, top)

2. 查找距离占位符最近的柱状图
   → 计算所有柱状图到占位符的距离
   → 选择距离最小的图表

3. 更新图表数据
   → 使用 chart.replace_data() 替换数据
   → 保留模板中配置的样式

4. 删除占位符文本框
   → 保持PPT整洁
```

---

#### 2.2 通用柱状图的其他格式

**格式B：对象数组格式**
```json
{
  "chart_config": {
    "data_source": "top_categories",
    "x_field": "category",
    "y_field": "count"
  }
}
```

**数据示例：**
```json
{
  "top_categories": [
    {"category": "钓鱼", "count": 1244},
    {"category": "暴力破解", "count": 1888}
  ]
}
```

---

#### 2.3 饼图类型

##### 2.3.1 通用饼图（pie_chart）

适用于一般的饼图展示，使用程序化创建图表。

```json
{
  "token": "SEVERITY_PIE",
  "type": "pie_chart",
  "ai_generate": false,
  "chart_config": {
    "data_source": "alerts.by_severity",
    "category_map": {
      "high": "高危",
      "medium": "中危",
      "low": "低危"
    },
    "position": {
      "left": 1.0,
      "top": 2.5,
      "width": 5.0,
      "height": 4.0
    }
  }
}
```

**数据要求：** 字典
```json
{
  "alerts": {
    "by_severity": {
      "high": 52,
      "medium": 473,
      "low": 10540
    }
  }
}
```

**category_map：** 将数据键名映射为图表标签（可选）

##### 2.3.2 P12_pie 威胁类型分布环形图 ⭐

**专门用于展示威胁类型分布数据**，使用**模板内置图表 + 占位符定位**的渲染模式。

**占位符配置：**
```json
{
  "token": "P12_pie",
  "type": "P12_pie",
  "ai_generate": false,
  "chart_config": {
    "data_source": "threat_distribution"
  }
}
```

**数据要求示例：**
```json
{
  "threat_distribution": {
    "categories": ["挖矿", "僵尸网络", "木马", "账号爆破", "代理工具"],
    "values": [100, 100, 200, 200, 400]
  }
}
```

---

##### 2.3.3 P13_pie 资产分布环形图 ⭐

**专门用于展示资产类型分布等数据**，使用**模板内置图表 + 占位符定位**的渲染模式。

**渲染模式特点：**
- 🎯 **占位符定位**：在模板中添加 `{{TOKEN}}` 占位符文本框，代码通过占位符位置查找最近的饼图
- 📊 **模板内置图表**：在模板中手动添加环形图（Donut Chart），保留手动配置的样式（图例位置、字体大小等）
- 🔄 **数据更新**：代码只更新图表数据，不改变图表的布局和样式配置
- 🗑️ **自动清理**：渲染完成后自动删除占位符文本框

**样式特点：**
- **图表类型**：环形图（Donut Chart，中间有空洞的饼图）
- **自定义颜色**：支持自定义每个扇区的颜色（如资产类型：服务器/终端/网络设备等）
- **数据标签**：显示百分比，支持自定义字体大小
- **图例位置**：由模板配置决定（建议右侧竖排）
- **扇区间隔**：白色边框分隔各扇区，视觉更清晰
- **可编辑性**：✅ 生成的是**原生PowerPoint图表对象**

**模板准备步骤：**
1. 在PPT模板中手动添加环形图（插入→图表→圆环图）
2. 手动配置图表样式：图例位置（建议右侧）、字体大小、数据标签格式等
3. 在图表附近添加文本框，内容为 `{{TOKEN_NAME}}`（如 `{{ASSET_CHART}}`）
4. 占位符位置建议：图表正上方或正下方

**占位符配置：**
```json
{
  "token": "ASSET_CHART",
  "type": "P13_pie",
  "ai_generate": false,
  "chart_config": {
    "data_source": "asset_distribution"
  }
}
```

**注意：**
- ⚠️ `chart_config` 中的 `position` 字段**不再使用**（由模板中的图表位置决定）
- ✅ 占位符可以使用任何名称，代码会自动匹配并删除
- ✅ 如果一个幻灯片有多个饼图，代码会自动查找距离占位符最近的饼图
- ✅ 支持自定义颜色方案（在代码中预定义或通过模板配置）

**数据要求示例：**
```json
{
  "asset_distribution": {
    "categories": ["服务器", "终端", "网络设备", "安全设备", "物联网设备"],
    "values": [100, 24, 19, 10, 8]
  }
}
```

**内置颜色方案示例（P13_pie）：**
```python
p13_colors = [
    (72, 116, 203),    # 服务器 - Blue
    (238, 130, 47),    # 终端 - Orange
    (117, 189, 66),    # 网络设备 - Green
    (242, 186, 2),     # 安全设备 - Yellow
    (48, 192, 180),    # 物联网设备 - Teal
]
```

**工作原理：**
```
1. 查找占位符 {{ASSET_CHART}}
   → 记录占位符位置 (left, top)

2. 查找距离占位符最近的饼图/环形图
   → 计算所有饼图到占位符的距离
   → 选择距离最小的图表

3. 更新图表数据
   → 使用 chart.replace_data() 替换数据
   → 应用自定义颜色到各扇区
   → 保留模板中配置的图例位置、字体等

4. 删除占位符文本框
   → 保持PPT整洁
```

---

##### 2.3.4 P14_pie 告警严重程度分布环形图 ⭐

**专门用于展示告警严重程度分布（高危/中危/低危）**，使用**模板内置图表 + 占位符定位**的渲染模式。

**占位符配置：**
```json
{
  "token": "P14_pie",
  "type": "P14_pie",
  "ai_generate": false,
  "chart_config": {
    "data_source": "severity_distribution"
  }
}
```

**数据要求示例：**
```json
{
  "severity_distribution": {
    "categories": ["高危", "中危", "低危"],
    "values": [62, 837, 9914]
  }
}
```

**内置颜色方案（P14_pie）：**
```python
p14_colors = [
    (220, 38, 38),     # 高危 - Red
    (234, 179, 8),     # 中危 - Yellow/Amber
    (34, 197, 94),     # 低危 - Green
]
```

---

### 3. 表格类占位符

```json
{
  "token": "INCIDENT_TABLE",
  "type": "native_table",
  "ai_generate": false,
  "table_config": {
    "data_source": "top_incidents",
    "max_rows": 10,
    "columns": [
      {"header": "事件ID", "field": "id", "width": 1.5},
      {"header": "类型", "field": "type", "width": 2.0},
      {"header": "状态", "field": "status", "width": 1.5}
    ],
    "position": {
      "left": 1.0,
      "top": 3.0,
      "width": 8.0,
      "height": 3.0
    }
  }
}
```

**数据要求：** 对象数组
```json
{
  "top_incidents": [
    {"id": "INC-001", "type": "钓鱼攻击", "status": "已解决"},
    {"id": "INC-002", "type": "恶意软件", "status": "处理中"}
  ]
}
```

**配置说明：**
- `max_rows`: 最多显示行数
- `columns[].field`: 数据字段名
- `columns[].width`: 列宽（英寸）

---

### 4. AI生成类占位符

#### 4.1 段落文本

```json
{
  "token": "EXECUTIVE_SUMMARY",
  "type": "paragraph",
  "ai_generate": true,
  "ai_instruction": "生成面向管理层的安全态势总结。要求：1) 包含告警总数和高危数量；2) 说明整体安全态势；3) 提及主要威胁类型。",
  "max_length": 150
}
```

**AI生成示例：**
```
本月共处理12,455条告警，其中高危52条已全部解决。整体安全态势平稳可控，主要威胁集中在钓鱼攻击和云配置错误，环比上升30%，建议加强云安全管控。
```

---

#### 4.2 项目列表

```json
{
  "token": "KEY_FINDINGS",
  "type": "bullet_list",
  "ai_generate": true,
  "ai_instruction": "总结本月的关键发现，包括重大事件、主要威胁、整改措施等。",
  "max_items": 5,
  "max_chars_per_item": 50
}
```

**AI生成示例：**
```
• 成功拦截52起高危告警，平均响应时间15分钟
• 处置2起安全事件，已完成根因分析和整改
• 云配置错误告警环比上升30%，需重点关注
• 钓鱼攻击尝试环比下降2%，防护措施有效
• 完成217个核心资产的安全加固
```

---

## 技术实现细节

### 数据路径访问

系统使用点号分隔的路径访问嵌套数据：

```json
{
  "alerts": {
    "total": 12455,
    "by_severity": {
      "high": 52
    }
  }
}
```

**路径示例：**
- `"source": "alerts.total"` → `12455`
- `"source": "alerts.by_severity.high"` → `52`
- `"source": "alerts.by_severity"` → `{"high": 52, ...}`

---

### 图表位置定义

所有位置坐标使用**英寸**为单位：

```json
{
  "position": {
    "left": 1.0,    // 距离左边距1英寸
    "top": 2.5,     // 距离上边距2.5英寸
    "width": 8.0,   // 宽度8英寸
    "height": 4.0   // 高度4英寸
  }
}
```

**PowerPoint标准页面尺寸：** 10英寸 × 7.5英寸（16:9比例）

---

### AI指令编写最佳实践

**好的AI指令：**
```
生成面向管理层的安全态势总结。要求：
1) 包含告警总数和高危数量（从alerts.total和alerts.by_severity.high获取）
2) 说明整体安全态势（平稳/恶化/改善）
3) 提及主要威胁类型（从top_alert_categories中选择前3项）
语言简洁专业，适合非技术管理层阅读。
```

**避免的写法：**
```
写个总结
```

**关键要素：**
1. 明确输出目标和受众
2. 列出具体要求（编号列表）
3. 指明数据来源路径
4. 说明语言风格

---

## 完整示例：端到端流程

### 输入：Excel文件

**security_report_2025-12.xlsx**

| 列A | 列B | 列C | 列D |
|-----|-----|-----|-----|
| 客户名称 | 报告开始 | 报告结束 | 告警总数 |
| ACME电子商务 | 2025-12-01 | 2025-12-31 | 12455 |

| 列A (分类) | 列B (数量) |
|-----------|-----------|
| 钓鱼攻击 | 1244 |
| 暴力破解 | 1888 |
| 云配置错误 | 1917 |

---

### 中间：JSON数据

**excel_report_2025-12.json**

```json
{
  "customer_name": "ACME电子商务",
  "period": {"start": "2025-12-01", "end": "2025-12-31"},
  "alerts": {"total": 12455},
  "top_alert_categories": [
    {"category": "钓鱼攻击", "count": 1244},
    {"category": "暴力破解", "count": 1888},
    {"category": "云配置错误", "count": 1917}
  ]
}
```

---

### 输出：PPT报告

**第1页效果：**
```
┌─────────────────────────────────────────┐
│  ACME电子商务  安全月报                  │  ← 数据填充
│                                         │
│  报告周期：2025-12-01 ~ 2025-12-31      │  ← 格式化
│                                         │
│  [饼图: 告警严重程度分布]                │  ← 图表渲染
│                                         │
│  本月共处理12,455条告警，其中高危52条    │  ← AI生成
│  已全部解决。整体安全态势平稳可控...     │
└─────────────────────────────────────────┘
```

---

## 系统优势

### 1. 灵活的数据源

- ✅ 支持Excel、CSV等任意数据源
- ✅ 只需转换为JSON即可接入
- ✅ 数据提取逻辑完全可自定义

### 2. 混合模式处理

- ✅ 数据直接填充：准确、快速、成本低
- ✅ AI智能生成：灵活、专业、有洞察
- ✅ 同一页面可混合使用

### 3. 丰富的可视化

- ✅ 柱状图、饼图原生渲染
- ✅ 表格自动排版和样式
- ✅ 支持自定义位置和尺寸

### 4. 零代码修改

- ✅ 利用现有V2系统架构
- ✅ 只需定义Descriptor文件
- ✅ 无需修改后端代码

---

## 扩展性说明

### 图表渲染架构（V2 新架构）

系统采用**特定图表类型 + 专用渲染函数 + 占位符定位**的架构，每种图表类型都有独立的渲染函数和样式配置。

#### 两种渲染模式

系统支持两种图表渲染模式，适用于不同场景：

##### 模式A：程序化创建图表（通用图表）

**适用场景**：通用柱状图（bar_chart）、通用饼图（pie_chart）等

**特点**：
- 代码完全控制图表创建和样式
- 使用 `position` 字段指定图表位置和尺寸
- 适合样式简单、不需要复杂手动配置的图表

**示例**：
```json
{
  "token": "ALERT_BAR",
  "type": "bar_chart",
  "chart_config": {
    "data_source": "alert_stats",
    "position": {
      "left": 1.0,
      "top": 2.0,
      "width": 8.0,
      "height": 4.0
    }
  }
}
```

##### 模式B：模板内置图表 + 占位符定位（专用图表）⭐

**适用场景**：P11_bar、P13_pie 等需要精细样式控制的专用图表

**特点**：
- 🎯 **占位符定位机制**：通过 `{{TOKEN}}` 文本框定位图表
- 📊 **保留模板样式**：不破坏模板中手动配置的图例位置、字体等
- 🔄 **只更新数据**：使用 `chart.replace_data()` 仅替换数据，不改变布局
- 🗑️ **自动清理**：渲染完成后删除占位符文本框

**工作流程**：
```
1. 模板准备
   ├─ 手动添加图表到模板（插入→图表）
   ├─ 手动配置样式（图例、颜色、字体等）
   └─ 添加占位符文本框 {{TOKEN}}

2. 代码渲染
   ├─ 查找占位符位置 (left, top)
   ├─ 查找距离最近的对应类型图表
   ├─ 更新图表数据 (chart.replace_data)
   └─ 删除占位符文本框

3. 结果
   └─ 数据更新✅ 样式保留✅ 占位符删除✅
```

**占位符定位算法**：
```python
# 1. 查找占位符
for shape in slide.shapes:
    if shape.text matches r'\{\{[^}]+\}\}':
        placeholder_position = (shape.left, shape.top)
        placeholder_shape = shape
        break

# 2. 查找最近的图表
min_distance = infinity
for shape in slide.shapes:
    if shape.has_chart and chart_type_matches:
        distance = sqrt((chart.left - placeholder.left)² +
                       (chart.top - placeholder.top)²)
        if distance < min_distance:
            selected_chart = shape.chart
            min_distance = distance

# 3. 更新数据
selected_chart.replace_data(new_data)

# 4. 删除占位符
placeholder_shape.delete()
```

**优势**：
- ✅ **完美保留手动配置**：PowerPoint会自动重新计算某些样式（如图例位置），只有模板内置的配置才能稳定保留
- ✅ **灵活的占位符命名**：占位符可以用任何名称（`{{abc}}`、`{{哈哈哈}}`），代码通过正则 `\{\{[^}]+\}\}` 自动识别
- ✅ **支持多图表幻灯片**：通过距离计算准确匹配,一个幻灯片可以有多个同类型图表
- ✅ **兼容不完整占位符**：正则 `\{\{[^}]+\}?\}?` 可以匹配 `{{abc}` 这样的不完整占位符

**示例**：
```json
{
  "token": "RESPONSE_TIME_BAR",
  "type": "P11_bar",
  "chart_config": {
    "data_source": "response_time",
    // 不需要 position 字段
  }
}
```

#### 架构设计

```python
# ppt_generator.py

class PPTGeneratorV2:
    def __init__(self, template_repo):
        # 图表渲染函数映射表
        self._chart_renderers = {
            # 模式A：程序化创建
            'bar_chart': self._render_bar_chart,        # 通用柱状图
            'pie_chart': self._render_pie_chart,        # 通用饼图

            # 模式B：模板内置 + 占位符定位
            'P11_bar': self._render_p11_bar,            # 处置时间柱状图
            'P13_pie': self._render_p13_pie,            # 资产分布环形图
        }

    def _render_p11_bar(self, slide, chart_data, token):
        """渲染 P11 柱状图（模式B）"""
        # 1. 查找占位符位置
        placeholder_position = self._find_placeholder(slide, token)

        # 2. 查找最近的柱状图
        chart = self._find_nearest_chart(slide, placeholder_position,
                                          chart_types=['COLUMN', 'BAR'])

        # 3. 更新数据
        chart.replace_data(chart_data)

        # 4. 删除占位符
        self._remove_placeholder(slide, token)

    def _render_bar_chart(self, slide, chart_data, position):
        """渲染通用柱状图（模式A）"""
        # 程序化创建图表
        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(position['left']),
            Inches(position['top']),
            Inches(position['width']),
            Inches(position['height']),
            chart_data
        )
        # 应用默认样式
        ...
```

#### 架构优势

✅ **每种图表完全独立**
- 不同图表类型互不干扰
- 每个渲染函数专注于一种图表样式
- 易于调试和维护

✅ **代码可读性强**
- 函数名清晰表达用途：`_render_p11_bar()`
- 每个函数内部逻辑简单明了
- 无需复杂的条件判断

✅ **易于扩展**
新增图表类型只需三步：
1. 在 `models/templates.py` 的 `type` 枚举中添加新类型
2. 编写新的 `_render_xxx()` 函数
3. 注册到 `_chart_renderers` 字典

示例：
```python
# 1. 添加类型枚举
type: Literal["text", "paragraph", "P11_bar", "P12_line", ...]

# 2. 编写渲染函数
def _render_p12_line(self, slide, chart_data):
    """渲染 P12 趋势折线图"""
    # 折线图特定的样式
    ...

# 3. 注册渲染器
self._chart_renderers = {
    ...
    'P12_line': self._render_p12_line,
}
```

✅ **灵活性高**
- 每种图表可以使用完全不同的 python-pptx API
- 例如：折线图可以设置 `smooth=True`，饼图可以设置 `explosion`
- 不受通用函数的限制

#### 数据流

```
模板 Descriptor JSON
  ↓
  type: "P11_bar"  (specific chart type)
  chart_config: {...}
  ↓
LLMOrchestrator._extract_chart_data()
  → 根据 chart_type 提取数据
  ↓
SlideSpecV2.placeholders[TOKEN] = {
  'categories': [...],
  'series': [{...}],
  'position': {...}
}
  ↓
PPTGeneratorV2.render()
  → 查找 _chart_renderers['P11_bar']
  → 调用 _render_p11_bar(slide, chart_data)
  ↓
生成原生 PowerPoint 图表对象
  → 用户可在 PPT 中编辑数据和样式
```

### 支持更多图表类型

如需添加折线图、散点图等，可扩展以下模块：

1. **添加类型枚举** (`models/templates.py`)
```python
type: Literal[..., "P12_line", "P13_scatter"]
```

2. **编写渲染函数** (`ppt_generator.py`)
```python
def _render_p12_line(self, slide, chart_data):
    """渲染 P12 趋势折线图"""
    chart_shape = slide.shapes.add_chart(
        self._XL_CHART_TYPE.LINE,
        ...
    )
    # 折线图特定样式
    chart.series[0].smooth = True
    ...
```

3. **注册渲染器**
```python
self._chart_renderers['P12_line'] = self._render_p12_line
```

4. **更新数据提取逻辑** (`llm_orchestrator.py`) - 如果需要特殊的数据格式
```python
if placeholder.type in ('bar_chart', 'pie_chart', 'P11_bar', 'P12_line'):
    ...
```

### 支持复杂数据转换

可在数据提取脚本中添加计算逻辑：

```python
# 计算环比变化
data["alert_mom_change"] = (
    (current_month_alerts - last_month_alerts) / last_month_alerts * 100
)

# 计算合规率
data["compliance_rate"] = passed_checks / total_checks
```

### 支持多语言

在Descriptor中添加locale字段，AI会根据语言生成内容：

```json
{
  "language": "en-US",
  "ai_instruction": "Generate an executive summary in English..."
}
```

---

## 注意事项

### 1. 数据质量

- Excel中的空值会被转换为 `null`
- 确保必填字段有数据
- 数字类型不要存储为文本

### 2. 性能考虑

- AI生成时间取决于占位符数量（通常1-5秒/页）
- 大量图表会增加渲染时间
- 建议单个报告不超过20页

### 3. 成本控制

- 只对必要内容使用AI生成
- 简单数据展示使用直接填充
- 可通过 `enable_llm=false` 禁用AI使用fallback

---

## 总结

本方案充分利用现有V2模板系统的能力，通过简单的数据提取脚本和Descriptor配置，即可实现从Excel到PPT的自动化报告生成。系统支持数据直接填充和AI智能生成的混合模式，既保证了数据准确性，又提供了灵活的内容生成能力。

**关键优势：**
- ✅ 无需修改任何后端代码
- ✅ 数据源灵活可扩展
- ✅ 支持丰富的可视化类型
- ✅ AI与数据填充完美结合

**实施路径：**
1. 编写Excel数据提取脚本（openpyxl）
2. 设计PPT模板文件（带占位符）
3. 编写Descriptor配置文件
4. 调用现有API生成报告

---

## 附录：新增图表类型快速指南

### 如何添加一个新的图表类型（如 P15_line）

**需要修改 4 个文件，按顺序操作：**

#### 1️⃣ 添加数据（可选）
**文件：** `mss_ai_ppt_sample_assets/backend/data/inputs/tenant_acme_2025-11_mss_input.json`

```json
{
  "your_new_data": {
    "categories": ["类别1", "类别2", "类别3"],
    "values": [100, 200, 300]
  }
}
```

---

#### 2️⃣ 注册图表类型
**文件：** `mss_ai_ppt_sample_assets/backend/models/templates.py`

**位置：** 第 17-26 行，`PlaceholderDefinition` 类的 `type` 字段

```python
type: Literal[
    "text", "paragraph", "bullet_list", "kpi", "kpi_group", "table",
    "chart_data", "incident_list", "incident_detail",
    "native_table",
    # Specific chart types
    "P11_bar",   # Response time bar chart
    "P12_pie",   # Threat type distribution donut chart
    "P13_pie",   # Asset distribution donut chart
    "P14_pie",   # Severity distribution donut chart
    "P15_line",  # 👈 新增：你的图表类型
]
```

---

#### 3️⃣ 编写渲染函数
**文件：** `mss_ai_ppt_sample_assets/backend/modules/ppt_generator.py`

**步骤 A：注册渲染器（第 102-110 行）**
```python
self._chart_renderers = {
    'bar_chart': self._render_bar_chart,
    'pie_chart': self._render_pie_chart,
    'P11_bar': self._render_p11_bar,
    'P12_pie': self._render_p12_pie,
    'P13_pie': self._render_p13_pie,
    'P14_pie': self._render_p14_pie,
    'P15_line': self._render_p15_line,  # 👈 新增
}
```

**步骤 B：编写渲染函数（在 `_render_p14_pie` 后面添加）**
```python
def _render_p15_line(
    self,
    slide,
    chart_data: Dict[str, Any],
    token: str = None
) -> None:
    """渲染 P15 折线图（你的图表说明）"""

    # 1. 查找占位符
    placeholder_position = None
    placeholder_shape_to_remove = None
    if token:
        import re
        placeholder_pattern = re.compile(r'\{\{[^}]+\}?\}?')
        for shape in slide.shapes:
            try:
                if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                    full_text = ''.join(
                        run.text for paragraph in shape.text_frame.paragraphs
                        for run in paragraph.runs
                    ).strip()
                    if placeholder_pattern.fullmatch(full_text):
                        placeholder_position = (shape.left, shape.top)
                        placeholder_shape_to_remove = shape
                        break
            except:
                continue

    # 2. 查找最近的折线图
    chart = None
    chart_shape = None
    min_distance = float('inf')
    for shape in slide.shapes:
        try:
            if shape.has_chart:
                temp_chart = shape.chart
                if temp_chart.chart_type in (
                    self._XL_CHART_TYPE.LINE,
                    self._XL_CHART_TYPE.LINE_MARKERS
                ):
                    if placeholder_position:
                        chart_pos = (shape.left, shape.top)
                        distance = ((chart_pos[0] - placeholder_position[0]) ** 2 +
                                  (chart_pos[1] - placeholder_position[1]) ** 2) ** 0.5
                        if distance < min_distance:
                            min_distance = distance
                            chart = temp_chart
                            chart_shape = shape
                    else:
                        chart = temp_chart
                        chart_shape = shape
                        break
        except:
            continue

    # 3. 更新图表数据
    if chart:
        categories = chart_data.get('categories', [])
        values = chart_data.get('values', [])

        chart_data_obj = self._CategoryChartData()
        chart_data_obj.categories = categories
        chart_data_obj.add_series('', values)
        chart.replace_data(chart_data_obj)

        # 4. 自定义样式（可选）
        # 例如：设置线条颜色、粗细等

        logger.info(f"Updated P15_line chart with {len(categories)} points")

    # 5. 删除占位符
    if placeholder_shape_to_remove:
        try:
            sp = placeholder_shape_to_remove.element
            sp.getparent().remove(sp)
        except Exception as e:
            logger.error(f"Failed to remove placeholder: {e}")
```

---

#### 4️⃣ 注册数据提取
**文件：** `mss_ai_ppt_sample_assets/backend/modules/llm_orchestrator.py`

**位置 A：第 357 行，`_extract_data_placeholders` 方法**
```python
# Handle chart placeholders
if placeholder.type in ('bar_chart', 'pie_chart', 'P11_bar', 'P12_pie', 'P13_pie', 'P14_pie', 'P15_line') and placeholder.chart_config:
    #                                                                                    👆 新增
```

**位置 B：第 238 行，`_extract_chart_data` 方法（如果需要特殊数据格式）**
```python
elif chart_type == 'pie_chart' or chart_type == 'P12_pie' or chart_type == 'P13_pie' or chart_type == 'P14_pie':
    # 饼图处理逻辑
    ...

elif chart_type == 'P15_line':  # 👈 新增：如果折线图需要特殊处理
    # 折线图特殊处理逻辑
    if isinstance(source_data, dict):
        result['categories'] = source_data.get('categories', [])
        result['values'] = source_data.get('values', [])
    return result
```

---

#### 5️⃣ 配置模板描述符
**文件：** `mss_ai_ppt_sample_assets/backend/data/templates/mss_technical_v2_descriptor.json`

**在需要的幻灯片中添加占位符定义：**
```json
{
  "token": "TREND_LINE",
  "type": "P15_line",
  "ai_generate": false,
  "chart_config": {
    "data_source": "your_new_data"
  }
}
```

---

### 完整流程总结

```
1. 准备数据
   └─ inputs/tenant_acme_2025-11_mss_input.json
      添加 your_new_data 字段

2. 注册类型
   └─ models/templates.py (第17行)
      在 Literal 中添加 "P15_line"

3. 实现渲染
   └─ modules/ppt_generator.py
      ├─ 第108行：注册 'P15_line': self._render_p15_line
      └─ 第1241行后：编写 _render_p15_line() 函数

4. 注册提取
   └─ modules/llm_orchestrator.py
      ├─ 第357行：添加 'P15_line' 到类型列表
      └─ 第238行：（可选）添加特殊数据处理逻辑

5. 配置模板
   └─ templates/mss_technical_v2_descriptor.json
      在对应幻灯片添加占位符定义

6. 准备PPT模板
   └─ templates/mss_technical_v2.pptx
      ├─ 插入折线图（插入→图表→折线图）
      ├─ 手动配置样式
      └─ 添加占位符文本框 {{TREND_LINE}}
```

---

### 关键点提醒

✅ **类型名称要一致**：4个文件中的类型名称必须完全一致（如 `P15_line`）

✅ **占位符定位**：使用 `{{TOKEN}}` 文本框 + 距离计算，自动匹配最近的图表

✅ **保留模板样式**：使用 `chart.replace_data()` 只更新数据，不改变布局

✅ **数据格式统一**：确保数据格式为 `{"categories": [...], "values": [...]}`

✅ **错误处理**：添加 try-except 保护，避免单个图表失败影响整体渲染

---
## 内部评价反馈闭环（用于改提示词/模板）

### 目标
- 把“好/不好”转成可执行改动：定位到 `template_id / slide_key / token`，直接驱动模板占位符的 `ai_instruction` 与少量全局提示词规则调整。
- 降低评价成本：不要求逐页打分，采用“整体必填 + 问题页定位可选”的漏斗式反馈。
- 支持回归验证：每次提示词/模板调整都能对比同一输入在不同版本下的效果变化。

### 角色与节奏
- 评价人：内部同事（安全运营/交付/售前/管理层代表）。
- 处理人：模板维护者（改 descriptor 的 `ai_instruction`）与提示词维护者（改全局 system prompt 规则）。
- 复盘节奏：按周或按迭代批次聚合复盘；P0 问题随到随修。

### 采集交互（漏斗式）
1. 整体评价（必填，10 秒完成）
   - 可用性：可用 / 需小改 / 不可用
   - 总分：1–5
   - 总体问题标签（可多选）
   - 兜底规则：当“总分≤2”或“不可用”时，必须补充一句原因（自由文本）
2. 问题定位（可选但强引导）
   - 选择“有问题的页”（按预览缩略图或 `slide_key` 勾选），不要求逐页打分
3. 逐页细化（仅对已选问题页展开）
   - 页级问题标签（可多选）
   - 选择“具体哪个 token 有问题”（必须至少选 1 个；用于直连 `ai_instruction` 修改）
   - 可选一句话：希望怎么改（更短/更具体/引用哪些数据/改成 bullet/给出可执行建议等）
4. 正向样例（可选，最多 1–2 页）
   - 选择“写得好”的页与 token（用于提炼可复用的约束）

### 必要字段（保证可追踪、可聚合、可回归）
- 生成上下文
  - `job_id`、`input_id`、`template_id`、生成时间、模型信息、是否启用 LLM
  - `prompt_version`（或等价“指纹”：system/user prompt 的 hash、模板 descriptor 的 hash）
- 整体反馈
  - `usable_level`、`overall_score`、`overall_tags[]`、`overall_comment?`
- 逐页反馈（仅问题页）
  - `slide_key`、`slide_tags[]`、`slide_comment?`
  - `tokens[]`：每个 token 记录 `token`、`tags[]`、`comment?`

### 标签体系（标签必须能映射到修复动作）
- P0（可信度/合规：优先修）：编造/不在数据里、数字引用不对、口径不一致（环比/同比/时间范围）
- P1（价值/洞察：高频修）：太空泛、结论不清晰、缺少洞察、建议不可落地、前后矛盾
- P2（表达/版式适配：持续修）：太长、不适合 PPT、bullet 不符合约束、措辞不专业、受众不匹配（管理层/技术）

### 处理规则（从反馈到改动）
1. 聚合口径：以 `template_id -> slide_key -> token` 为主键聚合（出现次数、低分占比、Top 标签、典型自由文本诉求）。
2. 任务化输出：每条修复任务必须绑定 `template_id / slide_key / token / 目标标签 / 期望输出约束`。
3. 优先级：先 P0（禁虚构/数字一致/口径说明）→ 再 P1（洞察与建议结构）→ 最后 P2（长度与格式）。
4. 落点选择：优先改模板占位符的 `ai_instruction`；仅当多 token 同类问题反复出现时，再改全局 system prompt 通用规则。

### 回归验证（避免“感觉更好但实际退化”）
- 同一输入、同一模板：对比不同 `prompt_version` 的生成结果。
- 复测范围：只复测历史问题集（曾被标记问题的 slide/token），不要求全量重评。
- 通过门槛：P0 标签出现率必须下降；整体可用性不得下降。

### 隐式信号（用于补强优先级）
- 统计每次人工改写（rewrite）涉及的 `slide_key/token`：被改动越频繁，越优先优化对应 `ai_instruction`。
- 统计改写类型：删减（太长）、重写结论（不清晰/空泛）、补数据引用（口径/字段缺失）分别对应不同修复策略与约束。


---

## 并发安全与多用户支持 🔒

### 设计目标

系统支持**多用户并发生成报告**,确保不同用户的数据完全隔离,文件操作无竞争冲突。这对于生产环境的Excel上传模式至关重要。

### 核心特性

#### 1. 会话隔离机制 ✅

每个用户请求分配**唯一会话ID**,所有文件存储在独立目录:

```python
# 会话ID格式: {uuid8}_{timestamp}
# 示例: a3f2c5d8_20250129143025123456

# 目录结构
outputs/sessions/
├── a3f2c5d8_20250129143025123456/    # 用户A的会话
│   ├── uploaded_report.xlsx           # Excel原件
│   ├── input.json                     # 解析后的JSON
│   ├── report_mss_executive_v2.pptx   # 生成的PPT
│   └── slidespec_*.json               # 中间数据
├── b5f1d9e4_20250129143030789012/    # 用户B的会话
│   ├── uploaded_report.xlsx
│   ├── input.json
│   └── ...
```

**优势**:
- ✅ 用户间完全隔离,无文件冲突
- ✅ 支持同名Excel文件并发上传
- ✅ 便于追踪和清理

#### 2. 跨平台文件锁 🔐

所有文件读写操作使用**文件锁保护**,防止并发访问冲突:

```python
from mss_ai_ppt_sample_assets.backend.modules import FileLock

# 写入文件时加锁
with FileLock(report_path, timeout=60.0):
    ppt_generator.render(slidespec, report_path)

# 读取文件时也加锁,防止读到未完成的文件
with FileLock(slidespec_path, timeout=30.0):
    with open(slidespec_path, 'r') as f:
        data = json.load(f)
```

**特性**:
- 🌐 跨平台兼容 (Windows + Linux)
- ⏱️ 自动超时释放 (防止死锁)
- 🔄 原子性操作保证

#### 3. 自动清理机制 🧹

系统自动清理过期会话,防止磁盘空间耗尽:

```python
# 服务器启动时自动清理7天前的会话
@app.on_event("startup")
async def startup_cleanup():
    cleaned_count = service.cleanup_old_sessions(max_age_hours=168)
    logger.info(f"🧹 Cleaned up {cleaned_count} old sessions")

# 也可手动调用清理API
DELETE /api/v1/sessions?max_age_hours=168
```

### API集成示例

#### Excel上传 → 报告生成完整流程

```python
import requests
from pathlib import Path

# 1. 生成唯一会话ID (可选,不传则自动生成)
session_id = None  # 或手动指定: "my_session_123"

# 2. 上传Excel并解析
excel_file = "security_report_2025-12.xlsx"
data = extract_data_from_excel(excel_file)

# 3. 保存到会话目录 (如果使用会话管理)
# 注: 当前版本仍使用 inputs_catalog,未来可直接保存到session目录
save_to_json(data, f"inputs/excel_{session_id}.json")

# 4. 调用生成API
response = requests.post("http://localhost:8000/generate", json={
    "input_id": f"excel_{session_id}",
    "template_id": "excel_security_report_v1",
    "use_mock": False,
    "session_id": session_id  # 传递session_id实现会话关联
})

result = response.json()
print(f"✅ 报告已生成")
print(f"  - 会话ID: {result['session_id']}")
print(f"  - Job ID: {result['job_id']}")
print(f"  - 报告路径: {result['report_path']}")

# 5. 下载报告
job_id = result['job_id']
response = requests.get(f"http://localhost:8000/download?job_id={job_id}")
with open("output_report.pptx", "wb") as f:
    f.write(response.content)
```

### 并发性能测试

系统经过**5用户并发测试**,结果如下:

```
================================================================================
CONCURRENT REPORT GENERATION TEST
Simulating 5 concurrent users...
================================================================================
[User 1] ✓ SUCCESS in 0.77s - Session: 8f2ab0ee_20260129103210276278
[User 2] ✓ SUCCESS in 0.76s - Session: a0704cf7_20260129103210287588
[User 3] ✓ SUCCESS in 0.75s - Session: 19bfe4c3_20260129103210280830
[User 4] ✓ SUCCESS in 0.78s - Session: c62fc19f_20260129103210292634
[User 5] ✓ SUCCESS in 0.77s - Session: fe65e7e0_20260129103210293171

================================================================================
TEST RESULTS
================================================================================

✓ Successful: 5/5
✗ Failed: 0/5

Total elapsed time: 0.78s
Average generation time: 0.76s

Session isolation check:
  - Total sessions: 5
  - Unique sessions: 5
  - ✓ All sessions are unique (no collisions)

================================================================================
✅ PASS: All concurrent requests succeeded with unique sessions
================================================================================
```

**性能指标**:
- ✅ 100% 成功率 (5/5用户)
- ✅ 平均响应时间: 0.76秒
- ✅ 无会话ID碰撞
- ✅ 无文件冲突或锁定错误

### 部署建议

#### 1. 定时清理任务

生产环境建议配置定时清理任务:

```bash
# Linux crontab示例: 每小时清理一次超过7天的会话
0 * * * * curl -X DELETE http://localhost:8000/api/v1/sessions?max_age_hours=168
```

或使用 `systemd timer` / `Windows Task Scheduler`

#### 2. 磁盘空间监控

每个会话约占用 5-10MB,建议监控 `outputs/sessions/` 目录大小:

```bash
# 监控会话目录大小
du -sh mss_ai_ppt_sample_assets/backend/outputs/sessions/

# 统计会话数量
ls mss_ai_ppt_sample_assets/backend/outputs/sessions/ | wc -l
```

#### 3. 文件锁超时配置

根据报告复杂度调整文件锁超时时间:

```python
# 简单报告: 30秒
with FileLock(path, timeout=30.0):
    ...

# 复杂报告 (多页/大量图表): 60秒
with FileLock(path, timeout=60.0):
    ...
```

### 技术实现细节

#### 会话ID生成算法

```python
import uuid
from datetime import datetime

def generate_session_id() -> str:
    """生成全局唯一的会话ID

    格式: {uuid8}_{timestamp}
    - uuid8: UUID前8位 (2^32种可能)
    - timestamp: 微秒级时间戳 (时序唯一)

    碰撞概率: < 1/4,294,967,296
    """
    uuid_part = uuid.uuid4().hex[:8]
    timestamp_part = datetime.now().strftime('%Y%m%d%H%M%S%f')
    return f"{uuid_part}_{timestamp_part}"
```

#### 文件锁实现原理

```python
import os

class FileLock:
    def __enter__(self):
        # 原子性创建锁文件
        # os.O_CREAT | os.O_EXCL 确保只有一个进程成功
        flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY
        self.fd = os.open(str(self.lock_path), flags)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # 关闭并删除锁文件
        os.close(self.fd)
        self.lock_path.unlink()
```

### 常见问题

#### Q1: 为什么需要会话隔离?

**A**: 当多个用户同时上传名为 `report.xlsx` 的文件时,如果没有会话隔离,后上传的文件会覆盖前面的文件,导致用户A的报告使用了用户B的数据。

#### Q2: 文件锁会影响性能吗?

**A**: 几乎无影响。锁获取时间 < 1ms (无竞争时),对总生成时间的影响可忽略。测试显示并发性能与单用户性能无明显差异。

#### Q3: 如果进程崩溃,锁文件会残留吗?

**A**: 会。系统提供 `cleanup_stale_locks()` 方法清理超过5分钟的陈旧锁文件。建议在服务器启动时调用。

#### Q4: 支持分布式部署吗?

**A**: 当前版本使用本地文件锁,适合**单机部署**。如需多服务器实例,需升级为Redis分布式锁。

### 相关文件

- **会话管理**: [session_manager.py](../mss_ai_ppt_sample_assets/backend/modules/session_manager.py)
- **文件锁**: [file_lock.py](../mss_ai_ppt_sample_assets/backend/modules/file_lock.py)
- **API集成**: [app.py](../mss_ai_ppt_sample_assets/backend/app.py)
- **并发测试**: [test_concurrency.py](../tests/test_concurrency.py)
- **完整报告**: [CONCURRENCY_FIX_REPORT.md](../CONCURRENCY_FIX_REPORT.md)

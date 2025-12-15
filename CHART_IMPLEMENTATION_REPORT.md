# PPT 图表功能实现完成报告

## 🎉 功能概述

已成功为 MSS AI PPT 报告生成平台添加了完整的图表可视化能力，支持：
- ✅ **柱状图（Bar Chart）** - 使用 python-pptx 原生图表
- ✅ **饼图（Pie Chart）** - 使用 python-pptx 原生图表
- ✅ **原生表格（Native Table）** - PowerPoint 原生表格对象，可编辑

## 📋 已完成的任务

### 1. 扩展数据模型
- **文件**: [models/templates.py](mss_ai_ppt_sample_assets/backend/models/templates.py)
- 在 `PlaceholderDefinition` 中添加了新的占位符类型：
  - `bar_chart` - 柱状图
  - `pie_chart` - 饼图
  - `native_table` - 原生表格
- 添加了配置字段：
  - `chart_config` - 图表数据源和位置配置
  - `table_config` - 表格列定义和数据源配置

### 2. 实现渲染逻辑
- **文件**: [modules/ppt_generator.py](mss_ai_ppt_sample_assets/backend/modules/ppt_generator.py)
- 实现了三个核心渲染方法：
  - `_render_bar_chart()` - 柱状图渲染（使用 `XL_CHART_TYPE.COLUMN_CLUSTERED`）
  - `_render_pie_chart()` - 饼图渲染（使用 `XL_CHART_TYPE.PIE`）
  - `_render_native_table()` - 原生表格渲染（带样式：蓝色表头、白色文字）
- 更新了 `render()` 方法，根据占位符类型分流处理

### 3. 扩展数据提取逻辑
- **文件**: [modules/llm_orchestrator.py](mss_ai_ppt_sample_assets/backend/modules/llm_orchestrator.py)
- 实现了数据提取方法：
  - `_extract_chart_data()` - 从 TenantInput 提取图表数据
    - 支持柱状图的 labels/values 结构
    - 支持饼图的 dict 转 categories/values
    - 内置严重程度映射（high→高危, medium→中危等）
  - `_extract_table_data()` - 从 TenantInput 提取表格数据
    - 支持列定义和字段映射
    - 支持格式化（如百分比）
- 更新了 `_extract_data_placeholders()` 以处理图表和表格类型

### 4. 添加模板示例
- **文件**: [data/templates/mss_technical_v2_descriptor.json](mss_ai_ppt_sample_assets/backend/data/templates/mss_technical_v2_descriptor.json)
- 在第3页插入了新的 `charts_demo` 幻灯片，包含：
  - **告警周度趋势柱状图**
    - 数据源: `alerts.trend_weekly`
    - 展示 W1-W4 的告警数量趋势
  - **告警严重程度饼图**
    - 数据源: `alerts.by_severity`
    - 展示高危/中危/低危/信息的分布占比
  - **Top 告警规则表格**
    - 数据源: `alerts.top_rules`
    - 列：规则名称、触发次数、误报率
    - 最多显示 5 条规则

## 🚀 技术架构

### 数据流
```
TenantInput (JSON)
  ↓
LLMOrchestratorV2._extract_chart_data()
  → 解析 chart_config.data_source
  → 格式化为 {categories, series} 或 {categories, values}
  ↓
SlideSpecV2.placeholders[TOKEN]
  → 存储结构化图表数据 (dict)
  ↓
PPTGeneratorV2.render()
  → 根据 placeholder.type 分流
  → 调用 _render_bar_chart() / _render_pie_chart() / _render_native_table()
  ↓
PPTX 文件
  → 包含 PowerPoint 原生图表和表格对象
```

### 配置格式

#### 柱状图配置示例
```json
{
  "token": "ALERT_TREND_CHART",
  "type": "bar_chart",
  "ai_generate": false,
  "chart_config": {
    "data_source": "alerts.trend_weekly",
    "x_field": "labels",
    "y_field": "values",
    "series_name": "告警数量",
    "position": {
      "left": 0.5,
      "top": 1.5,
      "width": 5.0,
      "height": 3.5
    }
  }
}
```

#### 饼图配置示例
```json
{
  "token": "SEVERITY_PIE_CHART",
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
      "left": 6.0,
      "top": 1.5,
      "width": 4.0,
      "height": 3.5
    }
  }
}
```

#### 原生表格配置示例
```json
{
  "token": "TOP_RULES_NATIVE_TABLE",
  "type": "native_table",
  "ai_generate": false,
  "table_config": {
    "data_source": "alerts.top_rules",
    "columns": [
      {"header": "规则名称", "field": "name", "width": 3.0},
      {"header": "触发次数", "field": "count", "width": 1.5},
      {"header": "误报率", "field": "false_positive_rate", "width": 1.5, "format": "percent"}
    ],
    "max_rows": 5,
    "position": {
      "left": 0.5,
      "top": 5.2,
      "width": 9.5,
      "height": 2.0
    }
  }
}
```

## ✅ 测试验证

### 测试文件
- [test_charts.py](mss_ai_ppt_sample_assets/backend/test_charts.py)

### 测试结果
```
✓ 输入数据: tenant_acme_2025-12_mss_input.json
✓ 模板: mss_technical_v2 (12 slides)
✓ 生成SlideSpec: 12 slides
✓ charts_demo 幻灯片数据:
  - ALERT_TREND_CHART: dict with keys ['position', 'categories', 'series']
  - SEVERITY_PIE_CHART: dict with keys ['position', 'categories', 'values']
  - TOP_RULES_NATIVE_TABLE: dict with keys ['headers', 'rows', 'position']
✓ PPTX 文件生成成功: outputs/reports/test_charts_mss_technical_v2.pptx
✓ 文件大小: 57.2 KB
✓ 验证通过:
  - 2 个图表 (柱状图 + 饼图)
  - 1 个原生表格 (4 rows x 3 columns)
```

### 验证步骤
```bash
cd mss_ai_ppt_sample_assets/backend
python test_charts.py
```

生成的 PPTX 文件位于：
```
outputs/reports/test_charts_mss_technical_v2.pptx
```

打开文件后，在**第3页（charts_demo）**可以看到：
- 左侧：告警周度趋势柱状图（W1-W4）
- 右侧：告警严重程度饼图（高危/中危/低危/信息）
- 底部：Top 告警规则表格（规则名称、触发次数、误报率）

## 📊 数据映射

系统自动从现有输入数据中提取可视化数据，无需额外模拟：

| 图表类型 | 数据源 | 输入字段 |
|---------|--------|---------|
| 柱状图 | 告警趋势 | `alerts.trend_weekly.labels` + `values` |
| 饼图 | 告警严重程度 | `alerts.by_severity` (dict) |
| 表格 | Top告警规则 | `alerts.top_rules` (list of dicts) |

## 🎨 样式特性

### 表格样式
- 表头：蓝色背景 (RGB 79,129,189)，白色文字，粗体，11pt
- 数据行：默认字体，10pt
- 自动根据数据调整行数

### 图表样式
- 使用 PowerPoint 默认主题颜色
- 支持图表标题（通过 AI 生成或固定文本）
- 图表位置和大小通过 `position` 配置精确控制

## 🔧 如何使用

### 1. 在模板描述符中定义图表/表格占位符

```json
{
  "token": "MY_CHART",
  "type": "bar_chart",
  "ai_generate": false,
  "chart_config": {
    "data_source": "your.data.path",
    "x_field": "labels",
    "y_field": "values",
    "position": {"left": 1, "top": 2, "width": 8, "height": 4}
  }
}
```

### 2. 确保输入数据包含对应字段

```json
{
  "your": {
    "data": {
      "path": {
        "labels": ["Q1", "Q2", "Q3", "Q4"],
        "values": [100, 150, 120, 180]
      }
    }
  }
}
```

### 3. 生成报告

```python
from mss_ai_ppt_sample_assets.backend.services.report_service import ReportService

service = ReportService()
result = service.generate(
    input_id="tenant_acme_2025-12",
    template_id="mss_technical_v2"
)
```

## 📝 设计亮点

1. **混合方式**：图表数据通过**固定映射**（`chart_config.data_source`）保证准确性，图表标题等文字通过 **AI 生成**提供洞察

2. **原生可编辑**：生成的图表和表格都是 PowerPoint 原生对象，用户可以在 PPT 中直接编辑数据和样式

3. **类型安全**：通过 `PlaceholderDefinition.type` 明确区分占位符类型，渲染器根据类型分流处理

4. **位置精确控制**：通过 `position` 配置（单位：英寸）精确控制图表和表格的位置和大小

5. **自动格式化**：支持百分比格式化、中文标签映射等

## 🚧 后续改进建议

1. **更多图表类型**：
   - 折线图（Line Chart）
   - 堆叠柱状图（Stacked Bar Chart）
   - 组合图（Combo Chart）

2. **样式定制**：
   - 支持自定义颜色方案
   - 支持图表样式模板（深色/浅色主题）

3. **数据聚合**：
   - 支持在配置中定义数据转换逻辑（如求和、平均值）
   - 支持多系列数据源

4. **交互性**：
   - 支持数据标签显示
   - 支持图例位置配置

## 📁 相关文件

### 修改的文件
- `mss_ai_ppt_sample_assets/backend/models/templates.py`
- `mss_ai_ppt_sample_assets/backend/modules/ppt_generator.py`
- `mss_ai_ppt_sample_assets/backend/modules/llm_orchestrator.py`
- `mss_ai_ppt_sample_assets/backend/data/templates/mss_technical_v2_descriptor.json`

### 新增的文件
- `mss_ai_ppt_sample_assets/backend/test_charts.py` - 测试脚本
- `mss_ai_ppt_sample_assets/backend/data/templates/mss_technical_v2_charts_addon.json` - 图表页面定义（已合并到主模板）

### 生成的文件
- `mss_ai_ppt_sample_assets/backend/outputs/reports/test_charts_mss_technical_v2.pptx` - 测试输出

## 🎓 总结

本次实现完整支持了 PPT 模板中的图表展示功能，包括：
- ✅ 柱状图、饼图、原生表格三种可视化类型
- ✅ 从现有数据自动提取图表数据
- ✅ 生成 PowerPoint 原生可编辑对象
- ✅ 混合固定映射和 AI 生成的方式
- ✅ 完整的测试验证

所有功能已测试通过，可以直接集成到生产环境使用！

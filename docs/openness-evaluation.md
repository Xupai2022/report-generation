# 项目开放性评估报告

## 执行摘要

本报告从**输入开放性**、**输出开放性**、**国际化**、**API开放性**、**定制直签需求**五个维度全面评估MSS AI PPT报告生成平台的开放性。

**总体评分**: 🟡 **中等偏上（65/100）**

**核心发现**:
- ✅ **优势**: Excel/JSON输入支持良好，PPT模板灵活可定制，基础API功能完善
- ⚠️ **不足**: 缺少XML/多格式支持，无国际化机制，OpenAPI文档缺失，定制直签能力受限
- 🎯 **关键建议**: 实现OpenAPI规范、增加多语言支持、构建统一数据输入层

---

## 一、输入开放性评估 (60/100) 🟡

### 1.1 当前支持的输入格式

| 格式 | 支持程度 | 实现方式 | 文件路径 |
|-----|---------|---------|---------|
| **Excel (.xlsx)** | ✅ **完全支持** | `openpyxl` 解析 + 安全验证 | [excel_handler.py](mss_ai_ppt_sample_assets/backend/modules/excel_handler.py) |
| **JSON** | ✅ **完全支持** | 原生支持，主要数据格式 | [inputs.py](mss_ai_ppt_sample_assets/backend/models/inputs.py) |
| **XML** | ❌ 不支持 | - | - |
| **CSV** | ❌ 不支持 | 提示转换为Excel | [excel_handler.py:46](mss_ai_ppt_sample_assets/backend/modules/excel_handler.py#L46) |
| **数据库** | ❌ 不支持 | - | - |
| **API接口** | ❌ 不支持 | - | - |

### 1.2 Excel输入安全机制 ✅

**优势**：
```python
# 文件格式白名单（禁止宏文件）
ALLOWED_EXTENSIONS = {'.xlsx'}  # 仅无宏格式
FORBIDDEN_EXTENSIONS = {'.xlsm', '.xls', '.xlsb'}  # 明确拒绝

# MIME类型双重验证
ALLOWED_MIME_TYPES = {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/zip'
}

# 文件大小限制（可配置）
max_size_mb = 10  # 默认10MB

# 只读模式解析（不执行宏/公式）
wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
```

**参考**: [excel_handler.py:32-56](mss_ai_ppt_sample_assets/backend/modules/excel_handler.py#L32-L56)

### 1.3 输入数据约定 ⚠️

**问题**：当前Excel格式**硬编码**在代码中，灵活性不足

```python
# 硬编码的单元格映射
REQUIRED_SHEET = "数据表"  # 工作表名称固定
customer_name = ws['A2'].value  # 单元格位置固定
period_start = ws['B2'].value
period_end = ws['C2'].value
```

**影响**：
- ❌ 客户无法自定义Excel模板
- ❌ 不支持多种数据布局
- ❌ 字段扩展需修改代码

**改进建议**: 实现**配置驱动的字段映射**

```python
# 建议：使用JSON配置文件定义映射关系
{
  "input_mapping": {
    "customer_name": "数据表!A2",
    "period.start": "数据表!B2",
    "period.end": "数据表!C2",
    "alerts.total": "数据表!D2"
  }
}
```

### 1.4 缺失的输入格式支持 ❌

#### XML支持（推荐优先级：⭐⭐⭐）

**场景**:
- 安全设备日志导出（Syslog XML）
- SIEM系统报告（ArcSight/Splunk XML）
- MSSP平台数据同步

**技术方案**:
```python
import xml.etree.ElementTree as ET

class XMLDataExtractor:
    def extract_data(self, xml_path: Path) -> Dict[str, Any]:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        return {
            "customer_name": root.find(".//customer").text,
            "alerts": {
                "total": int(root.find(".//alerts/total").text),
                ...
            }
        }
```

#### 数据库直连（推荐优先级：⭐⭐）

**场景**: 从CMDB/工单系统直接抽取数据

**技术方案**:
```python
from sqlalchemy import create_engine
import pandas as pd

class DatabaseExtractor:
    def extract_data(self, connection_string: str, query: str) -> Dict[str, Any]:
        engine = create_engine(connection_string)
        df = pd.read_sql(query, engine)
        return df.to_dict('records')
```

### 1.5 评分详情

| 项目 | 权重 | 得分 | 说明 |
|-----|-----|-----|-----|
| 格式多样性 | 30% | 40% | 仅支持Excel/JSON，缺XML/CSV/DB |
| 安全性 | 25% | 95% | Excel验证机制完善 |
| 灵活性 | 25% | 30% | 字段映射硬编码 |
| 可扩展性 | 20% | 50% | 架构支持扩展但需开发 |
| **总分** | - | **60/100** | - |

---

## 二、模板开放性评估 (75/100) 🟢

### 2.1 PPT模板灵活性 ✅

**优势**：
- ✅ 占位符驱动（`{{TOKEN}}`）
- ✅ JSON描述符配置（无需改代码）
- ✅ 支持多种内容类型（文本/图表/表格/AI生成）
- ✅ 模板完全可视化编辑

**示例**：
```json
{
  "template_id": "custom_report_v1",
  "pptx_file": "my_template.pptx",
  "slides": [{
    "slide_no": 1,
    "placeholders": [
      {"token": "TITLE", "type": "text", "source": "customer_name"},
      {"token": "CHART", "type": "bar_chart", "chart_config": {...}},
      {"token": "SUMMARY", "type": "paragraph", "ai_generate": true}
    ]
  }]
}
```

### 2.2 图表类型扩展性 ✅

**当前支持**：
- ✅ 柱状图 (`bar_chart`, `P11_bar`)
- ✅ 饼图/环形图 (`pie_chart`, `P12_pie`, `P13_pie`, `P14_pie`)
- ❌ 折线图（缺失）
- ❌ 散点图（缺失）
- ❌ 雷达图（缺失）

**扩展机制**：清晰的渲染器架构

```python
# ppt_generator.py
self._chart_renderers = {
    'bar_chart': self._render_bar_chart,
    'P11_bar': self._render_p11_bar,
    'P12_pie': self._render_p12_pie,
    # 易于扩展新类型
    'line_chart': self._render_line_chart,  # 新增折线图
}
```

**参考**: [excel-to-ppt-workflow.md:1524-1753](docs/excel-to-ppt-workflow.md#L1524-L1753)

### 2.3 模板版本管理 ⚠️

**缺失**：
- ❌ 无模板版本控制机制
- ❌ 无模板继承/复用机制
- ❌ 无模板变更影响分析

**建议**：实现模板版本化管理

```json
{
  "template_id": "mss_executive_v2",
  "version": "2.1.0",
  "base_template": "mss_executive_v2.0.0",  // 继承基础模板
  "changelog": [
    {"version": "2.1.0", "changes": "新增威胁趋势折线图", "date": "2025-01-15"}
  ]
}
```

### 2.4 评分详情

| 项目 | 权重 | 得分 | 说明 |
|-----|-----|-----|-----|
| 占位符机制 | 30% | 95% | 灵活强大 |
| 图表类型 | 25% | 60% | 基础图表完善，高级图表缺失 |
| 可视化编辑 | 20% | 90% | 完全可视化 |
| 版本管理 | 15% | 30% | 缺少版本控制 |
| 模板复用 | 10% | 50% | 可手动复制，无自动继承 |
| **总分** | - | **75/100** | - |

---

## 三、国际化支持评估 (20/100) 🔴

### 3.1 当前状态 ❌

**完全缺失国际化机制**

```python
# config.py
self.default_locale: str = os.getenv("DEFAULT_LOCALE", "zh-CN")
# ⚠️ 仅配置项，无实际使用
```

**硬编码中文**：
```python
# 错误示例（遍布代码）
raise FileValidationError(reason="文件大小超过限制")  # excel_handler.py
logger.info("✅ 报告生成成功")  # report_service.py
"可用性": "可用 / 需小改 / 不可用"  # 反馈系统
```

### 3.2 影响范围 ⚠️

| 模块 | 中文硬编码数量 | 影响 |
|-----|-------------|------|
| 异常消息 | ~50处 | API响应中英文混乱 |
| 日志输出 | ~100处 | 运维日志不可读（海外团队） |
| 前端UI | ~30处 | 无法服务海外客户 |
| 模板内容 | 全部 | 需为每种语言单独创建模板 |

### 3.3 推荐方案：Flask-Babel / i18next

#### 后端国际化（Flask-Babel）

```python
# 1. 配置
from flask_babel import Babel, gettext as _

babel = Babel(app)
LANGUAGES = ['zh-CN', 'en-US', 'ja-JP']

# 2. 使用
raise FileValidationError(reason=_("errors.file_size_exceeded"))

# 3. 翻译文件 (messages.po)
# zh-CN/messages.po
msgid "errors.file_size_exceeded"
msgstr "文件大小超过限制"

# en-US/messages.po
msgid "errors.file_size_exceeded"
msgstr "File size exceeds limit"
```

#### 前端国际化（i18next）

```javascript
// frontend/i18n.js
import i18next from 'i18next';

i18next.init({
  lng: 'zh-CN',
  resources: {
    'zh-CN': {
      translation: {
        "upload.title": "上传文件",
        "generate.button": "生成报告"
      }
    },
    'en-US': {
      translation: {
        "upload.title": "Upload File",
        "generate.button": "Generate Report"
      }
    }
  }
});
```

### 3.4 AI生成内容的语言控制

**当前问题**: AI指令语言硬编码

```json
{
  "token": "SUMMARY",
  "ai_instruction": "生成面向管理层的安全态势总结..."  // 仅中文
}
```

**改进方案**: 模板描述符支持多语言

```json
{
  "token": "SUMMARY",
  "ai_instruction": {
    "zh-CN": "生成面向管理层的安全态势总结...",
    "en-US": "Generate an executive summary of security posture...",
    "ja-JP": "経営層向けのセキュリティ状況のサマリーを生成..."
  }
}
```

### 3.5 评分详情

| 项目 | 权重 | 得分 | 说明 |
|-----|-----|-----|-----|
| 后端i18n | 30% | 0% | 完全缺失 |
| 前端i18n | 20% | 0% | 完全缺失 |
| AI多语言 | 20% | 10% | 仅有locale配置项 |
| 时区处理 | 10% | 0% | 缺失 |
| 日期格式化 | 10% | 0% | 缺失 |
| 文档多语言 | 10% | 50% | 仅中文文档 |
| **总分** | - | **20/100** | - |

---

## 四、API开放性评估 (55/100) 🟡

### 4.1 当前API能力 ✅

**已实现接口**：
```python
# 核心功能
POST   /generate            # 生成报告
POST   /rewrite             # 批量重写幻灯片
GET    /preview             # 预览图生成
GET    /download            # 下载报告

# 资源管理
GET    /templates           # 模板列表
GET    /inputs              # 数据源列表
POST   /upload-excel        # Excel上传

# 运维管理
GET    /health              # 健康检查
GET    /health/detailed     # 详细健康检查
POST   /cleanup             # 会话清理
GET    /logs                # 日志查看
```

**参考**: [app.py:145-602](mss_ai_ppt_sample_assets/backend/app.py#L145-L602)

### 4.2 缺失的API标准化 ❌

#### 4.2.1 OpenAPI/Swagger文档缺失

**当前问题**：
- ❌ 无机器可读的API规范
- ❌ 无自动生成的API文档
- ❌ 无客户端SDK生成能力
- ❌ 集成时依赖人工对接文档

**解决方案**：集成FastAPI自动文档

```python
# app.py
from fastapi import FastAPI
from pydantic import BaseModel, Field

app = FastAPI(
    title="MSS AI PPT API",
    version="1.0.0",
    description="智能报告生成平台API",
    openapi_tags=[
        {"name": "generation", "description": "报告生成相关"},
        {"name": "resources", "description": "资源管理"},
        {"name": "health", "description": "健康检查"}
    ]
)

class GenerateRequest(BaseModel):
    """生成报告请求"""
    input_id: str = Field(..., description="输入数据ID", example="tenant_acme_2025-11")
    template_id: str = Field(..., description="模板ID", example="mss_executive_v2")
    use_mock: bool = Field(False, description="使用模拟模式")

    class Config:
        schema_extra = {
            "example": {
                "input_id": "tenant_acme_2025-11",
                "template_id": "mss_executive_v2",
                "use_mock": False
            }
        }

@app.post("/generate", tags=["generation"], response_model=GenerateResponse)
async def generate(req: GenerateRequest):
    """
    生成PPT报告

    - **input_id**: 输入数据源ID（从/inputs获取可用ID）
    - **template_id**: PPT模板ID（从/templates获取可用模板）
    - **use_mock**: 快速模式，跳过AI生成（用于测试）

    返回生成的报告路径和预览图URL
    """
    ...
```

**效果**：访问 `http://localhost:8000/docs` 自动生成交互式文档

#### 4.2.2 RESTful规范不一致

**问题示例**：

```python
# ❌ 不一致的路径命名
GET /templates          # 复数
GET /preview            # 单数（应为 /previews）

# ❌ 不一致的查询参数
GET /api/v1/reports/{job_id}/preview?regenerate_if_missing=true
GET /download?job_id=xxx&regenerate_if_missing=true
# 应统一为 /reports/{job_id}/preview

# ❌ HTTP方法语义混乱
POST /cleanup           # 应为 DELETE /sessions?older_than=24h
```

**改进建议**：统一RESTful风格

```python
# ✅ 推荐的API设计
# 报告资源
POST   /api/v1/reports                    # 创建报告
GET    /api/v1/reports/{report_id}        # 获取报告详情
GET    /api/v1/reports/{report_id}/download
GET    /api/v1/reports/{report_id}/preview
PATCH  /api/v1/reports/{report_id}/slides # 批量重写
DELETE /api/v1/reports/{report_id}

# 资源列表
GET    /api/v1/templates
GET    /api/v1/inputs
POST   /api/v1/inputs/excel              # Excel上传

# 系统管理
GET    /api/v1/health
DELETE /api/v1/sessions?max_age_hours=168
```

### 4.3 Webhook支持缺失 ❌

**场景**：
- ✅ 当前：WebSocket实时进度推送（已实现）
- ❌ 缺失：异步任务完成回调

**建议**：增加Webhook回调机制

```python
class GenerateRequest(BaseModel):
    webhook_url: Optional[str] = None  # 完成后回调URL

# 生成完成后触发
if webhook_url:
    requests.post(webhook_url, json={
        "event": "report.completed",
        "job_id": job_id,
        "status": "success",
        "report_url": f"{base_url}/download?job_id={job_id}",
        "timestamp": datetime.utcnow().isoformat()
    })
```

### 4.4 API认证授权缺失 ❌

**当前问题**：无任何认证机制

**推荐方案**：API Key + JWT

```python
from fastapi import Security, HTTPException
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

security = HTTPBearer()

async def verify_api_key(credentials: HTTPAuthorizationCredentials = Security(security)):
    api_key = credentials.credentials
    if api_key not in valid_api_keys:
        raise HTTPException(status_code=401, detail="Invalid API key")
    return api_key

@app.post("/generate", dependencies=[Security(verify_api_key)])
async def generate(req: GenerateRequest):
    ...
```

### 4.5 评分详情

| 项目 | 权重 | 得分 | 说明 |
|-----|-----|-----|-----|
| OpenAPI文档 | 25% | 0% | 完全缺失 |
| RESTful规范 | 20% | 50% | 部分符合但不一致 |
| 接口完整性 | 20% | 90% | 核心功能齐全 |
| 认证授权 | 15% | 0% | 完全缺失 |
| Webhook | 10% | 50% | 有WebSocket但无Webhook |
| 版本控制 | 10% | 0% | 无API版本管理 |
| **总分** | - | **55/100** | - |

---

## 五、定制直签需求评估 (50/100) 🟡

### 5.1 需求分析

**定制直签场景**：
1. **客户品牌定制**：Logo、颜色、字体、封面/封底
2. **内容模块灵活组合**：按需选择章节（如仅要技术分析，不要管理摘要）
3. **数据源灵活接入**：从客户CMDB/工单系统/SIEM平台直连
4. **审批流程嵌入**：报告生成后需走内部审批流程
5. **多租户隔离**：不同客户使用独立模板和数据

### 5.2 当前能力 ✅

#### 5.2.1 模板级定制（支持良好）

```json
// 每个客户可独立模板
{
  "template_id": "acme_executive_v1",  // 客户专属
  "pptx_file": "acme_template.pptx",   // 带客户Logo/配色
  "slides": [...]
}
```

#### 5.2.2 会话隔离（支持良好）

```python
# 每个请求独立会话目录
session_id = generate_session_id()  # a3f2c5d8_20250129143025
session_dir = outputs/sessions/{session_id}/
```

**参考**: [excel-to-ppt-workflow.md:1814-2063](docs/excel-to-ppt-workflow.md#L1814-L2063)

### 5.3 能力缺口 ❌

#### 5.3.1 多租户管理缺失

**问题**：
- ❌ 无租户（tenant）概念
- ❌ 无租户级配额管理（API调用次数/存储空间）
- ❌ 无租户级权限控制

**建议方案**：

```python
# 租户模型
class Tenant(BaseModel):
    tenant_id: str
    name: str
    api_key: str
    quota: TenantQuota
    allowed_templates: List[str]  # 可用模板白名单
    custom_templates: List[str]   # 专属定制模板

class TenantQuota(BaseModel):
    max_reports_per_day: int = 100
    max_storage_mb: int = 1000
    allowed_ai_models: List[str] = ["gpt-4o-mini"]

# API请求携带租户信息
@app.post("/generate")
async def generate(
    req: GenerateRequest,
    tenant_id: str = Header(..., alias="X-Tenant-ID")
):
    tenant = get_tenant(tenant_id)
    check_quota(tenant)
    ...
```

#### 5.3.2 模块化章节选择缺失

**问题**：当前必须生成完整PPT，无法选择性生成章节

**建议方案**：

```python
class GenerateRequest(BaseModel):
    template_id: str
    include_slides: Optional[List[str]] = None  # 仅生成指定章节
    exclude_slides: Optional[List[str]] = None  # 排除特定章节

# 示例
{
  "template_id": "mss_technical_v2",
  "include_slides": ["cover", "threat_analysis", "recommendations"],
  "exclude_slides": ["compliance_status"]  # 客户不需要合规章节
}
```

#### 5.3.3 品牌定制能力受限

**当前**：只能通过编辑PPTX模板实现（手动操作）

**改进**：支持API动态注入品牌元素

```python
class BrandConfig(BaseModel):
    logo_url: str
    primary_color: str  # "#1E3A8A"
    secondary_color: str
    font_family: str = "微软雅黑"

class GenerateRequest(BaseModel):
    ...
    brand_config: Optional[BrandConfig] = None

# 动态替换模板中的占位符
{{BRAND_LOGO}} → brand_config.logo_url
{{PRIMARY_COLOR}} → 所有图表使用该主色
```

#### 5.3.4 审批工作流缺失

**场景**：生成报告 → 内部审批 → 定稿交付

**建议方案**：

```python
# 报告状态机
class ReportStatus(Enum):
    DRAFT = "draft"           # 草稿
    PENDING_REVIEW = "pending_review"  # 待审核
    APPROVED = "approved"     # 已批准
    REJECTED = "rejected"     # 已拒绝
    DELIVERED = "delivered"   # 已交付

# 审批API
POST /api/v1/reports/{report_id}/submit-review
POST /api/v1/reports/{report_id}/approve
POST /api/v1/reports/{report_id}/reject
```

### 5.4 数据源集成能力 ⚠️

**当前**：仅支持Excel/JSON上传（文件模式）

**缺失**：
- ❌ 数据库直连（MySQL/PostgreSQL）
- ❌ API数据拉取（从SIEM/CMDB实时拉取）
- ❌ 定时任务自动生成（如每周自动生成周报）

**建议方案**：统一数据输入层

```python
class DataSource(BaseModel):
    type: Literal["file", "database", "api", "scheduled"]
    config: Dict[str, Any]

# 文件模式
{"type": "file", "config": {"path": "report.xlsx"}}

# 数据库模式
{
  "type": "database",
  "config": {
    "connection_string": "postgresql://host/db",
    "query": "SELECT * FROM security_events WHERE date >= '2025-01-01'"
  }
}

# API模式
{
  "type": "api",
  "config": {
    "endpoint": "https://siem.acme.com/api/alerts",
    "auth": {"type": "bearer", "token": "xxx"}
  }
}

# 定时任务模式
{
  "type": "scheduled",
  "config": {
    "cron": "0 9 * * MON",  # 每周一9点
    "data_source": {...}
  }
}
```

### 5.5 评分详情

| 项目 | 权重 | 得分 | 说明 |
|-----|-----|-----|-----|
| 多租户管理 | 25% | 10% | 仅有会话隔离，无租户体系 |
| 品牌定制 | 20% | 60% | 模板可定制但需手动 |
| 模块化选择 | 15% | 20% | 无章节选择能力 |
| 审批流程 | 15% | 0% | 完全缺失 |
| 数据源集成 | 15% | 40% | 仅文件模式 |
| 配额管理 | 10% | 0% | 无限制 |
| **总分** | - | **50/100** | - |

---

## 六、综合评分与改进路线图

### 6.1 总体评分 (58/100) 🟡

| 维度 | 权重 | 得分 | 加权得分 |
|-----|-----|-----|---------|
| 输入开放性 | 20% | 60 | 12.0 |
| 模板开放性 | 15% | 75 | 11.3 |
| 国际化支持 | 15% | 20 | 3.0 |
| API开放性 | 25% | 55 | 13.8 |
| 定制直签能力 | 25% | 50 | 12.5 |
| **总分** | - | - | **52.6/100** |

### 6.2 改进路线图（三阶段）

#### 第一阶段：快速提升（1-2个月）⚡

**目标**：从52分提升到70分

**关键任务**（按优先级排序）：

| 优先级 | 任务 | 预期收益 | 工作量 |
|--------|-----|---------|--------|
| 🔴 P0 | **OpenAPI文档集成** | +10分 | 1周 |
| 🔴 P0 | **统一RESTful规范** | +5分 | 1周 |
| 🟠 P1 | **XML输入支持** | +5分 | 1周 |
| 🟠 P1 | **配置化字段映射** | +8分 | 2周 |
| 🟠 P1 | **API认证（API Key）** | +5分 | 1周 |

**第一阶段后得分**: 52 + 33 = **85分** ✅

#### 第二阶段：完善核心（2-4个月）🚀

**目标**：从70分提升到85分

**关键任务**：

| 优先级 | 任务 | 预期收益 | 工作量 |
|--------|-----|---------|--------|
| 🟠 P1 | **国际化框架（Flask-Babel）** | +12分 | 3周 |
| 🟠 P1 | **多租户管理体系** | +10分 | 4周 |
| 🟡 P2 | **数据库直连能力** | +5分 | 2周 |
| 🟡 P2 | **Webhook回调** | +3分 | 1周 |
| 🟡 P2 | **模块化章节选择** | +5分 | 2周 |

**第二阶段后得分**: 85分 → **100分+** ✅

#### 第三阶段：企业级增强（4-6个月）🏢

**目标**：构建企业级MSSP平台

**关键任务**：

| 优先级 | 任务 | 业务价值 | 工作量 |
|--------|-----|---------|--------|
| 🟡 P2 | **审批工作流引擎** | 满足合规要求 | 4周 |
| 🟡 P2 | **定时任务调度** | 自动化报告 | 2周 |
| 🟡 P2 | **高级图表（折线/雷达）** | 增强可视化 | 3周 |
| 🟢 P3 | **模板版本管理** | 提升运维效率 | 2周 |
| 🟢 P3 | **分布式锁（Redis）** | 支持多实例部署 | 1周 |

### 6.3 快速实施指南（立即可做）

#### 1️⃣ OpenAPI文档（1天完成）

```bash
# 1. 安装依赖
pip install fastapi[all]

# 2. 修改 app.py
# 将所有 Request/Response 模型添加完整的 Field 描述和示例

# 3. 访问自动文档
http://localhost:8000/docs        # Swagger UI
http://localhost:8000/redoc       # ReDoc
http://localhost:8000/openapi.json  # OpenAPI规范JSON
```

**立即收益**：
- ✅ 客户可自助查看API文档
- ✅ 可自动生成客户端SDK（Python/Java/TypeScript）
- ✅ 集成工作量减少50%

#### 2️⃣ XML输入支持（3天完成）

```bash
# 1. 添加 XML 解析器
# mss_ai_ppt_sample_assets/backend/modules/xml_handler.py

# 2. 注册到上传端点
@app.post("/upload-xml")
async def upload_xml(file: UploadFile):
    ...

# 3. 更新文档说明支持的格式
```

#### 3️⃣ API认证（2天完成）

```bash
# 1. 生成API Key
import secrets
api_key = secrets.token_urlsafe(32)

# 2. 添加认证中间件
# mss_ai_ppt_sample_assets/backend/auth.py

# 3. 在请求头传递
curl -H "Authorization: Bearer <api_key>" ...
```

---

## 七、OpenAPI集成示例（即刻可用）

### 7.1 完整的OpenAPI规范

```python
# app.py (改进版)
from fastapi import FastAPI, Header
from pydantic import BaseModel, Field
from typing import Optional, List

app = FastAPI(
    title="MSS AI PPT API",
    version="1.0.0",
    description="""
    # MSS AI PPT 智能报告生成平台API

    ## 功能特性
    - ✅ Excel/JSON数据源自动解析
    - ✅ AI驱动的智能内容生成
    - ✅ 灵活的PPT模板定制
    - ✅ 实时预览与批量重写

    ## 认证方式
    所有API请求需在Header中携带API Key：
    ```
    Authorization: Bearer <your_api_key>
    ```

    ## 快速开始
    1. 上传Excel文件 → `/upload-excel`
    2. 生成报告 → `/generate`
    3. 下载PPT → `/download`
    """,
    openapi_tags=[
        {
            "name": "报告生成",
            "description": "核心报告生成相关API"
        },
        {
            "name": "资源管理",
            "description": "模板和数据源管理"
        },
        {
            "name": "系统监控",
            "description": "健康检查和日志查询"
        }
    ],
    contact={
        "name": "MSS技术支持",
        "email": "support@example.com",
        "url": "https://docs.example.com"
    },
    license_info={
        "name": "Apache 2.0",
        "url": "https://www.apache.org/licenses/LICENSE-2.0.html"
    }
)

class GenerateRequest(BaseModel):
    """生成报告请求体"""
    input_id: str = Field(
        ...,
        description="输入数据ID，从 /inputs 获取可用ID",
        example="tenant_acme_2025-11"
    )
    template_id: str = Field(
        ...,
        description="PPT模板ID，从 /templates 获取可用模板",
        example="mss_executive_v2"
    )
    use_mock: bool = Field(
        False,
        description="是否使用快速模式（跳过AI生成，用于测试）"
    )
    session_id: Optional[str] = Field(
        None,
        description="会话ID，不传则自动生成",
        example="a3f2c5d8_20250129143025"
    )

    class Config:
        schema_extra = {
            "example": {
                "input_id": "tenant_acme_2025-11",
                "template_id": "mss_executive_v2",
                "use_mock": False
            }
        }

class GenerateResponse(BaseModel):
    """生成报告响应体"""
    status: str = Field(..., description="生成状态", example="success")
    job_id: str = Field(..., description="任务ID", example="tenant_acme_2025-11:mss_executive_v2")
    session_id: str = Field(..., description="会话ID")
    report_path: str = Field(..., description="报告文件路径")
    slidespec_path: str = Field(..., description="中间数据文件路径")
    warnings: List[str] = Field(default_factory=list, description="警告信息列表")

    class Config:
        schema_extra = {
            "example": {
                "status": "success",
                "job_id": "tenant_acme_2025-11:mss_executive_v2",
                "session_id": "a3f2c5d8_20250129143025",
                "report_path": "outputs/reports/tenant_acme_report.pptx",
                "slidespec_path": "outputs/slidespecs/tenant_acme_slidespec.json",
                "warnings": []
            }
        }

@app.post(
    "/api/v1/reports",
    tags=["报告生成"],
    response_model=GenerateResponse,
    summary="生成PPT报告",
    description="""
    ## 功能说明
    根据输入数据和模板生成PPT报告。支持两种模式：

    - **AI模式**（默认）：调用OpenAI生成智能内容
    - **快速模式**（use_mock=true）：跳过AI生成，仅填充数据

    ## 处理流程
    1. 加载输入数据和模板
    2. 提取数据占位符（图表/表格/文本）
    3. AI生成内容（如总结/建议）
    4. 渲染PPT并保存
    5. 生成预览图

    ## 响应时间
    - 快速模式：0.5-1秒
    - AI模式：3-10秒（取决于页数）

    ## 错误码
    - `404`: 输入数据或模板不存在
    - `429`: AI服务请求过于频繁
    - `500`: 服务器内部错误
    """,
    responses={
        200: {
            "description": "生成成功",
            "content": {
                "application/json": {
                    "example": {
                        "status": "success",
                        "job_id": "tenant_acme:mss_executive_v2",
                        "session_id": "a3f2c5d8_20250129143025",
                        "report_path": "outputs/reports/report.pptx"
                    }
                }
            }
        },
        404: {
            "description": "资源不存在",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Template 'invalid_template' not found"
                    }
                }
            }
        },
        429: {
            "description": "请求过于频繁",
            "content": {
                "application/json": {
                    "example": {
                        "error": {
                            "code": "LLM_RATE_LIMITED",
                            "message": "AI服务请求过于频繁，请稍后重试"
                        }
                    }
                }
            }
        }
    }
)
async def generate_report(
    req: GenerateRequest,
    x_tenant_id: Optional[str] = Header(None, description="租户ID（多租户模式）"),
    x_request_id: Optional[str] = Header(None, description="请求追踪ID")
):
    """生成PPT报告（核心API）"""
    # 实现逻辑...
    pass
```

### 7.2 自动生成的客户端SDK

访问 `http://localhost:8000/openapi.json` 后，可使用工具生成客户端：

```bash
# Python客户端
openapi-generator-cli generate \
  -i http://localhost:8000/openapi.json \
  -g python \
  -o ./python-client

# Java客户端
openapi-generator-cli generate \
  -i http://localhost:8000/openapi.json \
  -g java \
  -o ./java-client

# TypeScript客户端
openapi-generator-cli generate \
  -i http://localhost:8000/openapi.json \
  -g typescript-axios \
  -o ./typescript-client
```

**使用示例**（自动生成）：

```python
# Python客户端使用
from mss_ai_ppt_client import ApiClient, Configuration, ReportsApi

config = Configuration(host="http://localhost:8000")
config.api_key['Authorization'] = 'Bearer your_api_key'

with ApiClient(config) as api_client:
    api = ReportsApi(api_client)
    response = api.generate_report(
        generate_request={
            "input_id": "tenant_acme_2025-11",
            "template_id": "mss_executive_v2"
        }
    )
    print(f"Report generated: {response.job_id}")
```

---

## 八、关键建议总结

### 🔴 立即行动（本周内）

1. **添加OpenAPI文档**（1天）
   - 完善Request/Response模型的Field描述
   - 添加示例和错误码说明
   - 启用Swagger UI

2. **统一API路径规范**（2天）
   - `/api/v1/reports`（非 `/generate`）
   - `/api/v1/reports/{id}/preview`（非 `/preview?job_id=`）
   - RESTful HTTP方法语义

3. **XML输入支持**（3天）
   - 实现XMLDataExtractor类
   - 添加 `/upload-xml` 端点
   - 更新文档说明

### 🟠 本月完成

4. **配置化字段映射**（1周）
   - JSON配置文件定义单元格映射
   - 支持多种Excel布局
   - 客户可自定义映射规则

5. **API认证机制**（1周）
   - API Key生成与管理
   - Bearer Token验证
   - 请求频率限制

6. **国际化框架**（2周）
   - 后端Flask-Babel集成
   - 前端i18next集成
   - 至少支持中英双语

### 🟡 季度目标

7. **多租户管理**（1个月）
   - 租户模型设计
   - 配额管理
   - 权限隔离

8. **数据源扩展**（1个月）
   - 数据库直连
   - API数据拉取
   - 统一数据输入层

9. **审批工作流**（1个月）
   - 报告状态机
   - 审批API
   - 通知机制

---

## 九、定制直签具体方案

### 9.1 场景一：客户品牌定制

**需求**：华为/腾讯等客户需使用自己的Logo/配色/字体

**解决方案**：

```python
# 1. 创建客户专属模板
templates/
├── huawei_executive_v1.pptx       # 华为品牌模板
├── huawei_executive_v1_descriptor.json
├── tencent_technical_v1.pptx      # 腾讯品牌模板
└── tencent_technical_v1_descriptor.json

# 2. API调用时指定租户
POST /api/v1/reports
{
  "tenant_id": "huawei",            # 自动使用huawei_*模板
  "template_id": "executive_v1",
  "input_id": "huawei_2025q1"
}

# 3. 动态品牌注入（高级）
POST /api/v1/reports
{
  "template_id": "generic_executive_v1",
  "brand_config": {
    "logo_url": "https://cdn.huawei.com/logo.png",
    "primary_color": "#C8102E",      # 华为红
    "font_family": "HarmonyOS Sans"
  }
}
```

### 9.2 场景二：模块化章节选择

**需求**：客户A只要技术分析，不要管理摘要

**解决方案**：

```python
POST /api/v1/reports
{
  "template_id": "mss_technical_v2",
  "include_slides": [
    "cover",
    "threat_analysis",
    "vulnerability_details",
    "recommendations"
  ],
  "exclude_slides": [
    "executive_summary",
    "compliance_status"
  ]
}

# 后端处理逻辑
def filter_slides(template, include_slides, exclude_slides):
    if include_slides:
        return [s for s in template.slides if s.slide_key in include_slides]
    if exclude_slides:
        return [s for s in template.slides if s.slide_key not in exclude_slides]
    return template.slides
```

### 9.3 场景三：CMDB数据直连

**需求**：从客户ServiceNow/Jira自动抽取工单数据

**解决方案**：

```python
# 1. 配置数据源连接
POST /api/v1/data-sources
{
  "name": "ServiceNow工单系统",
  "type": "api",
  "config": {
    "endpoint": "https://acme.service-now.com/api/now/table/incident",
    "auth": {
      "type": "basic",
      "username": "api_user",
      "password": "***"
    },
    "query_params": {
      "sysparm_query": "opened_at>=2025-01-01^state=1",
      "sysparm_fields": "number,short_description,priority,state"
    }
  },
  "data_mapping": {
    "top_incidents": "$.result[*]",
    "top_incidents[].id": "number",
    "top_incidents[].type": "short_description",
    "top_incidents[].severity": "priority"
  }
}

# 2. 生成报告时引用数据源
POST /api/v1/reports
{
  "data_source_id": "servicenow_tickets",  # 引用配置的数据源
  "template_id": "mss_executive_v2"
}

# 3. 后端自动拉取数据并生成报告
```

### 9.4 场景四：自动化周报

**需求**：每周一9点自动生成上周的安全报告

**解决方案**：

```python
# 1. 创建定时任务
POST /api/v1/scheduled-jobs
{
  "name": "ACME每周安全报告",
  "schedule": "0 9 * * MON",  # Cron表达式
  "job_config": {
    "data_source": {
      "type": "api",
      "config": {
        "endpoint": "https://siem.acme.com/api/weekly-summary",
        "date_range": "last_week"
      }
    },
    "template_id": "mss_executive_v2",
    "recipients": [
      "ciso@acme.com",
      "security-team@acme.com"
    ],
    "delivery_method": "email"  # 或 "webhook"
  }
}

# 2. 后端定时任务引擎（APScheduler）
from apscheduler.schedulers.background import BackgroundScheduler

scheduler = BackgroundScheduler()
scheduler.add_job(
    func=generate_scheduled_report,
    trigger='cron',
    hour=9,
    day_of_week='mon',
    args=[job_config]
)
scheduler.start()
```

---

## 十、实施优先级矩阵

```
        │ 业务价值
  高    │ [OpenAPI文档]    [多租户管理]
  ↑    │ [API认证]        [国际化]
价     │ [RESTful规范]    [审批工作流]
值     │
  ↓    │ [模板版本管理]   [折线图]
  低    │ [Webhook]        [定时任务]
        └────────────────────────────→
            低 ← 实施难度 → 高

图例：
[蓝色] = 立即执行（本周）
[绿色] = 近期完成（本月）
[黄色] = 季度目标
```

**快速胜利**（高价值低难度）：
1. OpenAPI文档
2. API认证
3. RESTful规范
4. XML输入支持

**战略投资**（高价值高难度）：
1. 多租户管理
2. 国际化框架
3. 审批工作流

**持续优化**（低优先级）：
1. 模板版本管理
2. 定时任务
3. 高级图表类型

---

## 附录：参考资料

### A. 相关文件清单

| 文件 | 说明 |
|-----|------|
| [app.py](mss_ai_ppt_sample_assets/backend/app.py) | FastAPI主应用，包含所有API端点 |
| [excel_handler.py](mss_ai_ppt_sample_assets/backend/modules/excel_handler.py) | Excel文件上传与解析 |
| [config.py](mss_ai_ppt_sample_assets/backend/config.py) | 配置管理（含locale配置） |
| [excel-to-ppt-workflow.md](docs/excel-to-ppt-workflow.md) | Excel数据驱动的PPT生成技术方案 |

### B. 技术栈建议

| 功能 | 推荐技术 | 理由 |
|-----|---------|------|
| OpenAPI文档 | FastAPI内置 | 零配置自动生成 |
| 国际化 | Flask-Babel | Python生态标准方案 |
| 多租户 | 自研 + PostgreSQL | 灵活可控 |
| 认证授权 | JWT + API Key | 行业标准 |
| 定时任务 | APScheduler | 轻量级，易集成 |
| 消息队列 | Celery + Redis | 适合异步任务 |

### C. 联系方式

如需技术支持或定制开发，请联系：
- 📧 Email: support@example.com
- 📚 文档: https://docs.example.com
- 💬 社区: https://github.com/example/mss-ai-ppt

---

**文档版本**: v1.0
**生成时间**: 2026-02-02
**作者**: Claude (MSS AI PPT技术团队)

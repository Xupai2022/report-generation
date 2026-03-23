# MSS AI PPT Sample Assets

这个目录当前提供的是一个可运行的 FastAPI 后端样例，用来完成以下流程：

- 上传 Excel 或使用内置输入数据
- 异步生成 PPT 报告
- 生成 PDF 和预览图
- 支持任务轮询、WebSocket 进度通知、报告改写、评分和后台管理
- 可选启用 LLM 与 RAG

现阶段主要围绕 `mss_classic_ops` 这一套经典模板运行。

## 当前真实目录

```text
mss_ai_ppt_sample_assets/
├─ backend/
│  ├─ app.py                     # FastAPI 入口
│  ├─ config.py                  # 环境变量与目录配置
│  ├─ requirements.txt           # Python 依赖
│  ├─ data/
│  │  ├─ templates/
│  │  │  ├─ catalog.json         # 模板目录
│  │  │  ├─ 经典模板.pptx
│  │  │  └─ 经典模板_descriptor.json
│  │  └─ inputs/
│  │     └─ catalog.json         # 输入数据目录
│  ├─ frontend/                  # 已编译前端静态资源
│  ├─ routers/v1/                # REST API
│  ├─ services/                  # 报告、任务、评分、后台服务
│  ├─ modules/                   # Excel、预览、认证、RAG 等模块
│  └─ outputs/                   # 运行时输出目录
│     ├─ jobs/
│     ├─ logs/
│     ├─ previews/
│     ├─ reports/
│     ├─ sessions/
│     └─ slidespecs/
└─ README.md
```

## 当前内置资源

### 模板

当前 `backend/data/templates/catalog.json` 只登记了 1 个模板：

- `mss_classic_ops`
  - 名称：`经典模板`
  - PPT 文件：`经典模板.pptx`
  - 描述文件：`经典模板_descriptor.json`
  - 页数：25
  - 特性：`excel_driven_content`、`classic_layout`

### 输入数据

当前 `backend/data/inputs/catalog.json` 只登记了 1 份内置输入：

- `classic_ops_dataxlsx`
  - 租户：`tenant_classic_demo`
  - 说明：来自 `backend/data/data.xlsx` 对应的解析结果

注意：

- 对于经典模板流程，后端会优先在运行时解析 Excel，而不是只依赖静态 JSON。
- 解析优先级如下：
  1. `outputs/sessions/{session_id}/uploaded.xlsx`
  2. catalog 中配置的 `excel_file`
  3. `backend/data/data.xlsx`（当 `input_id=classic_ops_dataxlsx` 时）

## 当前服务能力

### 页面与入口

- `/`：重定向到 `/ui/index.html`
- `/ui`：前端页面
- `/docs`：Swagger 文档
- `/redoc`：ReDoc 文档
- `/ws/{client_id}`：WebSocket 进度通道

### API 路由

统一前缀：`/api/v1`

- `reports`
  - `POST /api/v1/reports`：异步创建报告任务
  - `GET /api/v1/reports/{report_id}/download`：下载 PPTX
  - `GET /api/v1/reports/{report_id}/download-pdf`：下载 PDF
  - `GET /api/v1/reports/{report_id}/preview`：获取预览图
  - `PATCH /api/v1/reports/{report_id}/slides`：批量改写指定 slide 内容
  - `POST /api/v1/reports/{report_id}/slides/ai-rewrite`：AI 改写单页
- `jobs`
  - `GET /api/v1/jobs/{job_id}/status`
  - `GET /api/v1/jobs`
  - `POST /api/v1/jobs/{job_id}/cancel`
  - `DELETE /api/v1/jobs/{job_id}`
- `templates`
  - `GET /api/v1/templates`
  - `GET /api/v1/templates/{template_id}/slides`
- `inputs`
  - `GET /api/v1/inputs`
  - `GET /api/v1/inputs/{input_id}`
  - `POST /api/v1/inputs/excel`
- `sessions`
  - `DELETE /api/v1/sessions`
  - `DELETE /api/v1/sessions/{session_id}`
- `system`
  - `GET /api/v1/system/health`
  - `GET /api/v1/system/health/detailed`
  - `GET /api/v1/system/logs`
- `ratings`
  - `POST /api/v1/ratings/jobs/{job_id}/rating`
  - `GET /api/v1/ratings/jobs/{job_id}/rating`
  - `GET /api/v1/ratings/jobs/{job_id}/can-rate`
- `admin`
  - `POST /api/v1/admin/login`
  - `POST /api/v1/admin/logout`
  - `GET /api/v1/admin/verify`
  - `GET /api/v1/admin/jobs`
  - `GET /api/v1/admin/jobs/{job_id}`
  - `PATCH /api/v1/admin/jobs/{job_id}/rating`
  - `DELETE /api/v1/admin/jobs`
  - `GET /api/v1/admin/statistics`
- `rag`
  - `POST /api/v1/rag/index/build`
  - `POST /api/v1/rag/index/update`
  - `GET /api/v1/rag/index/status`
  - `POST /api/v1/rag/query`

## 运行依赖

### Python

建议 Python 3.10 及以上。

安装依赖：

```powershell
cd f:\report-generation\mss_ai_ppt_sample_assets\backend
pip install -r requirements.txt
```

### LibreOffice

预览图和 PDF 依赖 LibreOffice 的 `soffice`：

- 用于 `PPTX -> PDF -> PNG`
- 健康检查也会检测 LibreOffice 是否可用

Windows 下默认会尝试这些位置：

- `C:\Program Files\LibreOffice\program\soffice.exe`
- `C:\Program Files (x86)\LibreOffice\program\soffice.exe`
- `C:\Program Files\OpenOffice 4\program\soffice.exe`

也可以通过环境变量指定：

```powershell
$env:LIBREOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
```

### LLM / OpenAI Compatible API

当 `ENABLE_LLM=true` 时，后端要求：

- `OPENAI_API_KEY`
- 可选 `OPENAI_BASE_URL`
- `OPENAI_MODEL`

当 `ENABLE_LLM=false` 时，系统可以用 mock 流程运行，但涉及 AI 改写的接口不可用。

### RAG

RAG 相关依赖已经在 `requirements.txt` 中声明，包括：

- `qdrant-client`
- `sentence-transformers`
- `huggingface-hub`
- `tiktoken`

只有当 `RAG_ENABLED=true` 时，RAG 检索才会真正启用。

## 环境变量

以下是当前代码真正使用、且最值得关注的环境变量：

### 基础目录

- `MSS_ENV_PATH`
- `MSS_DATA_DIR`
- `MSS_OUTPUTS_DIR`
- `MSS_OUTPUTS_URL_PREFIX`

### LLM

- `ENABLE_LLM`
- `OPENAI_API_KEY`
- `OPENAI_BASE_URL`
- `OPENAI_MODEL`
- `LLM_CONNECT_TIMEOUT_SECONDS`
- `LLM_READ_TIMEOUT_SECONDS`
- `LLM_WRITE_TIMEOUT_SECONDS`
- `LLM_POOL_TIMEOUT_SECONDS`
- `LLM_RETRY_ATTEMPTS`
- `LLM_RETRY_BACKOFF_MIN_SECONDS`
- `LLM_RETRY_BACKOFF_MAX_SECONDS`
- `LLM_DISABLE_LOCAL_ONLY_BATCH_SPLIT`

### 通用

- `DEFAULT_LOCALE`
- `PREVIEW_CLEANUP_DAYS`
- `SESSION_RETENTION_DAYS`
- `JOB_RETENTION_DAYS`
- `JOB_MAX_RETRIES`
- `LOG_LEVEL`
- `LOG_MAX_BYTES`
- `LOG_BACKUP_COUNT`

### 管理后台

- `ADMIN_USERNAME`
- `ADMIN_PASSWORD_HASH`
- `ADMIN_SESSION_SECRET`
- `ADMIN_SESSION_MAX_AGE`

生成管理员密码哈希：

```powershell
cd f:\report-generation
python -m mss_ai_ppt_sample_assets.backend.scripts.generate_admin_password
```

### RAG

- `RAG_ENABLED`
- `RAG_VECTOR_BACKEND`
- `RAG_QDRANT_URL`
- `RAG_QDRANT_API_KEY`
- `RAG_QDRANT_COLLECTION`
- `RAG_QDRANT_PATH`
- `RAG_SOURCE_DIR`
- `RAG_EMBED_MODEL`
- `RAG_TOP_K`
- `RAG_CANDIDATE_TOP_K`
- `RAG_MAX_CONTEXT_CHARS`
- `RAG_MAX_CONTEXT_CHARS_PER_SLIDE`
- `RAG_MIN_SCORE`
- `RAG_PROMPT_BUDGET_RATIO`
- `RAG_PROMPT_BUDGET_MIN_TOKENS`
- `RAG_PROMPT_BUDGET_MAX_TOKENS`
- `RAG_CHUNK_SIZE_TOKENS`
- `RAG_CHUNK_OVERLAP_TOKENS`
- `RAG_HF_LOCAL_FILES_ONLY`
- `RAG_ENABLE_RERANK`
- `RAG_RERANK_MODEL`
- `RAG_RERANK_SCENES`
- `RAG_FINAL_TOP_K_PER_SLIDE`
- `RAG_COMMON_FALLBACK_TOP_K`
- `RAG_SECTION_MIN_CHARS`
- `RAG_PRELOAD_ON_STARTUP`

## 启动方式

在仓库根目录启动：

```powershell
cd f:\report-generation
python -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000 --reload
```

或者直接运行：

```powershell
cd f:\report-generation
python -m mss_ai_ppt_sample_assets.backend.app
```

启动后常用地址：

- 前端：`http://localhost:8000/ui/index.html`
- 登录页：`http://localhost:8000/ui/login.html`
- 管理页：`http://localhost:8000/ui/admin.html`
- Swagger：`http://localhost:8000/docs`

## 典型使用流程

### 1. 使用内置数据生成报告

```http
POST /api/v1/reports
Content-Type: application/json

{
  "input_id": "classic_ops_dataxlsx",
  "template_id": "mss_classic_ops",
  "use_mock": false,
  "use_rag": false,
  "focus_options": ["vulnerability"]
}
```

说明：

- `focus_options` 在当前接口里是必填字段，不能为空
- 允许值会被归一化为：
  - `vulnerability`
  - `alert`
  - `business_protection`

### 2. 上传 Excel 后生成报告

先上传：

```http
POST /api/v1/inputs/excel
Content-Type: multipart/form-data
```

然后使用返回的 `session_id` 发起生成：

```http
POST /api/v1/reports
Content-Type: application/json

{
  "input_id": "custom",
  "template_id": "mss_classic_ops",
  "session_id": "your_session_id",
  "use_mock": false,
  "use_rag": false,
  "focus_options": ["alert"]
}
```

这里的关键点是：

- 对 `mss_classic_ops`，后端会优先读取该 `session_id` 目录下的 `uploaded.xlsx`
- 因此上传 Excel 后，即使 `input_id` 不是 catalog 中的内置 ID，也可以跑通经典模板流程

### 3. 查询任务状态

```http
GET /api/v1/jobs/{job_id}/status
```

### 4. 下载与预览

- 下载 PPTX：`GET /api/v1/reports/{job_id}/download`
- 下载 PDF：`GET /api/v1/reports/{job_id}/download-pdf`
- 获取预览图：`GET /api/v1/reports/{job_id}/preview`

## 当前任务与并发行为

当前实现有几个重要行为需要明确：

- 报告生成是异步任务，`POST /api/v1/reports` 返回 `202`
- 浏览器通过 cookie 维持 `browser_id`
- 同一个浏览器默认只允许 1 个运行中的任务
- 若重复提交相同请求，系统会通过 `idempotency_key` 复用已有任务
- 若传入 `force_new_task=true`，会取消当前浏览器正在运行的旧任务，并创建新任务
- 服务启动时会把“上次服务异常重启时仍处于 running 的任务”标记为失败

## 报告改写能力

### 批量改写指定页

`PATCH /api/v1/reports/{report_id}/slides`

- 直接更新 slide placeholder 内容
- 改写后会重新渲染 PPT，并重新生成预览图

### AI 改写单页

`POST /api/v1/reports/{report_id}/slides/ai-rewrite`

- 仅在 `ENABLE_LLM=true` 时可用
- `user_prompt` 必填
- `target_tokens` 可选，不传则改写当前页所有 `ai_generate=true` 的 token
- 当前实现里 `use_rag` 参数保留用于兼容，但后端会显式跳过 RAG 检索，优先严格遵循用户提示词

## 评分与后台管理

### 用户评分

- `liked` / `disliked`
- 点踩必须附带至少 10 个字符的评论
- 每个 session 对同一任务只能评一次

### 管理后台

当前后台认证依赖 FastAPI `SessionMiddleware` 的 cookie session。

常用接口：

- 登录：`POST /api/v1/admin/login`
- 校验登录态：`GET /api/v1/admin/verify`
- 任务分页、筛选、搜索：`GET /api/v1/admin/jobs`
- 统计信息：`GET /api/v1/admin/statistics`

## 自动清理与启动时行为

服务启动时会尝试执行以下清理动作：

- 可选预热 RAG
- 清理过期 session
- 清理预览临时目录和孤儿预览目录
- 清理陈旧文件锁
- 将重启前未完成任务标记为失败
- 清理过期 job
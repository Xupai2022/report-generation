# 后端说明

本文只记录当前 `mss_ai_ppt_sample_assets/backend/` 中能够直接被代码证明的事实。

## 1. 后端入口

后端主入口文件：

- `mss_ai_ppt_sample_assets/backend/app.py`

当前可直接确认：

- 使用 FastAPI
- `/docs` 为 OpenAPI 页面
- `/redoc` 为 ReDoc 页面
- `/` 会重定向到 `/ui/index.html`
- `/api` 会返回 API 概览
- WebSocket 入口为 `/ws/{client_id}`

来源：`mss_ai_ppt_sample_assets/backend/app.py:37`

## 2. `.env` 与配置加载

配置定义位于：

- `mss_ai_ppt_sample_assets/backend/config.py`

当前可直接确认：

- 仓库根目录通过 `REPO_ROOT_DIR` 计算得到
- 默认从仓库根目录 `.env` 读取配置
- 也可以通过 `MSS_ENV_PATH` 指定其他 `.env` 文件
- `load_dotenv(..., override=True)` 会覆盖同名环境变量

来源：`mss_ai_ppt_sample_assets/backend/config.py:6`

## 3. 目录约定

当前配置文件中定义了以下关键目录：

- `DATA_DIR`
- `TEMPLATES_DIR`
- `INPUTS_DIR`
- `OUTPUTS_DIR`
- `REPORTS_DIR`
- `LOGS_DIR`
- `PREVIEWS_DIR`
- `SLIDESPECS_DIR`
- `SESSIONS_DIR`
- `JOBS_DIR`
- `RAG_DIR`

其中默认关系可直接从代码确认：

- `DATA_DIR` 默认是 `backend/data/`
- `OUTPUTS_DIR` 默认是 `backend/outputs/`
- 其他目录都是 `OUTPUTS_DIR` 或 `DATA_DIR` 下的子目录

来源：`mss_ai_ppt_sample_assets/backend/config.py:13`

## 4. 输出物 URL 与静态挂载

当前后端会把运行期文件按以下方式挂载：

- `/static/previews` -> `PREVIEWS_DIR`
- `OUTPUTS_URL_PREFIX` -> `OUTPUTS_DIR`
- `/ui` -> `backend/frontend/`
- `/assets` -> `backend/frontend/assets/`
- `/i18n` -> `backend/frontend/i18n/`

其中：

- `OUTPUTS_URL_PREFIX` 默认值对应 `/outputs`
- `/outputs` 静态挂载在代码注释中被标注为更适合开发 / 管理用途
- 生产环境更适合优先走受控下载接口

来源：

- `mss_ai_ppt_sample_assets/backend/config.py:28`
- `mss_ai_ppt_sample_assets/backend/app.py:82`
- `mss_ai_ppt_sample_assets/backend/app.py:84`
- `mss_ai_ppt_sample_assets/backend/app.py:100`

## 5. 环境变量分类

### LLM / OpenAI 相关

当前配置文件中直接读取：

- `OPENAI_API_KEY`
- `OPENAI_BASE_URL`
- `OPENAI_MODEL`
- `ENABLE_LLM`
- `LLM_CONNECT_TIMEOUT_SECONDS`
- `LLM_READ_TIMEOUT_SECONDS`
- `LLM_WRITE_TIMEOUT_SECONDS`
- `LLM_POOL_TIMEOUT_SECONDS`
- `LLM_RETRY_ATTEMPTS`
- `LLM_RETRY_BACKOFF_MIN_SECONDS`
- `LLM_RETRY_BACKOFF_MAX_SECONDS`
- `LLM_DISABLE_LOCAL_ONLY_BATCH_SPLIT`
- `DEFAULT_LOCALE`

并且只有在 `ENABLE_LLM=true` 时，`OPENAI_API_KEY` 才是强制项。

来源：`mss_ai_ppt_sample_assets/backend/config.py:45`

### 清理、保留与日志相关

- `PREVIEW_CLEANUP_DAYS`
- `SESSION_RETENTION_DAYS`
- `LOG_LEVEL`
- `LOG_MAX_BYTES`
- `LOG_BACKUP_COUNT`
- `JOB_RETENTION_DAYS`
- `JOB_MAX_RETRIES`

来源：`mss_ai_ppt_sample_assets/backend/config.py:79`

### 管理后台认证相关

- `ADMIN_USERNAME`
- `ADMIN_PASSWORD_HASH`
- `ADMIN_SESSION_SECRET`
- `ADMIN_SESSION_MAX_AGE`

来源：`mss_ai_ppt_sample_assets/backend/config.py:94`

### RAG 相关

当前仓库已经存在完整的 RAG 配置集，包括：

- `RAG_ENABLED`
- `RAG_VECTOR_BACKEND`
- `RAG_QDRANT_URL`
- `RAG_QDRANT_API_KEY`
- `RAG_QDRANT_COLLECTION`
- `RAG_QDRANT_PATH`
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
- `RAG_SOURCE_DIR`
- `RAG_ENABLE_RERANK`
- `RAG_RERANK_MODEL`
- `RAG_RERANK_SCENES`
- `RAG_FINAL_TOP_K_PER_SLIDE`
- `RAG_COMMON_FALLBACK_TOP_K`
- `RAG_SECTION_MIN_CHARS`
- `RAG_PRELOAD_ON_STARTUP`

来源：`mss_ai_ppt_sample_assets/backend/config.py:103`

## 6. 中间件与运行时默认值

当前后端注册了：

- `ErrorLoggingMiddleware`
- `RequestIdMiddleware`
- `SessionMiddleware`

其中 `SessionMiddleware` 的关键参数包括：

- `session_cookie="admin_session"`
- `same_site="lax"`
- `https_only=False`

这说明当前默认值偏向开发环境。生产环境如果启用了 HTTPS，应重新审查该配置。

来源：`mss_ai_ppt_sample_assets/backend/app.py:45`

## 7. `/api/v1` 路由概览

当前统一前缀为：

- `/api/v1`

子路由包括：

- `/reports`
- `/templates`
- `/inputs`
- `/sessions`
- `/system`
- `/jobs`
- `/admin`
- `/ratings`
- `/rag`

这意味着当前文档不应再写旧式根路径接口。

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:14`

## 8. 与生成直接相关的核心接口

### reports

- `POST /api/v1/reports`
- `GET /api/v1/reports/{report_id}/download`
- `GET /api/v1/reports/{report_id}/download-pdf`
- `GET /api/v1/reports/{report_id}/preview`
- `PATCH /api/v1/reports/{report_id}/slides`
- `POST /api/v1/reports/{report_id}/slides/ai-rewrite`

其中必须明确：

- `POST /api/v1/reports` 是异步创建 job
- 下载与 PDF 下载是两个独立接口
- 预览返回的是预览图 URL，不是 PPT 二进制内容
- 手动改写与 AI 改写都基于既有 job 继续操作

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:430`

### templates / inputs

- `GET /api/v1/templates`
- `GET /api/v1/templates/{template_id}/slides`
- `GET /api/v1/inputs`
- `GET /api/v1/inputs/{input_id}`
- `POST /api/v1/inputs/excel`

来源：

- `mss_ai_ppt_sample_assets/backend/routers/v1/templates.py:16`
- `mss_ai_ppt_sample_assets/backend/routers/v1/inputs.py:25`

### jobs / sessions / system

- `GET /api/v1/jobs/{job_id}/status`
- `GET /api/v1/jobs`
- `POST /api/v1/jobs/{job_id}/cancel`
- `DELETE /api/v1/jobs/{job_id}`
- `DELETE /api/v1/sessions`
- `DELETE /api/v1/sessions/{session_id}`
- `GET /api/v1/system/health`
- `GET /api/v1/system/health/detailed`
- `GET /api/v1/system/logs`

其中一个重要真实性边界是：

- `POST /api/v1/jobs/{job_id}/cancel` 只会更新状态，不会真正中断正在执行的生成过程

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/jobs.py:150`

## 9. 作业状态与落盘事实

当前作业状态保存在文件系统中，而不是数据库中。

作业存储结构：

```text
outputs/jobs/
├── states/
├── index.json
└── idempotency/
```

并且 `job_id` 当前格式为：

- `{session_id}:{template_id}`

来源：

- `mss_ai_ppt_sample_assets/backend/modules/job_store.py:0`
- `mss_ai_ppt_sample_assets/backend/services/job_manager.py:98`

## 10. 启动时清理与恢复

后端启动时会执行：

- RAG 预热（可选）
- 清理旧 session
- 清理预览临时文件
- 清理孤立或过期 preview
- 清理 stale lock
- 将因重启中断的 running job 标记为 failed
- 清理旧 job

来源：`mss_ai_ppt_sample_assets/backend/app.py:246`

## 11. `/api/v1/system/logs` 的安全边界

系统日志接口：

- `GET /api/v1/system/logs`

当前路由文档中已明确说明该接口可能暴露敏感信息。生产环境至少应重新审查以下事项：

- 是否需要认证 / 授权
- 是否需要脱敏
- 是否需要限流

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py:152`

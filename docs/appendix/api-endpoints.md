# API 端点索引

本文按当前实际 router 分组整理 `/api/v1` 端点索引，只记录当前代码中可直接核实的入口。

## 1. 路由总表

当前 `/api/v1` 下挂载的子路由有：

- `/reports`
- `/templates`
- `/inputs`
- `/sessions`
- `/system`
- `/jobs`
- `/admin`
- `/ratings`
- `/rag`

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:14`

## 2. Reports

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py`

- `POST /api/v1/reports`
- `GET /api/v1/reports/{report_id}/download`
- `GET /api/v1/reports/{report_id}/download-pdf`
- `GET /api/v1/reports/{report_id}/preview`
- `PATCH /api/v1/reports/{report_id}/slides`
- `POST /api/v1/reports/{report_id}/slides/ai-rewrite`

说明：

- `POST /api/v1/reports` 是异步创建 job 的入口
- 下载、PDF、预览、改写都基于已有 `job_id`

## 3. Templates

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/templates.py`

- `GET /api/v1/templates`
- `GET /api/v1/templates/{template_id}/slides`

## 4. Inputs

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/inputs.py`

- `GET /api/v1/inputs`
- `GET /api/v1/inputs/{input_id}`
- `POST /api/v1/inputs/excel`

## 5. Sessions

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/sessions.py`

- `DELETE /api/v1/sessions`
- `DELETE /api/v1/sessions/{session_id}`

## 6. System

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py`

- `GET /api/v1/system/health`
- `GET /api/v1/system/health/detailed`
- `GET /api/v1/system/logs`

## 7. Jobs

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/jobs.py`

- `GET /api/v1/jobs/{job_id}/status`
- `GET /api/v1/jobs`
- `POST /api/v1/jobs/{job_id}/cancel`
- `DELETE /api/v1/jobs/{job_id}`

说明：

- `cancel` 当前只更新状态，不会真正中断正在执行的任务

## 8. Admin

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/admin.py`

- `POST /api/v1/admin/login`
- `POST /api/v1/admin/logout`
- `GET /api/v1/admin/verify`
- `GET /api/v1/admin/jobs`
- `GET /api/v1/admin/jobs/{job_id}`
- `PATCH /api/v1/admin/jobs/{job_id}/rating`
- `DELETE /api/v1/admin/jobs`
- `GET /api/v1/admin/statistics`

说明：

- `GET /api/v1/admin/jobs/{job_id}` 在前端会把第一个 `:` 编码成 `_` 再发送
- 管理接口依赖 session 认证

## 9. Ratings

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/ratings.py`

- `POST /api/v1/ratings/jobs/{job_id}/rating`
- `GET /api/v1/ratings/jobs/{job_id}/rating`
- `GET /api/v1/ratings/jobs/{job_id}/can-rate`

## 10. RAG

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/rag.py`

- `POST /api/v1/rag/index/build`
- `POST /api/v1/rag/index/update`
- `GET /api/v1/rag/index/status`
- `POST /api/v1/rag/query`

## 11. 非 `/api/v1` 但同样重要的入口

这些入口不属于 `v1_router`，但在排障和使用时经常一起核对：

- `GET /`
- `GET /api`
- `GET /docs`
- `GET /redoc`
- `WebSocket /ws/{client_id}`

来源：`mss_ai_ppt_sample_assets/backend/app.py:131`

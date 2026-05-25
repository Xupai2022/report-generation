# 排障与诊断

本文围绕当前代码中可以直接确认的诊断入口整理，不沿用过时接口名和旧同步流程描述。

## 1. 先确认服务是否在线

优先检查以下入口：

- `/`
- `/api`
- `/docs`
- `/api/v1/system/health`
- `/api/v1/system/health/detailed`

其中：

- `/` 会跳转到 `/ui/index.html`
- `/api` 会返回 API 概览
- `/api/v1/system/health` 提供基础健康状态
- `/api/v1/system/health/detailed` 提供更详细的组件检查

来源：

- `mss_ai_ppt_sample_assets/backend/app.py:131`
- `mss_ai_ppt_sample_assets/backend/app.py:138`
- `mss_ai_ppt_sample_assets/backend/routers/v1/system.py:16`
- `mss_ai_ppt_sample_assets/backend/routers/v1/system.py:83`

## 2. 健康检查怎么用

### 基础健康检查

接口：

- `GET /api/v1/system/health`

当前可直接确认它会返回：

- `status`
- LLM 并发状态
  - `max_concurrent_requests`
  - `available_slots`
  - `active_requests`
  - `is_saturated`

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py:16`

### 详细健康检查

接口：

- `GET /api/v1/system/health/detailed`

当前可直接确认它覆盖至少这些组件：

- 磁盘
- OpenAI API
- LibreOffice
- active sessions
- file locks
- LLM concurrency

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py:83`

## 3. 日志查看入口

接口：

- `GET /api/v1/system/logs`

适用场景：

- 报告生成失败
- 预览失败
- PDF 下载失败
- Excel 上传失败
- AI 重写失败
- 启动清理或 RAG 预热异常

当前代码中也明确提示：该接口在生产环境可能暴露敏感信息。

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py:152`

## 4. LLM 超时或连接失败

当前前后端都能直接确认以下情况：

- 后端会把一部分上游错误分类为机器可读错误码
- 前端把 `LLM_TIMEOUT_EXHAUSTED` 与 `LLM_UPSTREAM_CONNECTION_FAILED` 视为需要重新发起新任务的典型情况

后端在异步任务处理中也会对可重试错误执行有限重试，并写入进度消息。

排查顺序建议：

1. `GET /api/v1/jobs/{job_id}/status`
2. `GET /api/v1/system/health`
3. `GET /api/v1/system/health/detailed`
4. `GET /api/v1/system/logs`
5. 检查 `.env` 中 LLM 相关配置是否正确

来源：

- 前端超时错误码：`ui-src/src/pages/IndexApp.tsx:55`
- 后端错误分类与重试：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:144`
- LLM 环境变量：`mss_ai_ppt_sample_assets/backend/config.py:45`

## 5. LibreOffice / 预览 / PDF 问题

### 当前预览与 PDF 依赖什么

当前代码可以直接确认：

- 预览链路：`PPTX -> LibreOffice -> PDF -> PNG`
- PDF 下载同样依赖 LibreOffice 转换
- 程序会优先读取 `LIBREOFFICE_PATH`
- 然后尝试常见 Windows 路径
- 最后退回 `which soffice`

来源：

- `mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:42`
- `mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:69`
- `mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:691`

### 排查顺序

1. `GET /api/v1/jobs/{job_id}/status` 是否已完成
2. `GET /api/v1/system/health/detailed` 中 LibreOffice 是否正常
3. `GET /api/v1/system/logs` 是否已有转换报错
4. 服务器或本机是否真的安装了 LibreOffice
5. `LIBREOFFICE_PATH` 是否有效

### 如果 PPT 可以下载但 PDF 不行

优先怀疑：

- LibreOffice 不可用
- PDF 转换环节失败
- 相关缓存文件缺失或转换失败

## 6. 任务一直 running / pending

优先看：

- `GET /api/v1/jobs/{job_id}/status`
- `GET /api/v1/system/health`
- `GET /api/v1/system/logs`

当前任务状态明确包括：

- `pending`
- `running`
- `completed`
- `failed`
- `cancelled`

来源：`mss_ai_ppt_sample_assets/backend/models/job_state.py:14`

如果任务卡住，常见排查方向包括：

- LLM 上游超时 / 拒连
- LibreOffice / 预览阶段失败
- 运行环境缺少依赖
- 当前并发槽位饱和

## 7. 运行中任务复用与 supersede

这部分是当前实现中最容易被误判为 bug 的地方。

当前可以直接确认：

- 同一浏览器默认只允许 1 个运行中的任务
- 如果浏览器已有运行中的任务，新的生成请求可能复用旧任务
- 当 `force_new_task=true` 时，会创建新任务，并把旧任务标记为被新任务取代（superseded）

因此，如果用户感觉“重复点生成却没有新任务”，应先核对是否命中了任务复用逻辑。

来源：

- 浏览器级任务限制：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:32`
- 复用 running job：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:499`
- supersede 逻辑：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:522`

## 8. RAG 关闭、预热与状态问题

当前代码可以直接确认：

- `RAG_ENABLED` 控制 RAG 总开关
- `RAG_PRELOAD_ON_STARTUP` 控制是否在启动时预热
- 当未启用预热时，启动日志会明确写“skipped”
- 当前还提供 RAG 状态与索引接口

可用排查入口：

- `GET /api/v1/rag/index/status`
- `POST /api/v1/rag/index/build`
- `POST /api/v1/rag/index/update`
- `POST /api/v1/rag/query`
- 启动日志

来源：

- 启动预热逻辑：`mss_ai_ppt_sample_assets/backend/app.py:246`
- RAG 路由：`mss_ai_ppt_sample_assets/backend/routers/v1/rag.py:22`
- RAG 配置：`mss_ai_ppt_sample_assets/backend/config.py:103`

## 9. 启动清理行为

当前后端在启动时会自动执行一系列清理与恢复动作，因此有些“文件被删了 / running 任务变 failed”并不是异常，而是当前实现的一部分。

启动时会执行：

- 清理旧 session
- 清理 preview tmp
- 清理孤立或过期 preview
- 清理 stale lock
- 将因服务重启中断的 running job 标记为 failed
- 清理过期 job

来源：`mss_ai_ppt_sample_assets/backend/app.py:246`

如果你在重启后看到 job 状态变成失败，先核对是否属于 `RESTART_INTERRUPTED` 场景。

来源：

- 重启中断错误定义：`mss_ai_ppt_sample_assets/backend/services/job_manager.py:18`
- 启动恢复：`mss_ai_ppt_sample_assets/backend/app.py:303`

## 10. Excel 上传问题

当前实现可直接确认：

- 仅支持 `.xlsx`
- `ExcelHandler` 文件大小限制为 50MB
- 上传成功会返回 `session_id`
- 前端随后把输入源切换为 `custom`

来源：

- 后端上传路由：`mss_ai_ppt_sample_assets/backend/routers/v1/inputs.py:197`
- handler 初始化：`mss_ai_ppt_sample_assets/backend/routers/v1/inputs.py:22`
- 前端上传 API：`ui-src/src/api/main.ts:23`

若上传失败，优先检查：

1. 扩展名是否正确
2. 文件大小是否超过限制
3. 后端日志中是否已有校验异常
4. 返回体里是否已经给出明确错误信息

## 11. 手动编辑或 AI 重写失败

### 手动编辑失败

接口：

- `PATCH /api/v1/reports/{job_id}/slides`

优先检查：

- `job_id` 是否存在
- `slide_key` 是否存在于该报告的 slidespec 中
- `new_content` 是否符合原有结构要求
- 日志中是否出现 invalid request 或 slidespec not found

### AI 重写失败

接口：

- `POST /api/v1/reports/{job_id}/slides/ai-rewrite`

优先检查：

- `slide_key` 是否正确
- `user_prompt` 是否为空
- 当前页是否存在可 AI 重写的 token
- LLM 或限流错误是否已在日志中出现

来源：

- 手动更新路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:790`
- AI 重写路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:887`
- 请求模型：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:113`

## 12. 诊断时优先使用哪些事实来源

如果页面文案、旧 README 或口头约定与实际现象不一致，优先回指：

- `mss_ai_ppt_sample_assets/backend/app.py`
- `mss_ai_ppt_sample_assets/backend/config.py`
- `mss_ai_ppt_sample_assets/backend/routers/v1/reports.py`
- `mss_ai_ppt_sample_assets/backend/routers/v1/system.py`
- `mss_ai_ppt_sample_assets/backend/modules/preview_generator.py`
- `mss_ai_ppt_sample_assets/backend/services/job_manager.py`
- `ui-src/src/pages/IndexApp.tsx`

当前文档统一以这些真实实现为准。

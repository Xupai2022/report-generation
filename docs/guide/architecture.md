# 架构说明

本文只保留当前仓库中能够被代码直接证明的架构边界与主流程。

## 1. 总体边界

从当前实现可以直接确认，系统由以下几层组成：

1. 前端交互层：`ui-src/`
2. FastAPI / WebSocket 接入层：`mss_ai_ppt_sample_assets/backend/app.py`
3. `/api/v1` 路由层：`mss_ai_ppt_sample_assets/backend/routers/v1/*.py`
4. 业务编排层：`ReportService` + `JobManager`
5. 能力模块层：模板加载、LLM 编排、PPT 渲染、预览生成、RAG、文件锁、作业存储
6. 文件系统持久化层：`backend/data/` 与 `backend/outputs/`

对应事实来源：

- 应用入口与路由挂载：`mss_ai_ppt_sample_assets/backend/app.py:37`
- V1 路由总表：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:14`
- 报告服务初始化：`mss_ai_ppt_sample_assets/backend/services/report_service.py:42`
- 作业管理：`mss_ai_ppt_sample_assets/backend/services/job_manager.py:26`

## 2. API 与入口层

当前主 API 前缀统一为：

- `/api/v1`

同时后端还直接提供：

- `/` -> `/ui/index.html`
- `/api`
- `/docs`
- `/redoc`
- `/ws/{client_id}`

这说明当前系统的外部入口不只是 REST API，还包括前端静态站点与实时进度 WebSocket。

来源：

- `mss_ai_ppt_sample_assets/backend/app.py:37`
- `mss_ai_ppt_sample_assets/backend/app.py:88`
- `mss_ai_ppt_sample_assets/backend/app.py:116`
- `mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:14`

## 3. 路由层职责

`/api/v1` 下当前注册的子路由包括：

- `/reports`
- `/templates`
- `/inputs`
- `/sessions`
- `/system`
- `/jobs`
- `/admin`
- `/ratings`
- `/rag`

它们负责定义协议边界，但不承担完整生成流程本身。核心生成、改写、预览、落盘仍由 service / module 层负责。

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:17`

## 4. 核心主流程：从请求到报告文件

### 第一步：前端发起异步创建

前端通过 `POST /api/v1/reports` 发起生成请求。请求中会携带：

- `input_id`
- `template_id`
- `focus_options`
- `use_mock`
- `use_rag`
- `session_id`
- `client_id`
- `idempotency_key`
- `force_new_task`

来源：

- 前端 API：`ui-src/src/api/main.ts:58`
- 请求模型：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:6`
- 创建路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:430`

### 第二步：后端创建或复用 job

`reports.py` 中会：

- 生成或读取 browser cookie
- 按浏览器检查是否已有运行中任务
- 按 `idempotency_key` 决定复用已有 job 还是创建新 job
- 在 `force_new_task=true` 时把旧任务标记为 superseded / cancelled

这说明当前任务模型不是“每点一次都无条件新建任务”，而是带有浏览器级复用和幂等语义。

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:477`

### 第三步：后台异步执行报告生成

路由返回 `202` 后，会通过 `asyncio.create_task(...)` 在后台继续执行 `_process_report_async(...)`。

异步处理流程中会：

- 读取 job
- 根据是否 `use_mock` 决定是否进入 LLM 并发信号量
- 调用 `ReportService.generate(...)`
- 处理重试、失败分类与用户可见错误码
- 成功后继续生成预览图
- 最终把结果写回 job 状态

来源：

- 异步入口：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:245`
- 创建后台任务：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:614`

## 5. LLM 编排、SlideSpec 与 PPT 渲染

### ReportService 是主编排层

`ReportService` 初始化时统一持有：

- `TemplateRepository`
- `AuditLogger`
- `PPTPreviewGenerator`
- `SessionManager`
- `RAG service`
- `PPTGeneratorV2`
- `LLMOrchestratorV2`

这说明当前生成主链路通过 `ReportService` 把模板、输入、LLM、预览、session 管理串起来。

来源：`mss_ai_ppt_sample_assets/backend/services/report_service.py:42`

### LLMOrchestratorV2 负责生成内容，不直接产出最终 PPT

从服务层依赖可以确认，LLM 编排与 PPT 渲染是分开的：

- `LLMOrchestratorV2`：负责把输入、focus、RAG 上下文等整理成可渲染内容
- `PPTGeneratorV2`：负责把结果写入模板 PPTX

在文档层面，当前能稳定描述的中间契约是：系统会把生成结果与 slidespec 文件一起落盘，并把其路径记录到 job 状态中。

来源：

- 服务依赖：`mss_ai_ppt_sample_assets/backend/services/report_service.py:55`
- job 完成时写入 `slidespec_path`：`mss_ai_ppt_sample_assets/backend/services/job_manager.py:151`

## 6. 预览链路

预览能力不是前端本地生成，而是后端把已生成的 PPTX 转换为图片。

当前可以直接确认的链路为：

- `PPTX -> LibreOffice -> PDF -> PyMuPDF -> PNG`

并且当前实现带有：

- `staging` 目录
- 预览缓存复用
- 预览耗时记录
- 启动时的预览目录清理

来源：

- 预览生成器：`mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:42`
- staging 目录：`mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:245`
- 启动清理：`mss_ai_ppt_sample_assets/backend/app.py:282`

## 7. 异步 job + session/output 落盘模型

这是当前架构中最重要的事实边界之一。

### 作业状态持久化

`JobStore` 把作业状态写入文件系统，结构为：

```text
outputs/jobs/
├── states/
├── index.json
└── idempotency/
```

这说明当前系统不是数据库驱动的 job 队列，而是**文件系统持久化的作业状态模型**。

来源：`mss_ai_ppt_sample_assets/backend/modules/job_store.py:0`

### job_id 与 session_id 的关系

`JobManager.create_job(...)` 中可直接确认：

- `job_id` 格式为 `{session_id}:{template_id}`
- `session_id` 可由请求传入，也可由服务端生成

来源：`mss_ai_ppt_sample_assets/backend/services/job_manager.py:44`

### 运行产物会落到 outputs 目录

从配置中可直接确认，运行时会落盘到以下目录：

- `outputs/reports`
- `outputs/logs`
- `outputs/previews`
- `outputs/slidespecs`
- `outputs/sessions`
- `outputs/jobs`
- `outputs/rag`

来源：`mss_ai_ppt_sample_assets/backend/config.py:18`

## 8. 实时进度：WebSocket + 轮询双通道

当前系统不是只靠一种方式通知进度。

可以直接确认：

- 后端提供 `/ws/{client_id}` WebSocket
- 前端也会轮询 `/api/v1/jobs/{job_id}/status`
- WebSocket 进度会同步写回 job，用于轮询客户端获取一致状态

来源：

- WebSocket 入口：`mss_ai_ppt_sample_assets/backend/app.py:116`
- 同步 job 进度：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:82`
- 前端状态接口：`ui-src/src/api/main.ts:59`

## 9. 日志、健康检查与系统诊断

系统层当前提供：

- `GET /api/v1/system/health`
- `GET /api/v1/system/health/detailed`
- `GET /api/v1/system/logs`

这意味着当前后端不仅负责生成，也内置了：

- 并发状态暴露
- 组件健康检查
- 文件日志读取

来源：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py:16`

## 10. 启动时清理与恢复

后端在启动时会执行一系列清理与恢复动作，包括：

- 可选的 RAG 预热
- 清理过期 session
- 清理 preview tmp
- 清理过期或孤立的 preview 产物
- 清理陈旧文件锁
- 将因重启中断的 running job 标记为 failed
- 清理过期 job

这说明当前系统不是“纯请求驱动”模型，而是带有明显的启动恢复逻辑。

来源：`mss_ai_ppt_sample_assets/backend/app.py:246`

## 11. 当前架构不应再被描述成什么

为避免沿用旧说法，本文不再把当前系统描述成：

- 单次同步 `/generate` 直接返回 PPT
- 前端直接调用 AI 生成内容
- 只在内存中维护任务状态
- 不落盘 session / output 文件
- 只通过前端下载按钮完成所有导出逻辑

当前事实是：这是一个以 FastAPI 为入口、以异步 job 为主线、以文件系统为主要持久化介质、通过后端统一执行生成/改写/预览/导出的报告系统。

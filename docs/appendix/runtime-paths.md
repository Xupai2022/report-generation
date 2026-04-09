# 运行时路径

本文整理当前仓库中能够直接确认的运行时文件系统路径与对外 URL 路径。

## 1. 路径根

### 代码根

- `ROOT_DIR` -> `mss_ai_ppt_sample_assets/backend/`
- `FRONTEND_DIR` -> `mss_ai_ppt_sample_assets/backend/frontend/`

来源：`mss_ai_ppt_sample_assets/backend/config.py:10`

### 数据根

- `DATA_DIR` -> 默认 `backend/data/`
- `TEMPLATES_DIR` -> `DATA_DIR/templates/`
- `INPUTS_DIR` -> `DATA_DIR/inputs/`

来源：`mss_ai_ppt_sample_assets/backend/config.py:13`

### 运行输出根

- `OUTPUTS_DIR` -> 默认 `backend/outputs/`

来源：`mss_ai_ppt_sample_assets/backend/config.py:18`

## 2. outputs 下的主要目录与用途

### `REPORTS_DIR`

- 文件系统路径：`outputs/reports/`
- 用途：通用报告输出目录常量
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:19`

### `LOGS_DIR`

- 文件系统路径：`outputs/logs/`
- 用途：日志文件目录
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:20`

### `PREVIEWS_DIR`

- 文件系统路径：`outputs/previews/`
- 用途：预览图片与中间预览产物
- 对外静态路径：`/static/previews`
- 来源：
  - `mss_ai_ppt_sample_assets/backend/config.py:21`
  - `mss_ai_ppt_sample_assets/backend/app.py:82`

### `SLIDESPECS_DIR`

- 文件系统路径：`outputs/slidespecs/`
- 用途：slidespec 文件目录
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:22`

### `SESSIONS_DIR`

- 文件系统路径：`outputs/sessions/`
- 用途：隔离 session 目录，用于并发请求、上传输入、中间产物
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:23`

### `JOBS_DIR`

- 文件系统路径：`outputs/jobs/`
- 用途：作业状态持久化目录
- 典型结构：`states/`、`index.json`、`idempotency/`
- 来源：
  - `mss_ai_ppt_sample_assets/backend/config.py:24`
  - `mss_ai_ppt_sample_assets/backend/modules/job_store.py:0`

### `RAG_DIR`

- 文件系统路径：`outputs/rag/`
- 用途：RAG 运行时目录
- 元数据文件：`index_meta.json`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:25`

## 3. 对外 URL 路径

### `OUTPUTS_URL_PREFIX`

- 默认值：`/outputs`
- 作用：把整个 `OUTPUTS_DIR` 作为静态目录暴露
- 说明：代码注释中明确标为更适合开发 / 管理用途
- 来源：
  - `mss_ai_ppt_sample_assets/backend/config.py:28`
  - `mss_ai_ppt_sample_assets/backend/app.py:84`

### `/static/previews`

- 对应目录：`PREVIEWS_DIR`
- 用途：预览图静态访问
- 来源：`mss_ai_ppt_sample_assets/backend/app.py:82`

### `/ui`

- 对应目录：`backend/frontend/`
- 用途：托管构建后的前端页面
- 来源：`mss_ai_ppt_sample_assets/backend/app.py:101`

### `/assets`

- 对应目录：`backend/frontend/assets/`
- 用途：前端构建后的 JS/CSS/静态资源
- 来源：`mss_ai_ppt_sample_assets/backend/app.py:104`

### `/i18n`

- 对应目录：`backend/frontend/i18n/`
- 用途：国际化资源文件
- 来源：`mss_ai_ppt_sample_assets/backend/app.py:108`

## 4. 任务、session 与预览的关系

当前架构中可以直接确认：

- `job_id` 格式为 `{session_id}:{template_id}`
- session 目录负责隔离一次生成相关的输入与中间文件
- preview 目录则按 job / session 产出对外可访问的预览图片
- job 状态文件会记录 `report_path`、`slidespec_path`、`preview_urls`

来源：

- `mss_ai_ppt_sample_assets/backend/services/job_manager.py:98`
- `mss_ai_ppt_sample_assets/backend/models/job_state.py:82`

## 5. 与部署相关的路径真实性边界

以下路径属于代码中的默认值或仓库样例，部署时仍需按环境校验：

- `.env` 的真实位置
- `MSS_DATA_DIR` / `MSS_OUTPUTS_DIR` 是否被覆盖
- LibreOffice 安装路径
- 服务器对 outputs 目录的权限与容量
- systemd 样例中的绝对路径

本文只证明这些路径变量和挂载关系在代码中存在，不证明任何特定服务器已按这些值配置完成。

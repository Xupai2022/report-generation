# 快速开始

本文只记录当前仓库中可以直接核实的最小可用流程，不沿用根目录旧 `README.md` 中已经过时的目录名、接口名和启动描述。

## 1. 先了解当前项目结构

可以直接从仓库确认：

- 后端源码与运行入口在 `mss_ai_ppt_sample_assets/backend/`
- 后端应用入口是 `mss_ai_ppt_sample_assets/backend/app.py`
- 前端源码目录是 `ui-src/`
- 前端构建产物输出到 `mss_ai_ppt_sample_assets/backend/frontend/`
- 后端启动后会把根路径 `/` 重定向到 `/ui/index.html`

对应事实来源：

- 后端根路由：`mss_ai_ppt_sample_assets/backend/app.py:131`
- 前端脚本：`ui-src/package.json:5`
- 前端构建输出目录：`ui-src/vite.config.ts:6`

## 2. 本地依赖

当前仓库可以直接确认的基础依赖包括：

- Python：用于运行 FastAPI 后端
- Node.js / npm：用于安装与构建 `ui-src/`
- LibreOffice：用于预览和 PDF 转换链路

其中，LibreOffice 不是“项目才能启动”的通用硬前置，而是与以下能力直接相关：

- 报告预览图生成
- PDF 下载
- 详细健康检查中的 LibreOffice 状态项

预览链路来自：

- `mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:42`
- `mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:147`

## 3. 准备后端依赖

后端依赖文件位于：

- `mss_ai_ppt_sample_assets/backend/requirements.txt`

常见安装方式示例：

```bash
pip install -r mss_ai_ppt_sample_assets/backend/requirements.txt
```

如果你需要预览或 PDF 能力，还需要保证服务器或本机可以找到 `soffice`。程序会按以下顺序寻找 LibreOffice：

1. `LIBREOFFICE_PATH`
2. Windows 常见安装目录
3. `shutil.which("soffice")`

对应代码：`mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:69`

## 4. 准备前端依赖

前端源码在 `ui-src/`，可直接从 `package.json` 确认以下脚本：

```json
{
  "dev": "vite --host 0.0.0.0 --port 5173",
  "build": "vite build",
  "preview": "vite preview --host 0.0.0.0 --port 4173"
}
```

安装依赖：

```bash
cd ui-src
npm install
```

## 5. 准备 `.env`

后端会在启动时加载 `.env`：

- 默认从仓库根目录 `.env` 读取
- 也可以通过 `MSS_ENV_PATH` 指定其他文件路径

对应代码：`mss_ai_ppt_sample_assets/backend/config.py:6`

当前可以直接确认：

- 只有当 `ENABLE_LLM=true` 时，才强制要求 `OPENAI_API_KEY`
- 如果 `ENABLE_LLM=true` 但未提供 `OPENAI_API_KEY`，后端会在初始化配置时抛错

对应代码：`mss_ai_ppt_sample_assets/backend/config.py:139`

## 6. 构建前端

当前项目的真实交付方式不是单独部署一个 `frontend/` 目录，而是：

- 在 `ui-src/` 中开发
- 执行构建
- 直接把静态产物写入 `mss_ai_ppt_sample_assets/backend/frontend/`
- 最终由后端通过 `/ui`、`/assets`、`/i18n` 提供静态文件

构建命令：

```bash
cd ui-src
npm run build
```

构建输出路径来自：`ui-src/vite.config.ts:12`

## 7. 启动后端

当前仓库里可直接回指的启动方式有两个来源：

- `app.py` 中的本地入口：`uvicorn.run(app, host="0.0.0.0", port=8000)`
- `deploy/systemd/mss-ai-ppt.service` 中的参考命令：`python3 -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000`

因此，文档统一使用以下命令：

```bash
uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000
```

对应事实来源：

- `mss_ai_ppt_sample_assets/backend/app.py:323`
- `deploy/systemd/mss-ai-ppt.service:9`

## 8. 启动后优先检查哪些入口

后端成功启动后，可优先验证以下入口：

- `/`：应重定向到 `/ui/index.html`
- `/api`：应返回 API 概览
- `/docs`：FastAPI OpenAPI 文档
- `/redoc`：ReDoc 文档
- `/api/v1/system/health`：基础健康检查

其中：

- `/` -> `/ui/index.html`：`mss_ai_ppt_sample_assets/backend/app.py:131`
- `/api`：`mss_ai_ppt_sample_assets/backend/app.py:138`
- `/api/v1/system/health`：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py:16`

## 9. 最小可用业务闭环

当前仓库可直接确认的模板与输入 ID 为：

- 模板：`mss_classic_ops`、`mss_classic_ops_2`
- 输入：`classic_ops_dataxlsx`

来源：

- `mss_ai_ppt_sample_assets/backend/data/templates/catalog.json:1`
- `mss_ai_ppt_sample_assets/backend/data/inputs/catalog.json:1`

一个与当前代码一致的最小流程如下：

1. 在 `ui-src/` 执行 `npm install && npm run build`
2. 启动后端：`uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000`
3. 打开浏览器访问 `/`
4. 在 UI 中选择模板，例如 `mss_classic_ops`
5. 在 UI 中选择输入，例如 `classic_ops_dataxlsx`
6. 选择至少一个 `focus_options`
7. 发起异步报告生成
8. 通过 WebSocket 或 `/api/v1/jobs/{job_id}/status` 观察进度
9. 生成完成后查看预览、下载 PPTX 或 PDF

## 10. 关键真实性边界

开始使用前，建议明确以下事实边界：

- 当前生成接口是 `POST /api/v1/reports`，不是旧文档中的同步 `/generate`
- 报告生成是异步任务，会返回 `job_id`
- `focus_options` 是必填且至少一个
- Excel 上传是输入准备步骤，不等于自动生成报告
- PDF 与预览都依赖 LibreOffice 相关链路
- 前端真实源码目录是 `ui-src/`，不是旧 `frontend`

如果后续实现发生变化，优先重新核对：

- `mss_ai_ppt_sample_assets/backend/app.py`
- `mss_ai_ppt_sample_assets/backend/config.py`
- `mss_ai_ppt_sample_assets/backend/routers/v1/reports.py`
- `ui-src/package.json`
- `ui-src/vite.config.ts`

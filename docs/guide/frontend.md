# 前端说明

本文围绕当前 `ui-src/` 的真实开发与构建方式整理，只记录能被当前仓库直接证明的事实。

## 1. 前端源码与部署产物的关系

当前可以直接确认：

- 前端源码目录是 `ui-src/`
- 构建输出目录不是 `ui-src/dist/`
- 构建产物会直接写入 `mss_ai_ppt_sample_assets/backend/frontend/`
- 后端运行时再把这些产物挂载到 `/ui`、`/assets`、`/i18n`

这意味着当前项目的真实模型是：

- 源码在 `ui-src/`
- 部署静态产物在 `backend/frontend/`
- 最终由 FastAPI 托管，而不是独立部署一个前端服务

来源：

- `ui-src/vite.config.ts:6`
- `mss_ai_ppt_sample_assets/backend/app.py:100`

## 2. 技术栈

从 `ui-src/package.json` 可直接确认：

- 使用 React
- 使用 TypeScript
- 使用 Vite
- 运行时依赖包括 `react`、`react-dom`、`lucide-react`

来源：`ui-src/package.json:1`

## 3. npm 脚本

当前可直接确认的脚本为：

```json
{
  "dev": "vite --host 0.0.0.0 --port 5173",
  "build": "vite build",
  "preview": "vite preview --host 0.0.0.0 --port 4173"
}
```

因此前端常用命令是：

```bash
cd ui-src
npm install
npm run dev
npm run build
npm run preview
```

对应端口：

- dev：`5173`
- preview：`4173`

来源：`ui-src/package.json:5`

## 4. 多页面入口

当前 Vite 配置不是单入口 SPA，而是多页面构建（MPA）。

构建入口包括：

- `index.html`
- `login.html`
- `admin.html`

来源：`ui-src/vite.config.ts:16`

结合后端静态挂载，可以对应为：

- `/ui/index.html`
- `/ui/login.html`
- `/ui/admin.html`

来源：`mss_ai_ppt_sample_assets/backend/app.py:101`

## 5. 构建输出方式

`vite.config.ts` 中可直接确认：

- `outDir = ../mss_ai_ppt_sample_assets/backend/frontend`
- `emptyOutDir: false`
- JS / chunk / asset 输出到 `assets/[name]-[hash]...`

这说明：

- 构建本身就直接把文件输出到后端静态目录
- 当前没有额外复制脚本
- 构建不会先清空整个输出目录

来源：`ui-src/vite.config.ts:6`

## 6. 后端如何托管前端产物

当前后端会按目录存在情况挂载：

- `/ui` -> `backend/frontend/`
- `/assets` -> `backend/frontend/assets/`
- `/i18n` -> `backend/frontend/i18n/`

并且根路径 `/` 会跳转到 `/ui/index.html`。

这意味着在完整联调与部署场景中，页面访问通常不是直连 Vite dev server，而是走后端托管的静态页面。

来源：

- `mss_ai_ppt_sample_assets/backend/app.py:101`
- `mss_ai_ppt_sample_assets/backend/app.py:131`

## 7. 主工作台 `IndexApp`

当前主页面位于：

- `ui-src/src/pages/IndexApp.tsx`

从该文件可直接确认，主工作台负责：

- 加载模板与输入
- Excel 上传
- 发起报告生成
- 建立 WebSocket
- 轮询任务状态
- 展示预览图
- 导出 PPT / PDF
- 手动编辑 slide
- AI 重写单页
- 提交评分

这说明它是当前报告工作台，而不是简单的静态展示页。

来源：`ui-src/src/pages/IndexApp.tsx:217`

## 8. 主页面的关键交互点

### focus options

当前页面内置的三个关注焦点是：

- `business_protection`
- `vulnerability`
- `alert`

来源：`ui-src/src/pages/IndexApp.tsx:43`

### WebSocket

当前前端按如下逻辑生成 WebSocket 地址：

- 页面为 HTTPS 时使用 `wss:`
- 否则使用 `ws:`
- 路径固定为 `/ws/{client_id}`

来源：`ui-src/src/pages/IndexApp.tsx:110`

### 导出、AI 重写、编辑状态

页面中可以直接确认存在以下状态：

- 导出格式：`ppt` / `pdf`
- AI 重写弹层状态
- AI prompt 与 token 选择状态
- slide 编辑状态与修改字段缓存
- 评分状态

来源：`ui-src/src/pages/IndexApp.tsx:251`

## 9. API 调用方式

前端主 API 封装位于：

- `ui-src/src/api/main.ts`
- `ui-src/src/api/client.ts`

当前主工作台直接调用的核心接口包括：

- `GET /api/v1/templates`
- `GET /api/v1/inputs`
- `POST /api/v1/inputs/excel`
- `GET /api/v1/templates/{template_id}/slides`
- `POST /api/v1/reports`
- `GET /api/v1/jobs/{job_id}/status`
- `GET /api/v1/reports/{job_id}/preview`
- `PATCH /api/v1/reports/{job_id}/slides`
- `POST /api/v1/reports/{job_id}/slides/ai-rewrite`
- `POST /api/v1/ratings/jobs/{job_id}/rating`

来源：`ui-src/src/api/main.ts:20`

## 10. 登录页与管理后台页

### 登录页

登录页组件：

- `ui-src/src/pages/LoginApp.tsx`

当前行为可直接确认：

- 输入用户名与密码
- 调用 `AdminApi.login(...)`
- 登录成功后跳转到 `/ui/admin.html`

来源：`ui-src/src/pages/LoginApp.tsx:10`

### 管理后台页

管理后台组件：

- `ui-src/src/pages/AdminApp.tsx`

当前行为可直接确认：

- 先调用 `AdminApi.verify()` 校验登录态
- 未认证时跳转到 `/ui/login.html`
- 可查看任务列表、过滤、详情、统计
- 可批量删除任务
- 可退出登录

来源：`ui-src/src/pages/AdminApp.tsx:30`

## 11. 开发模式与部署模式的区别

### 开发模式

适合单独调试前端源码：

```bash
cd ui-src
npm run dev
```

### 构建预览模式

适合仅预览打包结果：

```bash
cd ui-src
npm run preview
```

### 实际部署 / 联调模式

当前仓库更接近真实交付的方式是：

1. `cd ui-src && npm run build`
2. 启动后端
3. 访问后端托管的 `/ui/index.html`

因为部署静态产物实际上由 FastAPI 提供，而不是 Vite preview 提供。

## 12. 当前文档特别不再沿用的旧说法

为避免继续传播过时信息，本文不再使用：

- `frontend/` 作为前端源码目录
- 构建产物默认输出到 `dist/`
- “前端独立部署服务”为默认事实

当前真实目录与部署关系统一以 `ui-src/` -> `backend/frontend/` -> `/ui` 为准。

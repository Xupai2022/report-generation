# 使用指南

本文面向已经能够访问当前 UI 的使用者，只描述当前前后端代码里能够直接核实的用户路径。

## 1. 主界面当前支持的能力

主工作台位于 `ui-src/src/pages/IndexApp.tsx`。从当前实现可直接确认，页面支持：

- 加载模板列表与输入列表
- 选择模板
- 选择关注焦点 `focus_options`
- 上传 `.xlsx` 作为自定义输入
- 发起异步报告生成
- 通过 WebSocket 与轮询跟踪任务进度
- 查看预览图
- 导出 PPTX / PDF
- 手动编辑 slide 内容
- 对单页执行 AI 重写
- 对任务提交评分

相关事实来源：

- focus options：`ui-src/src/pages/IndexApp.tsx:43`
- WebSocket URL：`ui-src/src/pages/IndexApp.tsx:110`
- 导出与 AI 重写状态：`ui-src/src/pages/IndexApp.tsx:259`

## 2. 模板与输入从哪里来

### 模板列表

模板列表接口：

- `GET /api/v1/templates`

模板 slide 元数据接口：

- `GET /api/v1/templates/{template_id}/slides`

当前仓库里可以直接确认的模板 ID：

- `mss_classic_ops`
- `mss_classic_ops_2`

来源：

- 路由注册：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:17`
- 模板 catalog：`mss_ai_ppt_sample_assets/backend/data/templates/catalog.json:1`
- 模板 slide 元数据：`mss_ai_ppt_sample_assets/backend/routers/v1/templates.py:90`

### 输入列表

输入列表接口：

- `GET /api/v1/inputs`

输入详情接口：

- `GET /api/v1/inputs/{input_id}`

当前仓库里可以直接确认的内置输入 ID：

- `classic_ops_dataxlsx`

来源：

- 路由注册：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py:18`
- 输入 catalog：`mss_ai_ppt_sample_assets/backend/data/inputs/catalog.json:1`

## 3. 关注焦点 `focus_options`

创建报告时，`focus_options` 是必填字段，且至少要有一项。

后端允许的标准值为：

- `vulnerability`
- `alert`
- `business_protection`

前端当前提供的三个选项正对应这三个值。

相关来源：

- 前端选项定义：`ui-src/src/pages/IndexApp.tsx:43`
- 请求模型约束：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:33`
- 校验逻辑：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:66`

## 4. Excel 上传的真实边界

上传接口：

- `POST /api/v1/inputs/excel`

当前实现可以直接确认：

- 只接受 `.xlsx`
- 上传成功后会返回 `session_id`
- 前端会把当前输入切换为 `custom`
- 上传不是自动生成报告，后续仍需显式调用 `POST /api/v1/reports`

相关来源：

- 前端 API：`ui-src/src/api/main.ts:23`
- 后端上传路由：`mss_ai_ppt_sample_assets/backend/routers/v1/inputs.py:197`
- 页面上传逻辑：`ui-src/src/pages/IndexApp.tsx`

## 5. 报告生成是异步任务，不是同步下载

创建报告接口：

- `POST /api/v1/reports`

当前实现的真实行为：

- 返回 `202 Accepted`
- 返回 `job_id` 与 `session_id`
- 后台异步继续执行生成
- 客户端通过 WebSocket 或 `/api/v1/jobs/{job_id}/status` 观察进度

这意味着当前系统不应再描述成旧式同步 `/generate`。

相关来源：

- 创建报告路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:430`
- 异步处理函数：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:245`
- 任务状态接口：`mss_ai_ppt_sample_assets/backend/routers/v1/jobs.py:28`
- WebSocket 入口：`mss_ai_ppt_sample_assets/backend/app.py:116`

### 一个与当前 catalog 一致的最小示例

```json
{
  "input_id": "classic_ops_dataxlsx",
  "template_id": "mss_classic_ops",
  "focus_options": ["vulnerability"],
  "use_mock": false,
  "use_rag": true
}
```

如果是 Excel 上传后的自定义流程，通常还会带上：

```json
{
  "input_id": "custom",
  "session_id": "<上传返回的 session_id>"
}
```

## 6. 任务进度如何查看

当前前端会同时使用两条通路：

- WebSocket：`/ws/{client_id}`
- HTTP 轮询：`GET /api/v1/jobs/{job_id}/status`

状态模型中可直接确认的状态值包括：

- `pending`
- `running`
- `completed`
- `failed`
- `cancelled`

相关来源：

- 前端 WebSocket URL：`ui-src/src/pages/IndexApp.tsx:110`
- 任务状态接口：`ui-src/src/api/main.ts:59`
- 状态枚举：`mss_ai_ppt_sample_assets/backend/models/job_state.py:14`

## 7. 预览、下载与导出

### 预览

预览接口：

- `GET /api/v1/reports/{job_id}/preview`

预览返回的是图片 URL 列表，底层不是前端自己生成，而是后端基于已生成报告执行：

- PPTX -> LibreOffice -> PDF -> PNG

来源：

- 预览路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:739`
- 预览链路：`mss_ai_ppt_sample_assets/backend/modules/preview_generator.py:42`

### 下载 PPTX

接口：

- `GET /api/v1/reports/{job_id}/download`

### 下载 PDF

接口：

- `GET /api/v1/reports/{job_id}/download-pdf`

PDF 下载不是纯前端行为，而是后端转换得到的文件，因此仍依赖 LibreOffice 能力。

来源：

- PPTX 下载：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:648`
- PDF 下载：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:691`

## 8. 手动编辑与 AI 重写

### 手动编辑 slide

接口：

- `PATCH /api/v1/reports/{job_id}/slides`

请求体核心字段：

- `slides`
- 每项含 `slide_key`
- 每项含 `new_content`

真实语义是：更新 slide 占位数据并重新渲染报告、刷新预览，而不是在浏览器里直接编辑二进制 PPT。

来源：

- 请求模型：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:113`
- 更新路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:790`

### AI 重写单页

接口：

- `POST /api/v1/reports/{job_id}/slides/ai-rewrite`

当前实现可直接确认：

- 这是**单页 AI 重写**
- `slide_key` 必填
- `user_prompt` 必填
- `target_tokens` 可选
- 若不传 `target_tokens`，后端会改写该页所有可 AI 生成的占位字段

来源：

- 请求模型：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:135`
- AI 重写路由：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py:887`
- 当前页 AI token 元数据：`mss_ai_ppt_sample_assets/backend/routers/v1/templates.py:100`

## 9. 评分

当前普通用户评分接口为：

- `POST /api/v1/ratings/jobs/{job_id}/rating`
- `GET /api/v1/ratings/jobs/{job_id}/rating`
- `GET /api/v1/ratings/jobs/{job_id}/can-rate`

其中可直接确认：

- 点赞：`liked`
- 点踩：`disliked`
- 点踩评论至少 10 个字符

来源：

- 评分路由：`mss_ai_ppt_sample_assets/backend/routers/v1/ratings.py:30`
- 评分请求模型：`mss_ai_ppt_sample_assets/backend/schemas/requests.py:215`

## 10. 四条真实用户路径

### 路径 A：使用内置输入生成报告

1. 打开 UI。
2. 选择模板，例如 `mss_classic_ops`。
3. 选择输入，例如 `classic_ops_dataxlsx`。
4. 选择至少一个关注焦点。
5. 提交 `POST /api/v1/reports`。
6. 通过 WebSocket 或状态接口跟踪进度。
7. 生成完成后查看预览并下载 PPTX 或 PDF。

### 路径 B：先上传 Excel，再生成报告

1. 调用 `POST /api/v1/inputs/excel` 上传 `.xlsx`。
2. 记录返回的 `session_id`。
3. 前端切换到 `custom` 输入。
4. 选择模板与关注焦点。
5. 提交 `POST /api/v1/reports`，并带上 `input_id: "custom"` 与上传返回的 `session_id`。
6. 等待任务完成后查看预览与导出结果。

### 路径 C：生成后再手动修改

1. 先完成一次正常生成。
2. 在编辑区选择需要修改的 slide。
3. 编辑占位字段内容。
4. 调用 `PATCH /api/v1/reports/{job_id}/slides`。
5. 等待重新渲染与预览刷新。
6. 确认后再导出。

### 路径 D：生成后对单页执行 AI 重写

1. 先完成一次正常生成。
2. 切换到目标页。
3. 打开 AI 重写入口。
4. 输入 `user_prompt`。
5. 选择要重写的 token，或让后端处理当前页全部可改写 token。
6. 调用 `POST /api/v1/reports/{job_id}/slides/ai-rewrite`。
7. 等待预览刷新后再导出。

## 11. 本文特别删除了哪些旧说法

为避免失真，本文不再沿用以下旧式表述：

- 同步 `/generate`
- 笼统 `/rewrite`
- 笼统 `/download`
- 使用历史模板 ID 或输入 ID 作为示例

当前文档统一以 `/api/v1/...`、异步 job 模型、真实 catalog ID 为准。

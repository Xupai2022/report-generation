# 项目知识库

本文档站点只记录**能被当前仓库代码、配置、catalog 数据与部署样例直接证明**的信息。

## 真实性原则

维护本文档时统一遵循以下规则：

- 只记录可从代码、配置、catalog、部署样例直接核实的事实。
- 对依赖服务器环境的内容，明确标注“需按部署环境校验”。
- 不复用根目录 `README.md` 中已经与现状不一致的接口、路径或启动方式。
- 若代码与旧文档冲突，以当前仓库实现为准。

## 知识库范围

当前知识库覆盖以下主题：

- 快速开始：本地最小可用启动与验证流程
- 使用指南：模板、输入、异步生成、预览、导出、改写、评分
- 架构说明：前后端边界、异步任务、SlideSpec、预览链路、落盘位置
- 后端说明：配置、目录约定、API 前缀、静态挂载、运行时边界
- 前端说明：`ui-src/` 开发方式、多页面入口、构建产物与后端托管关系
- 运维部署：服务器运行方式、运行时配置、安全注意事项
- 排障与诊断：健康检查、日志、预览、RAG、任务复用与启动清理
- 参考附录：环境变量、API 索引、运行时路径、systemd 样例说明

## 阅读入口

- [快速开始](quickstart/getting-started.md)
- [使用指南](guide/usage.md)
- [架构说明](guide/architecture.md)
- [数据接入迁移方案](guide/data-ingestion-migration-plan.md)
- [后端说明](guide/backend.md)
- [前端说明](guide/frontend.md)
- [运维部署](guide/deployment.md)
- [排障与诊断](guide/troubleshooting.md)
- [环境变量清单](appendix/env-vars.md)
- [API 端点索引](appendix/api-endpoints.md)
- [运行时路径](appendix/runtime-paths.md)
- [部署样例说明](appendix/deployment-sample.md)

## 推荐阅读顺序

首次接手仓库时，建议按以下顺序阅读：

1. [快速开始](quickstart/getting-started.md)
2. [使用指南](guide/usage.md)
3. [架构说明](guide/architecture.md)
4. [数据接入迁移方案](guide/data-ingestion-migration-plan.md)
5. [后端说明](guide/backend.md)
6. [前端说明](guide/frontend.md)
7. [运维部署](guide/deployment.md)

## 主要事实来源

本文档站点主要回指以下真实来源：

- 配置与路径：`mss_ai_ppt_sample_assets/backend/config.py`
- 应用入口与挂载：`mss_ai_ppt_sample_assets/backend/app.py`
- V1 路由总表：`mss_ai_ppt_sample_assets/backend/routers/v1/__init__.py`
- 报告生成与改写：`mss_ai_ppt_sample_assets/backend/routers/v1/reports.py`
- 系统检查与日志：`mss_ai_ppt_sample_assets/backend/routers/v1/system.py`
- 请求模型：`mss_ai_ppt_sample_assets/backend/schemas/requests.py`
- 前端开发与构建：`ui-src/package.json`、`ui-src/vite.config.ts`
- 主工作台：`ui-src/src/pages/IndexApp.tsx`
- 模板与输入 catalog：`backend/data/templates/catalog.json`、`backend/data/inputs/catalog.json`
- 部署样例：`deploy/systemd/mss-ai-ppt.service`

## 维护说明

如果后续代码发生变化，更新文档时应优先重新核对：

- 命令是否仍可执行
- 路由是否仍存在于 `/api/v1/...`
- 示例 ID 是否仍存在于 catalog
- 前端构建输出路径与后端静态挂载是否一致
- 服务器样例中的路径是否仅为环境专用值

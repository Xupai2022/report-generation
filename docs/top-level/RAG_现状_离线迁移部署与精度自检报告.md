# RAG 现状、离线迁移部署与精度自检报告

## 1. 目的与范围

本文基于当前仓库代码与配置，给出：

- 当前 RAG 实现的完整说明（索引、检索、注入 Prompt 的链路）
- 迁移到不能联网的内网 Windows 环境的部署方案
- 迁移到不能联网的 CentOS 环境的部署方案
- 对当前 RAG 相关设置的自检结论，重点回答“为什么会检索出额外文本、精度是否偏低”

说明：本文完全依据当前代码实现，而不是历史文档假设。

## 2. 当前 RAG 实现详解

### 2.1 代码入口与职责

- RAG 核心服务：`mss_ai_ppt_sample_assets/backend/services/rag_service.py`
- RAG API 路由：`mss_ai_ppt_sample_assets/backend/routers/v1/rag.py`
- 配置来源：`mss_ai_ppt_sample_assets/backend/config.py`
- 生成流程接入点：`mss_ai_ppt_sample_assets/backend/services/report_service.py`
- Prompt 注入点：`mss_ai_ppt_sample_assets/backend/modules/llm_orchestrator.py`

### 2.2 能力边界（当前版本）

当前 RAG 支持：

- 文档类型：`.pdf/.docx/.md/.markdown/.txt`
- 向量库：Qdrant（本地路径模式或 URL 模式）
- 向量模型：`SentenceTransformer`
- 检索模式：单路 Dense Vector 相似度检索（Cosine）
- 过滤维度：`template_id`（匹配 `[template_id, "__all__"]`）

当前不支持：

- BM25/关键词召回
- reranker（二阶段重排）
- MMR 去冗余
- 按 slide_key 的向量过滤（只有 query 文本里含 slide 信息，但不是硬过滤）

### 2.3 索引构建流程

`POST /api/v1/rag/index/build` 调用 `RAGService.build_index()`，流程如下：

1. 检查 `RAG_ENABLED`
2. 初始化嵌入模型与 Qdrant（懒加载）
3. 遍历 `source_dir` 下支持格式文件
4. 文本抽取：
   - PDF：按页抽取
   - DOCX：解析 `word/document.xml`
   - 文本类：UTF-8 读取
5. 分块 `_split_text()`：
   - 采用“字符窗口近似 token”切分（`chars_per_token=2`）
   - 默认块大小：`RAG_CHUNK_SIZE_TOKENS=600`（约 1200 字符）
   - 默认重叠：`RAG_CHUNK_OVERLAP_TOKENS=80`（约 160 字符）
6. 每个 chunk 生成 `chunk_id/point_id`，写入 payload（source_file/page_no/template_id 等）
7. 分批 upsert 到 Qdrant
8. 写入 `outputs/rag/index_meta.json`

### 2.4 检索流程

`POST /api/v1/rag/query` 或生成/改写内部调用 `RAGService.query()`，流程如下：

1. 清洗 query 文本（压缩空白）
2. query 向量化
3. 在 Qdrant 检索 `top_k`（默认 8）
4. 依据 `min_score`（默认 0.30）做阈值裁剪
5. 组装：
   - `hits`：命中的文本与元数据
   - `retrieval_trace`：追踪信息
   - `context`：把命中文本拼成可注入 Prompt 的证据段
   - `retrieval_stats`：命中统计

### 2.5 生成/改写如何使用 RAG

- 生成：`retrieve_for_generation()`
  - 优先按每页构建结构化查询（`QueryBuilderV2.build_generation_query`）
  - 每页检索后按 `rag_max_context_chars_per_slide` 截断
  - 汇总为 `context_by_slide`
- 改写：`retrieve_for_rewrite()`
  - 构造 rewrite 结构化查询
  - 返回单页上下文

### 2.6 Prompt 注入策略

`llm_orchestrator.py` 中将检索内容作为“辅助上下文”注入，规则是：

- 文案明确写了“若与结构化数据冲突，以结构化数据优先”
- 批量生成时还会按预算比例裁剪 RAG 上下文（`rag_prompt_budget_ratio/min/max`）

### 2.7 运行期调试与审计

每次生成/改写检索会在 session 下落盘：

- `rag_query_*.md`：查询文本
- `rag_retrieval_*.md`：命中内容、trace、stats

这对定位“为什么检索进来额外文本”非常关键。

## 3. 当前配置与真实状态

### 3.1 当前 `.env`（已生效）中的关键项

- `RAG_ENABLED=true`
- `RAG_VECTOR_BACKEND=qdrant`
- `RAG_QDRANT_COLLECTION=kb_chunks`
- `RAG_EMBED_MODEL=mss_ai_ppt_sample_assets/backend/models/bge-small-zh-v1.5`（本地模型路径）
- `RAG_TOP_K=8`
- `RAG_MAX_CONTEXT_CHARS=4000`
- `RAG_HF_LOCAL_FILES_ONLY=true`

### 3.2 代码默认值（若 `.env` 未设置）

- `RAG_MIN_SCORE=0.30`
- `RAG_MAX_CONTEXT_CHARS_PER_SLIDE=1600`
- `RAG_PROMPT_BUDGET_RATIO=0.20`
- `RAG_PROMPT_BUDGET_MIN_TOKENS=400`
- `RAG_PROMPT_BUDGET_MAX_TOKENS=2500`
- `RAG_CHUNK_SIZE_TOKENS=600`
- `RAG_CHUNK_OVERLAP_TOKENS=80`

### 3.3 本地自检到的运行状态

分两种 Python 环境看：

- 系统 Python：会出现 `sentence-transformers is not available`，说明该解释器未装齐依赖。
- 项目 `.venv`：依赖可用，但当前返回 `ready=false`，原因为本地 Qdrant 存储目录被其他实例占用：`Storage folder ... is already accessed by another instance of Qdrant client`。

元数据里显示最近索引信息存在（`indexed_chunks_in_last_run=5`）。

结论：当前核心问题不是模型库缺失，而是“本地 Qdrant path 模式的并发锁冲突”。如果要多进程/多实例并发访问，建议改为 `RAG_QDRANT_URL` 指向独立 Qdrant 服务。

## 4. 精度与“额外文本”问题自检

## 4.1 结论先行

是的，按当前实现，检索结果“夹带额外文本”的概率偏高，属于实现特性导致，不是偶发。

### 4.2 主要原因（按影响排序）

1. 单文档大而全 + 固定窗口切块，天然混主题
- 当前 `kb.md` 被切成 5 个大 chunk（约每块 1200 字符），每块通常包含多个章节主题。
- 一旦命中某块，会把该块里与当前问题无关的段落一并带出。

2. 仅 Dense 检索，无重排
- 现有流程是“向量检索 -> 分数阈值过滤”，没有 reranker 做精排。
- 语义相关但不够精确的段落会保留，尤其在企业知识库术语高度相似时。

3. query 本身是结构化大 JSON，带入了较多字段
- `QueryBuilderV2` 会把 `ai_instructions/facts/focus/context_policy` 等统一编码到一个 query。
- 优点是覆盖信息全；缺点是召回面变宽，可能拉到“同模板不同页”的通用描述。

4. 过滤条件过粗
- 当前只按 `template_id` 过滤，不按 `slide_key/context_policy/doc_section` 做硬过滤。
- 即使某页是 `local_only`，检索层面仍可能命中同模板其他页知识。

5. 默认阈值偏宽
- `RAG_MIN_SCORE=0.30` + `RAG_TOP_K=8` 对中文通用知识库场景偏“召回优先”。
- 这会提升命中率，但会增加边缘相关文本进入 context 的概率。

6. chunk 后处理较少
- `_build_context()` 仅做长度截断，不做去重、句级提纯、标题噪声清理。

### 4.3 还有两个实现一致性问题

1. `docs/top-level/RAG_DEPLOYMENT.md` 与当前接口不一致
- 文档示例里含 `tenant_id/tags/rag_scope` 字段，但当前 `RAGBuildIndexRequest/RAGQueryRequest` 不支持这些字段。
- 直接照旧文档调用会报参数错误。

2. 测试中存在历史字段痕迹
- `tests/test_rag_integration.py` 里仍出现 `rag_scope` 断言，但当前 `CreateReportRequest` 已无该字段。

## 5. 精度优化建议（在现有架构上可落地）

### 5.1 先做配置级优化（低成本）

建议先试以下参数组合：

- `RAG_TOP_K`：`8 -> 4` 或 `5`
- `RAG_MIN_SCORE`：`0.30 -> 0.45`（可在 `0.40~0.55` 网格测试）
- `RAG_CHUNK_SIZE_TOKENS`：`600 -> 260~360`
- `RAG_CHUNK_OVERLAP_TOKENS`：`80 -> 40~60`
- `RAG_MAX_CONTEXT_CHARS_PER_SLIDE`：`1600 -> 900~1200`

预期：减少“单块混入太多主题”与“低相关段落被拼入 context”。

### 5.2 建议的代码级增强（中成本）

1. 元数据细化
- 入库时新增 `slide_key/topic_tag/section`（可由文档结构或文件名映射）
- 检索时增加 must filter，避免跨主题污染

2. 二阶段重排
- 向量召回 topN（如 20）后，用 reranker 取 topM（如 4）
- 可显著降低“看起来相关但不够精确”的 chunk

3. 句级提纯
- 对命中 chunk 做句级切分，只保留与 query 相似度最高的句段
- 拼接 context 时不直接塞整块

4. 多路召回融合
- Dense + BM25 融合（RRF）
- 对包含固定术语/指标名的企业报告场景通常更稳

## 6. 迁移到不能联网的内网 Windows：部署方案

## 6.1 方案建议

Windows 内网建议优先用“本地 Qdrant 路径模式”（不依赖独立 Qdrant 服务）：

- 不设置 `RAG_QDRANT_URL`
- 设置 `RAG_QDRANT_PATH` 指向本地目录

这样部署最简单，适合单机或小规模。

### 6.2 在有网机器准备离线资产

在可联网机器执行：

```powershell
python mss_ai_ppt_sample_assets/backend/scripts/rag_prepare_offline_assets.py `
  --output-dir offline_assets `
  --embed-model BAAI/bge-small-zh-v1.5
```

如需在 CentOS 侧用 Docker Qdrant，再加：

```powershell
--include-docker-image
```

产物至少包括：

- `offline_assets/wheelhouse/`
- `offline_assets/models/`
- `offline_assets/requirements.txt`
- `offline_assets/rag.env.example`

### 6.3 拷贝到内网 Windows

将以下目录整体拷贝到内网机：

- 项目代码目录
- `offline_assets`

### 6.4 内网 Windows 安装依赖（离线）

```powershell
cd <project_root>
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install --no-index --find-links .\offline_assets\wheelhouse -r .\offline_assets\requirements.txt
```

### 6.5 内网 Windows 推荐 `.env`（离线）

```env
RAG_ENABLED=true
RAG_VECTOR_BACKEND=qdrant
RAG_QDRANT_COLLECTION=kb_chunks
RAG_QDRANT_PATH=./mss_ai_ppt_sample_assets/backend/outputs/rag/qdrant
RAG_EMBED_MODEL=./offline_assets/models/BAAI__bge-small-zh-v1.5
RAG_HF_LOCAL_FILES_ONLY=true
RAG_SOURCE_DIR=./mss_ai_ppt_sample_assets/backend/data/rag_docs

# 先用保守精度参数
RAG_TOP_K=5
RAG_MIN_SCORE=0.45
RAG_CHUNK_SIZE_TOKENS=320
RAG_CHUNK_OVERLAP_TOKENS=48
RAG_MAX_CONTEXT_CHARS=3000
RAG_MAX_CONTEXT_CHARS_PER_SLIDE=1100
```

说明：`RAG_EMBED_MODEL` 需与实际离线模型目录一致。脚本默认目录名是 `模型名中的 / 替换为 __`。

### 6.6 构建索引与验证

启动服务后执行：

```powershell
curl.exe -X POST http://127.0.0.1:8000/api/v1/rag/index/build `
  -H "Content-Type: application/json" `
  -d '{"source_dir":"mss_ai_ppt_sample_assets/backend/data/rag_docs","template_id":"mss_classic_ops","reset_collection":true}'
```

诊断检索：

```powershell
curl.exe -X POST http://127.0.0.1:8000/api/v1/rag/query `
  -H "Content-Type: application/json" `
  -d '{"query_text":"漏洞闭环率提升的业务价值","template_id":"mss_classic_ops","scene":"diagnostic","top_k":5}'
```

检查点：

- `/api/v1/rag/index/status` 的 `ready=true`
- `kept_hits` 是否稳定 >0 且文本主题相关
- session 下 `rag_retrieval_*.md` 是否仍出现明显跨主题文本

## 7. 迁移到不能联网的 CentOS：部署方案

### 7.1 前置准备（有网机器）

同第 6.2，建议执行：

```bash
python mss_ai_ppt_sample_assets/backend/scripts/rag_prepare_offline_assets.py \
  --output-dir offline_assets \
  --include-docker-image \
  --embed-model BAAI/bge-small-zh-v1.5
```

### 7.2 内网 CentOS 离线安装

```bash
cd /opt/report-generation
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --no-index --find-links ./offline_assets/wheelhouse -r ./offline_assets/requirements.txt
```

或使用仓库脚本：

```bash
bash mss_ai_ppt_sample_assets/backend/scripts/rag_install_offline.sh ./offline_assets ./offline_assets/requirements.txt
```

### 7.3 Qdrant 部署（二选一）

方案 A：Docker（推荐）

```bash
docker image load -i ./offline_assets/qdrant.tar
docker run -d --name qdrant --restart always -p 6333:6333 -v /data/qdrant:/qdrant/storage qdrant/qdrant:latest
```

方案 B：二进制 + systemd

- 参考模板：`mss_ai_ppt_sample_assets/backend/scripts/rag-qdrant.service`

### 7.4 CentOS `.env` 推荐

```env
RAG_ENABLED=true
RAG_VECTOR_BACKEND=qdrant
RAG_QDRANT_URL=http://127.0.0.1:6333
RAG_QDRANT_COLLECTION=kb_chunks
RAG_EMBED_MODEL=/opt/report-generation/offline_assets/models/BAAI__bge-small-zh-v1.5
RAG_HF_LOCAL_FILES_ONLY=true
RAG_SOURCE_DIR=/opt/report-generation/mss_ai_ppt_sample_assets/backend/data/rag_docs

RAG_TOP_K=5
RAG_MIN_SCORE=0.45
RAG_CHUNK_SIZE_TOKENS=320
RAG_CHUNK_OVERLAP_TOKENS=48
RAG_MAX_CONTEXT_CHARS=3000
RAG_MAX_CONTEXT_CHARS_PER_SLIDE=1100
```

### 7.5 后端服务化（systemd）

- 模板：`mss_ai_ppt_sample_assets/backend/scripts/rag-backend.service`
- `WorkingDirectory`、`EnvironmentFile`、`ExecStart` 改为你的实际路径

启动后执行：

```bash
curl -X POST http://127.0.0.1:8000/api/v1/rag/index/build \
  -H "Content-Type: application/json" \
  -d '{"source_dir":"/opt/report-generation/mss_ai_ppt_sample_assets/backend/data/rag_docs","template_id":"mss_classic_ops","reset_collection":true}'
```

## 8. 迁移与精度自检清单（建议上线前逐项确认）

- 依赖可用：`sentence-transformers`、`qdrant-client` 已安装
- 状态正常：`/api/v1/rag/index/status` 返回 `ready=true`
- 索引成功：`collection_total_chunks > 0`
- 命中质量：`kept_hits` 稳定，低相关文本占比可接受
- Prompt 可控：`rag_context_by_slide` 没有明显跨页污染
- 性能可控：索引耗时、查询耗时在目标范围内
- 文档一致性：不要再使用旧 `RAG_DEPLOYMENT.md` 中已失效字段

## 9. 本次自检结论摘要

- 当前代码的 RAG 主链路完整；在 `.venv` 下依赖齐全，但本机存在本地 Qdrant 目录锁冲突，导致状态为 `ready=false`。
- 你观察到的“检索出额外文本”在当前实现中是高概率现象，根因是“大块切分 + 单路 Dense 检索 + 过滤粒度粗 + 阈值偏宽”。
- 先调参（`top_k/min_score/chunk`）就能明显改善；若要进一步稳定，需要补 reranker 与元数据过滤。
- 现有 `docs/top-level/RAG_DEPLOYMENT.md` 与代码接口不一致，建议以后以本文步骤为准。

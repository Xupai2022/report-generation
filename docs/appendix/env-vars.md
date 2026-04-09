# 环境变量清单

本文按 `mss_ai_ppt_sample_assets/backend/config.py` 中当前实际读取的变量整理，不扩展仓库外的约定。

## 1. `.env` 读取位置相关

### `MSS_ENV_PATH`

- 作用：指定 `.env` 文件路径
- 默认值：仓库根目录 `.env`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:6`

## 2. 路径类变量

### `MSS_DATA_DIR`

- 作用：覆盖 `DATA_DIR`
- 默认值：`backend/data/`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:13`

### `MSS_OUTPUTS_DIR`

- 作用：覆盖 `OUTPUTS_DIR`
- 默认值：`backend/outputs/`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:18`

### `MSS_OUTPUTS_URL_PREFIX`

- 作用：控制 outputs 目录的 URL 前缀
- 默认值：`/outputs`
- 代码行为：会被标准化为带前导 `/` 的路径
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:28`

## 3. OpenAI / LLM 相关

### `OPENAI_API_KEY`

- 作用：OpenAI 兼容接口密钥
- 默认值：无
- 重要边界：仅当 `ENABLE_LLM=true` 时为必填
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:46`

### `OPENAI_BASE_URL`

- 作用：OpenAI 兼容接口基地址
- 默认值：无
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:47`

### `OPENAI_MODEL`

- 作用：模型名
- 默认值：`GLM4.7`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:48`

### `ENABLE_LLM`

- 作用：是否启用 LLM
- 默认值：`false`
- 代码行为：值为字符串 `true` 时启用
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:71`

### `LLM_CONNECT_TIMEOUT_SECONDS`

- 作用：LLM 连接超时
- 默认值：`5`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:50`

### `LLM_READ_TIMEOUT_SECONDS`

- 作用：LLM 读取超时
- 默认值：`240`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:53`

### `LLM_WRITE_TIMEOUT_SECONDS`

- 作用：LLM 写入超时
- 默认值：`15`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:56`

### `LLM_POOL_TIMEOUT_SECONDS`

- 作用：连接池等待超时
- 默认值：`5`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:59`

### `LLM_RETRY_ATTEMPTS`

- 作用：LLM 重试次数
- 默认值：`2`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:62`

### `LLM_RETRY_BACKOFF_MIN_SECONDS`

- 作用：LLM 重试退避最小等待时间
- 默认值：`1`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:65`

### `LLM_RETRY_BACKOFF_MAX_SECONDS`

- 作用：LLM 重试退避最大等待时间
- 默认值：`4`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:68`

### `LLM_DISABLE_LOCAL_ONLY_BATCH_SPLIT`

- 作用：控制 local_only batch split 行为
- 默认值：`false`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:74`

### `DEFAULT_LOCALE`

- 作用：默认语言环境
- 默认值：`zh-CN`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:77`

## 4. 预览 / Session / Job retention

### `PREVIEW_CLEANUP_DAYS`

- 作用：预览清理保留天数
- 默认值：`7`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:80`

### `SESSION_RETENTION_DAYS`

- 作用：session 目录保留天数
- 默认值：`7`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:83`

### `JOB_RETENTION_DAYS`

- 作用：job 保留天数
- 默认值：`7`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:91`

### `JOB_MAX_RETRIES`

- 作用：job 最大重试次数
- 默认值：`3`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:92`

## 5. Logging

### `LOG_LEVEL`

- 作用：日志级别
- 默认值：`INFO`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:86`

### `LOG_MAX_BYTES`

- 作用：单个日志文件最大字节数
- 默认值：`52428800`（50MB）
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:87`

### `LOG_BACKUP_COUNT`

- 作用：日志滚动备份数量
- 默认值：`10`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:88`

## 6. Admin auth

### `ADMIN_USERNAME`

- 作用：管理员用户名
- 默认值：`admin`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:95`

### `ADMIN_PASSWORD_HASH`

- 作用：管理员密码哈希
- 默认值：空字符串
- 重要边界：为空时后台登录会报“管理员认证未配置”
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:96`

### `ADMIN_SESSION_SECRET`

- 作用：SessionMiddleware 密钥
- 默认值：代码内默认字符串
- 部署建议：生产环境应替换
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:97`

### `ADMIN_SESSION_MAX_AGE`

- 作用：管理员 session 最大有效期（秒）
- 默认值：`604800`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:101`

## 7. RAG 相关

### `RAG_ENABLED`

- 作用：RAG 总开关
- 默认值：`false`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:104`

### `RAG_VECTOR_BACKEND`

- 作用：向量后端类型
- 默认值：`qdrant`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:105`

### `RAG_QDRANT_URL`

- 作用：Qdrant URL
- 默认值：无
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:106`

### `RAG_QDRANT_API_KEY`

- 作用：Qdrant API key
- 默认值：无
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:107`

### `RAG_QDRANT_COLLECTION`

- 作用：集合名
- 默认值：`kb_chunks`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:108`

### `RAG_QDRANT_PATH`

- 作用：本地 qdrant 路径
- 默认值：`outputs/rag/qdrant`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:109`

### `RAG_EMBED_MODEL`

- 作用：向量模型名
- 默认值：`BAAI/bge-small-zh-v1.5`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:110`

### `RAG_TOP_K`

- 作用：检索 top-k
- 默认值：`3`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:111`

### `RAG_CANDIDATE_TOP_K`

- 作用：候选 top-k
- 默认值：`16`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:112`

### `RAG_MAX_CONTEXT_CHARS`

- 作用：总上下文字符数上限
- 默认值：`4000`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:113`

### `RAG_MAX_CONTEXT_CHARS_PER_SLIDE`

- 作用：每页上下文字符数上限
- 默认值：`1600`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:114`

### `RAG_MIN_SCORE`

- 作用：最小分数阈值
- 默认值：`0.30`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:115`

### `RAG_PROMPT_BUDGET_RATIO`

- 作用：prompt budget 比例
- 默认值：`0.20`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:116`

### `RAG_PROMPT_BUDGET_MIN_TOKENS`

- 作用：prompt budget 最小 token
- 默认值：`400`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:117`

### `RAG_PROMPT_BUDGET_MAX_TOKENS`

- 作用：prompt budget 最大 token
- 默认值：`2500`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:118`

### `RAG_CHUNK_SIZE_TOKENS`

- 作用：chunk 大小
- 默认值：`220`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:119`

### `RAG_CHUNK_OVERLAP_TOKENS`

- 作用：chunk overlap
- 默认值：`40`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:120`

### `RAG_HF_LOCAL_FILES_ONLY`

- 作用：是否仅使用本地 HF 文件
- 默认值：`false`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:121`

### `RAG_SOURCE_DIR`

- 作用：RAG 文档源目录
- 默认值：`DATA_DIR/rag_docs`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:122`

### `RAG_ENABLE_RERANK`

- 作用：是否启用 rerank
- 默认值：`true`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:123`

### `RAG_RERANK_MODEL`

- 作用：rerank 模型路径
- 默认值：`mss_ai_ppt_sample_assets/backend/models/bge-reranker-base`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:124`

### `RAG_RERANK_SCENES`

- 作用：启用 rerank 的场景列表
- 默认值：`diagnostic,generate,rewrite`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:128`

### `RAG_FINAL_TOP_K_PER_SLIDE`

- 作用：每页最终 top-k
- 默认值：`3`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:134`

### `RAG_COMMON_FALLBACK_TOP_K`

- 作用：通用 fallback top-k
- 默认值：`2`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:135`

### `RAG_SECTION_MIN_CHARS`

- 作用：section 最小字符数
- 默认值：`80`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:136`

### `RAG_PRELOAD_ON_STARTUP`

- 作用：启动时是否预热 RAG
- 默认值：`false`
- 来源：`mss_ai_ppt_sample_assets/backend/config.py:137`

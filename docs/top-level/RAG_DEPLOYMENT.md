# RAG Deployment Guide (Windows Dev -> Linux Intranet)

## 1. External network (Windows) prepare offline assets

```bash
python mss_ai_ppt_sample_assets/backend/scripts/rag_prepare_offline_assets.py ^
  --output-dir offline_assets ^
  --include-docker-image ^
  --embed-model BAAI/bge-small-zh-v1.5
```

Generated artifacts:

- `offline_assets/wheelhouse/` Python wheels
- `offline_assets/models/` embedding model snapshot
- `offline_assets/qdrant.tar` (optional, docker image)
- `offline_assets/requirements.txt`
- `offline_assets/rag.env.example`

## 2. Intranet Linux install

```bash
bash mss_ai_ppt_sample_assets/backend/scripts/rag_install_offline.sh /path/to/offline_assets /path/to/requirements.txt
```

## 3. Environment variables

Add to `.env`:

```env
RAG_ENABLED=true
RAG_VECTOR_BACKEND=qdrant
RAG_QDRANT_URL=http://127.0.0.1:6333
RAG_QDRANT_COLLECTION=kb_chunks
RAG_EMBED_MODEL=BAAI/bge-small-zh-v1.5
RAG_TOP_K=8
RAG_MAX_CONTEXT_CHARS=4000
RAG_HF_LOCAL_FILES_ONLY=true
```

## 4. Build index

```bash
curl -X POST http://127.0.0.1:8000/api/v1/rag/index/build \
  -H "Content-Type: application/json" \
  -d '{
    "source_dir": "/data/knowledge",
    "tenant_id": "classic_ops_dataxlsx",
    "template_id": "mss_classic_ops",
    "tags": ["domain_knowledge"],
    "reset_collection": true
  }'
```

## 5. Diagnostic retrieval

```bash
curl -X POST http://127.0.0.1:8000/api/v1/rag/query \
  -H "Content-Type: application/json" \
  -d '{
    "query_text": "高危漏洞治理策略",
    "tenant_id": "classic_ops_dataxlsx",
    "template_id": "mss_classic_ops",
    "scene": "diagnostic",
    "top_k": 8,
    "rag_scope": ["domain_knowledge"]
  }'
```

## 6. Optional systemd deployment

- Qdrant service template: `mss_ai_ppt_sample_assets/backend/scripts/rag-qdrant.service`
- Backend service template: `mss_ai_ppt_sample_assets/backend/scripts/rag-backend.service`

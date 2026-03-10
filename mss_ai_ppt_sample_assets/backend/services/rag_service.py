from __future__ import annotations

import hashlib
import json
import logging
import threading
import uuid
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence
from zipfile import ZipFile

import fitz
from lxml import etree

from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.models.inputs import TenantInput

logger = logging.getLogger(__name__)


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _sanitize_filename(value: str) -> str:
    text = re.sub(r"[^0-9A-Za-z._-]+", "_", value or "").strip("._")
    return text or "unknown"


def _clean_whitespace(text: str) -> str:
    compact = " ".join((text or "").replace("\u3000", " ").split())
    return compact.strip()


def _estimate_chars_per_token() -> int:
    # Simple and stable approximation for mixed zh/en content.
    return 2


@dataclass
class RetrievalResult:
    rag_used: bool
    context: str = ""
    retrieval_trace: List[Dict[str, Any]] = field(default_factory=list)
    hits: List[Dict[str, Any]] = field(default_factory=list)
    context_by_slide: Dict[str, str] = field(default_factory=dict)
    retrieval_stats: Dict[str, Any] = field(default_factory=dict)


@dataclass
class GenerationRetrievalTask:
    slide_key: str
    slide_title: str
    ai_tokens: List[str]
    ai_instructions: List[str]
    facts: Dict[str, Any]
    focus_options: List[str]
    context_policy: str = "auto"


@dataclass
class RewriteRetrievalTask:
    slide_key: str
    ai_tokens: List[str]
    ai_instructions: List[str]
    user_prompt: str
    facts: Dict[str, Any]
    current_slide_content: Dict[str, Any]


class QueryBuilderV2:
    """Construct compact retrieval queries aligned to generation tasks."""

    MAX_INSTRUCTION_CHARS = 260
    MAX_FACT_VALUE_CHARS = 220
    MAX_FACT_ITEMS = 18

    @staticmethod
    def _truncate(text: Any, limit: int) -> str:
        value = str(text or "").strip()
        if len(value) <= limit:
            return value
        return value[: max(0, limit - 1)].rstrip() + "…"

    @classmethod
    def _summarize_scalar(cls, value: Any) -> str:
        if value is None:
            return "null"
        if isinstance(value, bool):
            return "true" if value else "false"
        if isinstance(value, (int, float)):
            return str(value)
        if isinstance(value, str):
            return cls._truncate(value, cls.MAX_FACT_VALUE_CHARS)
        return cls._truncate(json.dumps(value, ensure_ascii=False), cls.MAX_FACT_VALUE_CHARS)

    @classmethod
    def _summarize_value(cls, value: Any) -> Any:
        if isinstance(value, dict):
            items = list(value.items())
            summary: Dict[str, Any] = {}
            for key, child in items[:8]:
                if isinstance(child, (dict, list)):
                    summary[str(key)] = cls._summarize_scalar(child)
                else:
                    summary[str(key)] = cls._summarize_scalar(child)
            if len(items) > 8:
                summary["_truncated_items"] = len(items) - 8
            return summary
        if isinstance(value, list):
            summary_items = [cls._summarize_scalar(item) for item in value[:6]]
            if len(value) > 6:
                summary_items.append(f"...(+{len(value) - 6} items)")
            return summary_items
        return cls._summarize_scalar(value)

    @classmethod
    def _normalize_facts(cls, facts: Dict[str, Any]) -> Dict[str, Any]:
        if not facts:
            return {}
        normalized: Dict[str, Any] = {}
        for idx, key in enumerate(sorted(facts.keys())):
            if idx >= cls.MAX_FACT_ITEMS:
                normalized["_truncated_facts"] = len(facts) - cls.MAX_FACT_ITEMS
                break
            normalized[key] = cls._summarize_value(facts[key])
        return normalized

    @classmethod
    def build_generation_query(
        cls,
        template_id: str,
        task: GenerationRetrievalTask,
    ) -> str:
        focus_text = ",".join(task.focus_options or [])
        instructions = [
            cls._truncate(item, cls.MAX_INSTRUCTION_CHARS)
            for item in task.ai_instructions
            if (item or "").strip()
        ]
        facts_payload = cls._normalize_facts(task.facts)
        payload = {
            "scene": "generate",
            "template": template_id,
            "slide_key": task.slide_key,
            "slide_title": task.slide_title,
            "focus": focus_text,
            "context_policy": task.context_policy,
            "ai_tokens": task.ai_tokens,
            "ai_instructions": instructions,
            "facts": facts_payload,
        }
        return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    @classmethod
    def build_rewrite_query(
        cls,
        template_id: str,
        task: RewriteRetrievalTask,
    ) -> str:
        instructions = [
            cls._truncate(item, cls.MAX_INSTRUCTION_CHARS)
            for item in task.ai_instructions
            if (item or "").strip()
        ]
        payload = {
            "scene": "rewrite",
            "template": template_id,
            "slide_key": task.slide_key,
            "ai_tokens": task.ai_tokens,
            "ai_instructions": instructions,
            "user_prompt": cls._truncate(task.user_prompt, 600),
            "current_slide": cls._normalize_facts(task.current_slide_content),
            "facts": cls._normalize_facts(task.facts),
        }
        return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


class RAGService:
    """RAG indexing and retrieval service with lazy dependency loading."""

    SUPPORTED_SUFFIXES = {".pdf", ".docx", ".md", ".markdown", ".txt"}

    def __init__(self):
        self._lock = threading.Lock()
        self._embedder = None
        self._qdrant_client = None
        self._qmodels = None
        self._embedding_dim: Optional[int] = None
        config.RAG_DIR.mkdir(parents=True, exist_ok=True)

    @property
    def enabled(self) -> bool:
        return bool(config.settings.rag_enabled)

    def _check_enabled(self) -> None:
        if not self.enabled:
            raise ValueError("RAG is disabled. Set RAG_ENABLED=true to use RAG features.")

    def _load_dependencies(self) -> None:
        if self._embedder is not None and self._qdrant_client is not None and self._qmodels is not None:
            return

        with self._lock:
            if self._embedder is None:
                try:
                    from sentence_transformers import SentenceTransformer
                except Exception as e:
                    raise RuntimeError(
                        "sentence-transformers is not available. Install backend requirements first."
                    ) from e

                model_ref = self._resolve_embed_model_ref(config.settings.rag_embed_model)
                self._embedder = SentenceTransformer(
                    model_ref,
                    device="cpu",
                    trust_remote_code=False,
                    local_files_only=config.settings.rag_hf_local_files_only,
                )

            if self._qdrant_client is None or self._qmodels is None:
                try:
                    from qdrant_client import QdrantClient
                    from qdrant_client.http import models as qmodels
                except Exception as e:
                    raise RuntimeError("qdrant-client is not available. Install backend requirements first.") from e

                if config.settings.rag_qdrant_url:
                    self._qdrant_client = QdrantClient(
                        url=config.settings.rag_qdrant_url,
                        api_key=config.settings.rag_qdrant_api_key or None,
                    )
                else:
                    local_path = Path(config.settings.rag_qdrant_path).resolve()
                    local_path.mkdir(parents=True, exist_ok=True)
                    self._qdrant_client = QdrantClient(path=str(local_path))
                self._qmodels = qmodels

            if self._embedding_dim is None:
                probe = self._encode_texts(["embedding dimension probe"])
                if not probe or not probe[0]:
                    raise RuntimeError("Failed to initialize embedding model.")
                self._embedding_dim = len(probe[0])

    @staticmethod
    def _resolve_embed_model_ref(model_ref: str) -> str:
        """Resolve local model path when configured as a relative path.

        Priority:
        1) absolute path if exists
        2) path relative to current working directory
        3) path relative to repository root
        4) raw model_ref (treated as huggingface repo id)
        """
        if not model_ref:
            return model_ref

        candidate = Path(model_ref)
        if candidate.is_absolute() and candidate.exists():
            return str(candidate)

        cwd_path = Path.cwd() / candidate
        if cwd_path.exists():
            return str(cwd_path.resolve())

        repo_path = config.REPO_ROOT_DIR / candidate
        if repo_path.exists():
            return str(repo_path.resolve())

        return model_ref

    def _encode_texts(self, texts: List[str]) -> List[List[float]]:
        vectors = self._embedder.encode(
            texts,
            normalize_embeddings=True,
            convert_to_numpy=True,
            show_progress_bar=False,
        )
        return vectors.tolist()

    def _ensure_collection(self, reset_collection: bool = False) -> None:
        self._load_dependencies()
        assert self._qdrant_client is not None
        assert self._qmodels is not None
        assert self._embedding_dim is not None

        collection = config.settings.rag_qdrant_collection
        qmodels = self._qmodels

        if reset_collection:
            try:
                self._qdrant_client.delete_collection(collection_name=collection)
            except Exception:
                pass

        try:
            self._qdrant_client.get_collection(collection_name=collection)
            return
        except Exception:
            pass

        self._qdrant_client.create_collection(
            collection_name=collection,
            vectors_config=qmodels.VectorParams(
                size=self._embedding_dim,
                distance=qmodels.Distance.COSINE,
            ),
        )

    def _read_meta(self) -> Dict[str, Any]:
        meta_path = config.RAG_META_FILE
        if not meta_path.exists():
            return {}
        try:
            return json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _write_meta(self, payload: Dict[str, Any]) -> None:
        meta_path = config.RAG_META_FILE
        meta_path.parent.mkdir(parents=True, exist_ok=True)
        meta_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _extract_docx_text(self, path: Path) -> str:
        try:
            with ZipFile(path, "r") as zf:
                raw = zf.read("word/document.xml")
        except Exception:
            return ""

        try:
            root = etree.fromstring(raw)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            nodes = root.xpath(".//w:t", namespaces=ns)
            texts = [node.text for node in nodes if node.text]
            return "\n".join(texts)
        except Exception:
            return ""

    def _split_text(self, text: str, page_no: Optional[int]) -> List[Dict[str, Any]]:
        cleaned = _clean_whitespace(text)
        if not cleaned:
            return []

        chunk_tokens = max(200, int(config.settings.rag_chunk_size_tokens))
        overlap_tokens = max(0, int(config.settings.rag_chunk_overlap_tokens))
        chars_per_token = _estimate_chars_per_token()
        chunk_chars = chunk_tokens * chars_per_token
        overlap_chars = min(overlap_tokens * chars_per_token, max(0, chunk_chars - 10))

        chunks: List[Dict[str, Any]] = []
        start = 0
        while start < len(cleaned):
            end = min(len(cleaned), start + chunk_chars)
            part = cleaned[start:end].strip()
            if part:
                chunks.append({"text": part, "page_no": page_no})
            if end >= len(cleaned):
                break
            start = max(0, end - overlap_chars)
        return chunks

    def _extract_chunks_from_file(self, path: Path) -> List[Dict[str, Any]]:
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            chunks: List[Dict[str, Any]] = []
            try:
                doc = fitz.open(path)
                for i in range(doc.page_count):
                    page_text = doc.load_page(i).get_text("text")
                    chunks.extend(self._split_text(page_text, page_no=i + 1))
                doc.close()
            except Exception as e:
                logger.warning("Failed to parse PDF %s: %s", path, e)
            return chunks

        if suffix == ".docx":
            return self._split_text(self._extract_docx_text(path), page_no=None)

        try:
            raw = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            raw = path.read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            logger.warning("Failed to read text file %s: %s", path, e)
            return []
        return self._split_text(raw, page_no=None)

    def _iter_source_files(self, source_dir: Path) -> Iterable[Path]:
        for path in sorted(source_dir.rglob("*")):
            if not path.is_file():
                continue
            if path.suffix.lower() in self.SUPPORTED_SUFFIXES:
                yield path

    @staticmethod
    def _get_nested_raw(raw: Dict[str, Any], path: str) -> Any:
        if not path:
            return None
        current: Any = raw
        for part in path.split("."):
            if isinstance(current, dict):
                current = current.get(part)
            elif isinstance(current, list) and part.isdigit():
                idx = int(part)
                current = current[idx] if 0 <= idx < len(current) else None
            else:
                return None
            if current is None:
                return None
        return current

    def _dump_query_markdown(
        self,
        *,
        session_id: Optional[str],
        scene: str,
        template_id: str,
        query_entries: List[Dict[str, Any]],
    ) -> None:
        """Persist built RAG query text into session directory for auditing/debugging."""
        if not session_id or not query_entries:
            return
        try:
            session_dir = config.SESSIONS_DIR / session_id
            session_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            filename = f"rag_query_{_sanitize_filename(scene)}_{ts}.md"
            path = session_dir / filename

            lines: List[str] = [
                "# RAG Query Dump",
                "",
                "## Metadata",
                f"- scene: `{scene}`",
                f"- template_id: `{template_id}`",
                f"- session_id: `{session_id}`",
                f"- query_count: `{len(query_entries)}`",
                "",
            ]
            for idx, item in enumerate(query_entries, start=1):
                slide_key = item.get("slide_key")
                lines.append(f"## Query {idx}")
                if slide_key:
                    lines.append(f"- slide_key: `{slide_key}`")
                if item.get("query_chars") is not None:
                    lines.append(f"- query_chars: `{item.get('query_chars')}`")
                lines.extend([
                    "",
                    "```text",
                    item.get("query_text", ""),
                    "```",
                    "",
                ])

            path.write_text("\n".join(lines), encoding="utf-8")
            logger.info("RAG query markdown dumped: %s", path)
        except Exception as e:
            logger.warning("Failed to dump RAG query markdown for session %s: %s", session_id, e)

    def _dump_retrieval_markdown(
        self,
        *,
        session_id: Optional[str],
        scene: str,
        template_id: str,
        retrieval_result: RetrievalResult,
    ) -> None:
        """Persist retrieval output into session directory for auditing/debugging."""
        if not session_id:
            return
        try:
            session_dir = config.SESSIONS_DIR / session_id
            session_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            filename = f"rag_retrieval_{_sanitize_filename(scene)}_{ts}.md"
            path = session_dir / filename

            def _json(obj: Any) -> str:
                return json.dumps(obj, ensure_ascii=False, indent=2)

            lines: List[str] = [
                "# RAG Retrieval Dump",
                "",
                "## Metadata",
                f"- scene: `{scene}`",
                f"- template_id: `{template_id}`",
                f"- session_id: `{session_id}`",
                f"- rag_used: `{bool(retrieval_result.rag_used)}`",
                f"- hits_count: `{len(retrieval_result.hits or [])}`",
                f"- trace_count: `{len(retrieval_result.retrieval_trace or [])}`",
                "",
                "## Context",
                "```text",
                retrieval_result.context or "",
                "```",
                "",
                "## Context By Slide",
                "```json",
                _json(retrieval_result.context_by_slide or {}),
                "```",
                "",
                "## Retrieval Stats",
                "```json",
                _json(retrieval_result.retrieval_stats or {}),
                "```",
                "",
                "## Retrieval Trace",
                "```json",
                _json(retrieval_result.retrieval_trace or []),
                "```",
                "",
                "## Hits",
                "```json",
                _json(retrieval_result.hits or []),
                "```",
                "",
            ]
            path.write_text("\n".join(lines), encoding="utf-8")
            logger.info("RAG retrieval markdown dumped: %s", path)
        except Exception as e:
            logger.warning("Failed to dump RAG retrieval markdown for session %s: %s", session_id, e)

    def _upsert_records(self, records: List[Dict[str, Any]]) -> None:
        self._load_dependencies()
        assert self._qdrant_client is not None
        assert self._qmodels is not None
        qmodels = self._qmodels

        if not records:
            return

        texts = [item["text"] for item in records]
        vectors = self._encode_texts(texts)
        points = [
            qmodels.PointStruct(
                id=item["point_id"],
                vector=vector,
                payload={
                    "chunk_id": item["chunk_id"],
                    "text": item["text"],
                    "source_file": item["source_file"],
                    "page_no": item["page_no"],
                    "tenant_id": item["tenant_id"],
                    "template_id": item["template_id"],
                    "doc_type": item["doc_type"],
                    "updated_at": item["updated_at"],
                },
            )
            for item, vector in zip(records, vectors)
        ]

        self._qdrant_client.upsert(
            collection_name=config.settings.rag_qdrant_collection,
            points=points,
            wait=True,
        )

    def build_index(
        self,
        source_dir: Optional[str] = None,
        template_id: Optional[str] = None,
        reset_collection: bool = True,
    ) -> Dict[str, Any]:
        self._check_enabled()
        self._ensure_collection(reset_collection=reset_collection)

        source = Path(source_dir or config.settings.rag_source_dir).resolve()
        if not source.exists() or not source.is_dir():
            raise ValueError(f"RAG source directory not found: {source}")

        tenant_value = "__all__"
        template_value = (template_id or "__all__").strip() or "__all__"
        records: List[Dict[str, Any]] = []
        docs_indexed = 0
        for file_path in self._iter_source_files(source):
            relative = str(file_path.relative_to(source)).replace("\\", "/")
            chunks = self._extract_chunks_from_file(file_path)
            if not chunks:
                continue
            docs_indexed += 1
            for idx, chunk in enumerate(chunks):
                text = chunk["text"]
                hash_input = f"{tenant_value}|{template_value}|{relative}|{chunk.get('page_no')}|{idx}|{text}"
                chunk_id = hashlib.sha1(hash_input.encode("utf-8")).hexdigest()
                point_id = str(uuid.UUID(hashlib.md5(hash_input.encode("utf-8")).hexdigest()))
                records.append(
                    {
                        "point_id": point_id,
                        "chunk_id": chunk_id,
                        "text": text,
                        "source_file": relative,
                        "page_no": chunk.get("page_no"),
                        "tenant_id": tenant_value,
                        "template_id": template_value,
                        "doc_type": file_path.suffix.lower().lstrip("."),
                        "updated_at": _utc_now_iso(),
                    }
                )

        batch_size = 32
        for i in range(0, len(records), batch_size):
            self._upsert_records(records[i:i + batch_size])

        count = 0
        try:
            assert self._qdrant_client is not None
            count = int(
                self._qdrant_client.count(
                    collection_name=config.settings.rag_qdrant_collection,
                    exact=True,
                ).count
            )
        except Exception:
            count = len(records)

        meta = {
            "rag_enabled": self.enabled,
            "collection": config.settings.rag_qdrant_collection,
            "last_indexed_at": _utc_now_iso(),
            "last_source_dir": str(source),
            "last_tenant_id": tenant_value,
            "last_template_id": template_value,
            "last_reset_collection": reset_collection,
            "indexed_docs": docs_indexed,
            "indexed_chunks_in_last_run": len(records),
            "collection_total_chunks": count,
        }
        self._write_meta(meta)
        return meta

    def update_index(
        self,
        source_dir: Optional[str] = None,
        template_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        return self.build_index(
            source_dir=source_dir,
            template_id=template_id,
            reset_collection=False,
        )

    def get_status(self) -> Dict[str, Any]:
        meta = self._read_meta()
        base = {
            "rag_enabled": self.enabled,
            "vector_backend": config.settings.rag_vector_backend,
            "collection": config.settings.rag_qdrant_collection,
            "top_k": config.settings.rag_top_k,
            "max_context_chars": config.settings.rag_max_context_chars,
            "max_context_chars_per_slide": config.settings.rag_max_context_chars_per_slide,
            "min_score": config.settings.rag_min_score,
            "embed_model": config.settings.rag_embed_model,
        }

        if not self.enabled:
            base.update({"ready": False, "reason": "RAG_DISABLED"})
            return {**base, **meta}

        try:
            self._ensure_collection(reset_collection=False)
            assert self._qdrant_client is not None
            count = int(
                self._qdrant_client.count(
                    collection_name=config.settings.rag_qdrant_collection,
                    exact=True,
                ).count
            )
            base.update({"ready": True, "collection_total_chunks": count})
        except Exception as e:
            base.update({"ready": False, "reason": str(e), "collection_total_chunks": 0})

        return {**base, **meta}

    def _build_query_filter(
        self,
        template_id: Optional[str],
    ):
        assert self._qmodels is not None
        qmodels = self._qmodels

        must = []
        if template_id:
            must.append(
                qmodels.FieldCondition(
                    key="template_id",
                    match=qmodels.MatchAny(any=[template_id, "__all__"]),
                )
            )

        if not must:
            return None
        return qmodels.Filter(must=must)

    def _build_context(self, hits: List[Dict[str, Any]], max_chars: int) -> str:
        lines: List[str] = []
        used = 0
        for idx, hit in enumerate(hits, start=1):
            prefix = (
                f"[{idx}] source={hit.get('source_file')}"
                f", page={hit.get('page_no')}, score={hit.get('score'):.4f}\n"
            )
            text = (hit.get("text") or "").strip()
            candidate = prefix + text + "\n"
            if used + len(candidate) > max_chars:
                break
            lines.append(candidate)
            used += len(candidate)
        return "\n".join(lines).strip()

    def query(
        self,
        query_text: str,
        template_id: Optional[str] = None,
        scene: str = "diagnostic",
        top_k: Optional[int] = None,
        min_score: Optional[float] = None,
    ) -> RetrievalResult:
        if not self.enabled:
            return RetrievalResult(rag_used=False, context="", retrieval_trace=[], hits=[])

        normalized_query = _clean_whitespace(query_text)
        if not normalized_query:
            return RetrievalResult(rag_used=False, context="", retrieval_trace=[], hits=[])

        self._ensure_collection(reset_collection=False)
        assert self._qdrant_client is not None

        vector = self._encode_texts([normalized_query])[0]
        query_filter = self._build_query_filter(template_id=template_id)
        limit = int(top_k or config.settings.rag_top_k)
        score_threshold = float(config.settings.rag_min_score if min_score is None else min_score)

        if hasattr(self._qdrant_client, "query_points"):
            response = self._qdrant_client.query_points(
                collection_name=config.settings.rag_qdrant_collection,
                query=vector,
                query_filter=query_filter,
                limit=max(1, limit),
                with_payload=True,
                with_vectors=False,
            )
            raw_hits = getattr(response, "points", []) or []
        else:
            # Backward compatibility for older qdrant-client versions.
            raw_hits = self._qdrant_client.search(
                collection_name=config.settings.rag_qdrant_collection,
                query_vector=vector,
                query_filter=query_filter,
                limit=max(1, limit),
                with_payload=True,
                with_vectors=False,
            )

        hits: List[Dict[str, Any]] = []
        trace: List[Dict[str, Any]] = []
        raw_count = 0
        dropped_count = 0
        for point in raw_hits:
            raw_count += 1
            payload = dict(getattr(point, "payload", {}) or {})
            score = float(getattr(point, "score", 0.0))
            if score < score_threshold:
                dropped_count += 1
                continue
            item = {
                "chunk_id": payload.get("chunk_id") or str(getattr(point, "id", "")),
                "score": score,
                "source_file": payload.get("source_file"),
                "page_no": payload.get("page_no"),
                "text": payload.get("text", ""),
                "tenant_id": payload.get("tenant_id"),
                "template_id": payload.get("template_id"),
                "doc_type": payload.get("doc_type"),
            }
            hits.append(item)
            trace.append(
                {
                    "chunk_id": item["chunk_id"],
                    "score": item["score"],
                    "source_file": item["source_file"],
                    "page_no": item["page_no"],
                    "scene": scene,
                }
            )

        context = self._build_context(hits, max_chars=config.settings.rag_max_context_chars)
        stats = {
            "scene": scene,
            "query_chars": len(normalized_query),
            "top_k": max(1, limit),
            "min_score": score_threshold,
            "raw_hits": raw_count,
            "kept_hits": len(hits),
            "dropped_hits": dropped_count,
            "top_score": max((item["score"] for item in hits), default=0.0),
        }
        return RetrievalResult(
            rag_used=bool(hits),
            context=context,
            retrieval_trace=trace,
            hits=hits,
            retrieval_stats=stats,
        )

    def _tenant_snapshot(self, tenant_input: TenantInput, max_chars: int = 1800) -> str:
        raw = json.dumps(tenant_input.raw, ensure_ascii=False)
        return raw[:max_chars]

    def _build_generation_tasks(
        self,
        tenant_input: TenantInput,
        template_descriptor: Any,
        focus_options: Optional[List[str]] = None,
    ) -> List[GenerationRetrievalTask]:
        tasks: List[GenerationRetrievalTask] = []
        raw = tenant_input.raw or {}
        focus = [item.strip() for item in (focus_options or []) if (item or "").strip()]

        for slide in getattr(template_descriptor, "slides", []):
            ai_placeholders = [ph for ph in slide.placeholders if getattr(ph, "ai_generate", False)]
            if not ai_placeholders:
                continue

            facts: Dict[str, Any] = {}
            for placeholder in slide.placeholders:
                source = getattr(placeholder, "source", None)
                if not source:
                    continue
                value = self._get_nested_raw(raw, source)
                if value is None:
                    continue
                facts[source] = value

            tasks.append(
                GenerationRetrievalTask(
                    slide_key=slide.slide_key,
                    slide_title=getattr(slide, "title", slide.slide_key),
                    ai_tokens=[placeholder.token for placeholder in ai_placeholders],
                    ai_instructions=[
                        (getattr(placeholder, "ai_instruction", "") or "").strip()
                        for placeholder in ai_placeholders
                        if (getattr(placeholder, "ai_instruction", "") or "").strip()
                    ],
                    facts=facts,
                    focus_options=focus,
                    context_policy=(getattr(slide, "context_policy", "auto") or "auto"),
                )
            )
        return tasks

    @staticmethod
    def _join_context_by_slide(
        context_by_slide: Dict[str, str],
        slide_order: Sequence[str],
    ) -> str:
        if not context_by_slide:
            return ""
        chunks: List[str] = []
        for slide_key in slide_order:
            context = (context_by_slide.get(slide_key) or "").strip()
            if not context:
                continue
            chunks.append(f"## slide={slide_key}\n{context}")
        return "\n\n".join(chunks).strip()

    def _retrieve_by_generation_tasks(
        self,
        template_id: str,
        tasks: List[GenerationRetrievalTask],
        session_id: Optional[str] = None,
    ) -> RetrievalResult:
        if not tasks:
            return RetrievalResult(rag_used=False)

        context_by_slide: Dict[str, str] = {}
        retrieval_trace: List[Dict[str, Any]] = []
        all_hits: List[Dict[str, Any]] = []
        per_slide_stats: List[Dict[str, Any]] = []
        query_entries: List[Dict[str, Any]] = []
        max_chars_per_slide = max(200, int(config.settings.rag_max_context_chars_per_slide))

        for task in tasks:
            query_text = QueryBuilderV2.build_generation_query(template_id=template_id, task=task)
            query_entries.append(
                {
                    "slide_key": task.slide_key,
                    "query_text": query_text,
                    "query_chars": len(query_text),
                }
            )
            result = self.query(
                query_text=query_text,
                template_id=template_id,
                scene="generate",
            )
            slide_context = self._build_context(result.hits, max_chars=max_chars_per_slide)
            if slide_context:
                context_by_slide[task.slide_key] = slide_context

            query_preview = query_text[:220]
            per_slide_stats.append(
                {
                    "slide_key": task.slide_key,
                    "query_chars": len(query_text),
                    **(result.retrieval_stats or {}),
                }
            )

            for trace_item in result.retrieval_trace:
                retrieval_trace.append(
                    {
                        **trace_item,
                        "slide_key": task.slide_key,
                        "query_preview": query_preview,
                    }
                )

            for hit in result.hits:
                all_hits.append({**hit, "slide_key": task.slide_key})

        ordered_keys = [task.slide_key for task in tasks]
        context = self._join_context_by_slide(context_by_slide, ordered_keys)
        retrieval_stats = {
            "mode": "generation_tasks",
            "slides_total": len(tasks),
            "slides_with_hits": len(context_by_slide),
            "per_slide": per_slide_stats,
        }
        self._dump_query_markdown(
            session_id=session_id,
            scene="generate",
            template_id=template_id,
            query_entries=query_entries,
        )
        retrieval_result = RetrievalResult(
            rag_used=bool(context_by_slide),
            context=context,
            retrieval_trace=retrieval_trace,
            hits=all_hits,
            context_by_slide=context_by_slide,
            retrieval_stats=retrieval_stats,
        )
        self._dump_retrieval_markdown(
            session_id=session_id,
            scene="generate",
            template_id=template_id,
            retrieval_result=retrieval_result,
        )
        return retrieval_result

    def retrieve_for_generation(
        self,
        tenant_input: TenantInput,
        input_id: str,
        template_id: str,
        focus_options: Optional[List[str]] = None,
        use_rag: bool = True,
        template_descriptor: Optional[Any] = None,
        session_id: Optional[str] = None,
    ) -> RetrievalResult:
        if not use_rag:
            return RetrievalResult(rag_used=False, context="", retrieval_trace=[], hits=[])

        if template_descriptor is not None:
            tasks = self._build_generation_tasks(
                tenant_input=tenant_input,
                template_descriptor=template_descriptor,
                focus_options=focus_options,
            )
            return self._retrieve_by_generation_tasks(
                template_id=template_id,
                tasks=tasks,
                session_id=session_id,
            )

        focus_text = ",".join(focus_options or [])
        query_text = (
            f"scene=generate\n"
            f"template={template_id}\n"
            f"focus={focus_text}\n"
            f"tenant_snapshot={self._tenant_snapshot(tenant_input)}"
        )
        self._dump_query_markdown(
            session_id=session_id,
            scene="generate",
            template_id=template_id,
            query_entries=[
                {
                    "query_text": query_text,
                    "query_chars": len(query_text),
                }
            ],
        )
        retrieval_result = self.query(
            query_text=query_text,
            template_id=template_id,
            scene="generate",
        )
        self._dump_retrieval_markdown(
            session_id=session_id,
            scene="generate",
            template_id=template_id,
            retrieval_result=retrieval_result,
        )
        return retrieval_result

    def retrieve_for_rewrite(
        self,
        tenant_input: TenantInput,
        input_id: str,
        template_id: str,
        slide_key: str,
        user_prompt: str,
        current_slide_content: Optional[Dict[str, Any]] = None,
        use_rag: bool = True,
        template_descriptor: Optional[Any] = None,
        target_tokens: Optional[List[str]] = None,
        structured_slide_data: Optional[Dict[str, Any]] = None,
        session_id: Optional[str] = None,
    ) -> RetrievalResult:
        if not use_rag:
            return RetrievalResult(rag_used=False, context="", retrieval_trace=[], hits=[])

        if template_descriptor is not None:
            target_slide = next(
                (slide for slide in getattr(template_descriptor, "slides", []) if slide.slide_key == slide_key),
                None,
            )
            if target_slide is None:
                return RetrievalResult(rag_used=False)

            ai_placeholders = [ph for ph in target_slide.placeholders if getattr(ph, "ai_generate", False)]
            if target_tokens:
                requested = set(target_tokens)
                ai_placeholders = [ph for ph in ai_placeholders if ph.token in requested]

            ai_tokens = [ph.token for ph in ai_placeholders]
            ai_instructions = [
                (getattr(ph, "ai_instruction", "") or "").strip()
                for ph in ai_placeholders
                if (getattr(ph, "ai_instruction", "") or "").strip()
            ]

            facts = dict(structured_slide_data or {})
            if not facts:
                raw = tenant_input.raw or {}
                for placeholder in target_slide.placeholders:
                    source = getattr(placeholder, "source", None)
                    if not source:
                        continue
                    value = self._get_nested_raw(raw, source)
                    if value is None:
                        continue
                    facts[source] = value

            task = RewriteRetrievalTask(
                slide_key=slide_key,
                ai_tokens=ai_tokens,
                ai_instructions=ai_instructions,
                user_prompt=user_prompt,
                facts=facts,
                current_slide_content=dict(current_slide_content or {}),
            )
            query_text = QueryBuilderV2.build_rewrite_query(
                template_id=template_id,
                task=task,
            )
            self._dump_query_markdown(
                session_id=session_id,
                scene="rewrite",
                template_id=template_id,
                query_entries=[
                    {
                        "slide_key": slide_key,
                        "query_text": query_text,
                        "query_chars": len(query_text),
                    }
                ],
            )
            result = self.query(
                query_text=query_text,
                template_id=template_id,
                scene="rewrite",
            )
            context = self._build_context(
                result.hits,
                max_chars=max(200, int(config.settings.rag_max_context_chars_per_slide)),
            )
            context_by_slide = {slide_key: context} if context else {}
            retrieval_stats = {
                "mode": "rewrite_task",
                "slide_key": slide_key,
                "query_chars": len(query_text),
                **(result.retrieval_stats or {}),
            }
            retrieval_trace = [
                {
                    **trace_item,
                    "slide_key": slide_key,
                    "query_preview": query_text[:220],
                }
                for trace_item in result.retrieval_trace
            ]
            hits = [{**hit, "slide_key": slide_key} for hit in result.hits]
            retrieval_result = RetrievalResult(
                rag_used=bool(context),
                context=context,
                retrieval_trace=retrieval_trace,
                hits=hits,
                context_by_slide=context_by_slide,
                retrieval_stats=retrieval_stats,
            )
            self._dump_retrieval_markdown(
                session_id=session_id,
                scene="rewrite",
                template_id=template_id,
                retrieval_result=retrieval_result,
            )
            return retrieval_result

        current_text = json.dumps(current_slide_content or {}, ensure_ascii=False)
        query_text = (
            f"scene=rewrite\n"
            f"template={template_id}\n"
            f"slide_key={slide_key}\n"
            f"user_prompt={user_prompt}\n"
            f"current_slide={current_text[:1000]}\n"
            f"tenant_snapshot={self._tenant_snapshot(tenant_input, max_chars=1200)}"
        )
        self._dump_query_markdown(
            session_id=session_id,
            scene="rewrite",
            template_id=template_id,
            query_entries=[
                {
                    "slide_key": slide_key,
                    "query_text": query_text,
                    "query_chars": len(query_text),
                }
            ],
        )
        retrieval_result = self.query(
            query_text=query_text,
            template_id=template_id,
            scene="rewrite",
        )
        self._dump_retrieval_markdown(
            session_id=session_id,
            scene="rewrite",
            template_id=template_id,
            retrieval_result=retrieval_result,
        )
        return retrieval_result


_shared_rag_service: Optional[RAGService] = None
_shared_rag_lock = threading.Lock()


def get_rag_service() -> RAGService:
    global _shared_rag_service
    if _shared_rag_service is not None:
        return _shared_rag_service
    with _shared_rag_lock:
        if _shared_rag_service is None:
            _shared_rag_service = RAGService()
    return _shared_rag_service

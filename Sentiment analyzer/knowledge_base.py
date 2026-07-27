import hashlib
import logging
import os
import re

import chromadb
import pandas as pd
from chromadb.config import Settings
from chromadb.utils import embedding_functions
from sqlalchemy import create_engine

from config import (
    APP_MODE,
    DATA_DIR,
    EMBEDDING_BATCH_SIZE,
    EMBED_MODEL,
    ENABLE_VECTOR_INDEX,
    EXTERNAL_DATA_MODE,
    INTERNAL_DATA_MODE,
    OLLAMA_HOST,
    VECTOR_INDEX_DIR,
)
from data_pipeline import DemoCsvProvider, InternalApiProvider
from feedback_repository import FeedbackRepository
from timeframe_filters import filter_by_timeframe, rolling_month_count


logger = logging.getLogger(__name__)


def _dataframe_fingerprint(data_frame):
    normalized = data_frame.fillna("").astype(str)
    digest = hashlib.sha256()
    digest.update("\x1f".join(map(str, normalized.columns)).encode("utf-8"))
    digest.update(pd.util.hash_pandas_object(normalized, index=True).values.tobytes())
    return digest.hexdigest()


def _collection_prefix(model_name):
    normalized = re.sub(r"[^a-z0-9]+", "_", str(model_name or "").lower()).strip("_")
    return f"cx_{normalized or 'embedding'}"


def _unique_text(values, *, limit, max_length):
    output = []
    seen = set()
    for value in values:
        text = str(value or "").strip()
        if not text or text.lower() in {"nan", "none", "null", "-"}:
            continue
        normalized = text.casefold()
        if normalized in seen:
            continue
        seen.add(normalized)
        output.append(text[:max_length])
        if len(output) >= limit:
            break
    return output


def _build_feedback_embedding_records(data_frame):
    group_columns = [
        column
        for column in ("Layanan", "Customer Journey Hint")
        if column in data_frame.columns
    ]
    if not group_columns:
        group_columns = [data_frame.columns[0]]

    ids, documents, metadata = [], [], []
    for position, (group_key, group) in enumerate(
        data_frame.fillna("").groupby(group_columns, dropna=False, sort=True)
    ):
        group_values = group_key if isinstance(group_key, tuple) else (group_key,)
        group_metadata = {
            column: str(value or "")
            for column, value in zip(group_columns, group_values)
        }
        ratings = pd.to_numeric(group.get("Rating"), errors="coerce").dropna()
        rating_counts = (
            ratings.round(2).value_counts().sort_index().to_dict()
            if not ratings.empty
            else {}
        )
        comments = _unique_text(
            group.get("Komentar", pd.Series(dtype=str)),
            limit=5,
            max_length=240,
        )
        reasons = _unique_text(
            group.get("Representative Why", pd.Series(dtype=str)),
            limit=3,
            max_length=200,
        )
        stakeholder_types = _unique_text(
            group.get("Tipe Stakeholder", pd.Series(dtype=str)),
            limit=8,
            max_length=80,
        )
        instructor_types = _unique_text(
            group.get("Tipe Instruktur", pd.Series(dtype=str)),
            limit=8,
            max_length=80,
        )
        fields = [
            *(f"{column}: {value}" for column, value in group_metadata.items() if value),
            f"Jumlah respons: {len(group)}",
            f"Rata-rata rating: {ratings.mean():.2f}" if not ratings.empty else "",
            "Distribusi rating: "
            + ", ".join(f"{rating}: {count}" for rating, count in rating_counts.items())
            if rating_counts
            else "",
            "Tipe stakeholder: " + ", ".join(stakeholder_types) if stakeholder_types else "",
            "Tipe instruktur: " + ", ".join(instructor_types) if instructor_types else "",
            "Komentar representatif: " + " ; ".join(comments) if comments else "",
            "Alasan representatif: " + " ; ".join(reasons) if reasons else "",
        ]
        ids.append(str(position))
        documents.append(" | ".join(field for field in fields if field))
        metadata.append(
            {
                **group_metadata,
                "record_count": int(len(group)),
            }
        )
    return ids, documents, metadata


class KnowledgeBase:
    def __init__(self, db_uri):
        os.makedirs(DATA_DIR, exist_ok=True)
        self.engine = create_engine(db_uri)
        self.repository = FeedbackRepository(self.engine)
        self.repository.ensure_schema()
        self.app_mode = APP_MODE
        self.internal_data_mode = INTERNAL_DATA_MODE
        self.external_data_mode = EXTERNAL_DATA_MODE
        self.enable_vector_index = ENABLE_VECTOR_INDEX
        self.provider = self._build_provider()
        self.chroma = None
        self.embed_fn = None
        self.collection = None
        if self.enable_vector_index:
            os.makedirs(VECTOR_INDEX_DIR, exist_ok=True)
            self.chroma = chromadb.PersistentClient(
                path=VECTOR_INDEX_DIR,
                settings=Settings(anonymized_telemetry=False),
            )
            self.embed_fn = embedding_functions.OllamaEmbeddingFunction(
                url=f"{OLLAMA_HOST}/api/embeddings",
                model_name=EMBED_MODEL,
                timeout=300,
            )
        self.df = None
        self.refresh_data()

    def _build_provider(self):
        if self.internal_data_mode == "api":
            return InternalApiProvider()
        return DemoCsvProvider()

    def activate_internal_api_provider(self):
        self.internal_data_mode = "api"
        self.provider = InternalApiProvider()

    def _load_cached_dataframe(self):
        try:
            cached_df = self.repository.load_feedback_dataframe()
            if cached_df is not None and not cached_df.empty:
                logger.warning("Using cached internal data from SQLite.")
                return cached_df
        except Exception as exc:
            logger.warning("No cached internal dataset available: %s", exc)
        return None

    def _rebuild_vector_store(self):
        if not self.enable_vector_index or self.chroma is None:
            return True
        if self.df is None or self.df.empty:
            return False

        ids, documents, metadata = _build_feedback_embedding_records(self.df)

        if not ids:
            return False

        prefix = _collection_prefix(EMBED_MODEL)
        fingerprint = _dataframe_fingerprint(self.df)
        collection_name = f"{prefix}_{fingerprint[:12]}"
        previous_collection = self.collection
        candidate_collection = self.chroma.get_or_create_collection(
            name=collection_name,
            embedding_function=self.embed_fn,
            metadata={
                "model": EMBED_MODEL,
                "source_fingerprint": fingerprint,
            },
        )
        if candidate_collection.count() == len(ids):
            self.collection = candidate_collection
            logger.info(
                "Reusing feedback vector index %s with %s records.",
                collection_name,
                len(ids),
            )
            return True

        existing_ids = candidate_collection.get().get("ids", [])
        if existing_ids:
            candidate_collection.delete(ids=existing_ids)

        try:
            logger.info(
                "Indexing %s feedback records with %s in batches of %s.",
                len(ids),
                EMBED_MODEL,
                EMBEDDING_BATCH_SIZE,
            )
            for start in range(0, len(ids), EMBEDDING_BATCH_SIZE):
                end = start + EMBEDDING_BATCH_SIZE
                candidate_collection.add(
                    documents=documents[start:end],
                    metadatas=metadata[start:end],
                    ids=ids[start:end],
                )
            if candidate_collection.count() != len(ids):
                raise RuntimeError("vector index count does not match feedback record count")
        except Exception:
            self.collection = previous_collection
            try:
                self.chroma.delete_collection(collection_name)
            except Exception:
                logger.warning("Failed to remove incomplete vector collection %s.", collection_name)
            raise

        self.collection = candidate_collection
        for existing in self.chroma.list_collections():
            existing_name = getattr(existing, "name", str(existing))
            if existing_name.startswith(f"{prefix}_") and existing_name != collection_name:
                self.chroma.delete_collection(existing_name)
        return True

    def refresh_data(self):
        try:
            latest_df = self.provider.load_feedback_data()
            self.repository.save_feedback_dataframe(
                latest_df,
                source_name=self.provider.source_name,
                metadata=getattr(self.provider, "last_ingestion_metadata", {}),
            )
            self.df = latest_df
        except Exception as exc:
            logger.error(
                "Failed to load internal data from %s: %s",
                self.provider.source_name,
                exc,
            )
            self.df = self._load_cached_dataframe()
            if self.df is None or self.df.empty:
                return False

        try:
            return self._rebuild_vector_store()
        except Exception as exc:
            logger.error("Failed to rebuild vector store: %s", exc)
            return False

    @staticmethod
    def _rerank_documents(query_text, documents, metadatas=None, distances=None, limit=25, enabled=None):
        if enabled is None:
            enabled = os.getenv("EVIDENCE_QUALITY_ENABLED", "true").strip().lower() in {"1", "true", "yes", "on"}
        if not enabled:
            return list(documents[:limit])
        try:
            from evidence_quality import rerank
            candidates = []
            for index, document in enumerate(documents):
                metadata = metadatas[index] if metadatas and index < len(metadatas) else {}
                distance = distances[index] if distances and index < len(distances) else 1.0
                try:
                    semantic_score = 1.0 / (1.0 + max(0.0, float(distance)))
                except (TypeError, ValueError):
                    semantic_score = 0.0
                normalized_meta = dict(metadata or {})
                normalized_meta.setdefault("question", normalized_meta.get("Pertanyaan", normalized_meta.get("response_name", "")))
                normalized_meta.setdefault("service", normalized_meta.get("Layanan", ""))
                normalized_meta.setdefault("trainer", normalized_meta.get("PIC Layanan", ""))
                candidates.append({"id": str(index), "text": str(document or ""), "semantic_score": semantic_score, "metadata": normalized_meta})
            retrieval_intent = {
                "goal": "find feedback themes, representative comments, service dimensions, and segment evidence",
                "preferred_datasets": ["ClassReport", "ReferenceClassReport"],
                "exclude": ["FinanceInvoice", "ProjectStandards"],
                "preferred_terms": [query_text],
            }
            ranked = rerank("feedback", query_text, candidates, limit=limit, retrieval_intent=retrieval_intent)
            ordered = [str(item.get("text") or "") for item in ranked if isinstance(item, dict)]
            return ordered or list(documents[:limit])
        except Exception:
            return list(documents[:limit])

    def query(self, timeframe, context_keywords=""):
        if (
            not self.enable_vector_index
            or self.collection is None
            or self.collection.count() <= 0
            or (timeframe and rolling_month_count(timeframe) is None)
        ):
            return self._fallback_query(timeframe)

        base_query = (
            "Feedback umum, keluhan, apresiasi, tema layanan, dan masalah operasional. "
            f"{context_keywords}"
        ).strip()
        query_text = (
            "Instruct: Temukan bukti feedback internal yang paling relevan untuk analisis layanan.\n"
            f"Query: {base_query}"
        )
        query_payload = {
            "query_texts": [query_text],
            "n_results": min(40, max(1, self.collection.count())),
            "include": ["documents", "metadatas", "distances"],
        }
        try:
            result = self.collection.query(**query_payload)
            documents = result.get("documents", [[]])
            if documents and documents[0]:
                metadatas = (result.get("metadatas") or [[]])[0]
                distances = (result.get("distances") or [[]])[0]
                ranked = self._rerank_documents(query_text, documents[0], metadatas, distances, limit=25)
                return "\n---\n".join(ranked)
        except Exception as exc:
            logger.error("Query error: %s", exc)

        return self._fallback_query(timeframe)

    def _fallback_query(self, timeframe):
        filtered_df = self.df if self.df is not None else pd.DataFrame()
        if timeframe and not filtered_df.empty:
            filtered_df = filter_by_timeframe(filtered_df, timeframe)
        if filtered_df.empty:
            return "Tidak ada data feedback internal untuk periode ini."
        documents = [
            " | ".join(f"{column}: {value}" for column, value in row.items())
            for _, row in filtered_df.head(25).iterrows()
        ]
        return "\n---\n".join(documents)

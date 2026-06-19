import logging
import os

import chromadb
import pandas as pd
from chromadb.config import Settings
from chromadb.utils import embedding_functions
from sqlalchemy import create_engine

from config import (
    APP_MODE,
    DATA_DIR,
    EMBED_MODEL,
    ENABLE_VECTOR_INDEX,
    EXTERNAL_DATA_MODE,
    INTERNAL_DATA_MODE,
    OLLAMA_HOST,
)
from data_pipeline import DemoCsvProvider, InternalApiProvider
from feedback_repository import FeedbackRepository
from timeframe_filters import filter_by_timeframe, rolling_month_count


logger = logging.getLogger(__name__)


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
            self.chroma = chromadb.Client(Settings(anonymized_telemetry=False))
            self.embed_fn = embedding_functions.OllamaEmbeddingFunction(
                url=f"{OLLAMA_HOST}/api/embeddings",
                model_name=EMBED_MODEL,
            )
            self.collection = self.chroma.get_or_create_collection(
                name="cx_holistic_db",
                embedding_function=self.embed_fn,
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
        if not self.enable_vector_index or self.collection is None:
            return True
        if self.df is None or self.df.empty:
            return False

        existing_ids = self.collection.get().get("ids", [])
        if existing_ids:
            self.collection.delete(existing_ids)

        ids, documents, metadata = [], [], []
        for index, row in self.df.iterrows():
            text_representation = " | ".join(
                f"{column}: {value}" for column, value in row.items()
            )
            ids.append(str(index))
            documents.append(text_representation)
            metadata.append(row.astype(str).to_dict())

        if not ids:
            return False

        logger.info(
            "Sending %s feedback records to Ollama embeddings (%s).",
            len(ids),
            OLLAMA_HOST,
        )
        self.collection.add(documents=documents, metadatas=metadata, ids=ids)
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
        if not self.enable_vector_index or self.collection is None:
            filtered_df = self.df if self.df is not None else pd.DataFrame()
            if timeframe and not filtered_df.empty:
                filtered_df = filter_by_timeframe(filtered_df, timeframe)
            if filtered_df.empty:
                return "Tidak ada data feedback internal untuk periode ini."
            limited_rows = filtered_df.head(25)
            documents = []
            for _, row in limited_rows.iterrows():
                documents.append(
                    " | ".join(f"{column}: {value}" for column, value in row.items())
                )
            return "\n---\n".join(documents)

        query_text = (
            f"General feedback, complaints, praise, and operational issues. "
            f"{context_keywords}"
        ).strip()
        query_payload = {
            "query_texts": [query_text],
            "n_results": min(40, max(1, self.collection.count())),
            "include": ["documents", "metadatas", "distances"],
        }
        if timeframe:
            if rolling_month_count(timeframe) is None:
                query_payload["where"] = {"Rentang Waktu": timeframe}

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

        return "Tidak ada data feedback internal untuk periode ini."

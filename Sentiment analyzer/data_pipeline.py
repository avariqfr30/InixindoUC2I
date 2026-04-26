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
    CSV_PATH,
    DATA_DIR,
    DB_URI,
    EMBED_MODEL,
    ENABLE_VECTOR_INDEX,
    EXTERNAL_DATA_MODE,
    INTERNAL_DATA_MODE,
    OLLAMA_HOST,
)
from internal_api import InternalApiClient
from internal_connector import load_internal_connector

logger = logging.getLogger(__name__)

CANONICAL_INTERNAL_COLUMNS = (
    "Record ID",
    "Sumber Feedback",
    "Kanal Feedback",
    "Tanggal Feedback",
    "Tipe Stakeholder",
    "Layanan",
    "Lokasi",
    "Tipe Instruktur",
    "Rentang Waktu",
    "Rating",
    "Komentar",
    "Customer Journey Hint",
)

COLUMN_ALIASES = {
    "Record ID": ("record_id", "id", "feedback_id", "ticket_id", "case_id", "uuid", "kode", "no_tiket", "nomor_tiket"),
    "Sumber Feedback": ("sumber feedback", "source", "feedback_source", "origin", "source_name", "sumber", "asal_data", "source_type"),
    "Kanal Feedback": ("kanal feedback", "channel", "feedback_channel", "touchpoint", "platform", "kanal", "media", "channel_name"),
    "Tanggal Feedback": ("tanggal feedback", "feedback_date", "created_at", "submitted_at", "date", "tanggal", "tanggal_submit", "tgl_feedback", "created_date"),
    "Tipe Stakeholder": ("tipe stakeholder", "stakeholder_type", "stakeholder", "customer_segment", "customer_type", "segment", "segmen", "tipe_pelanggan", "jenis_pelanggan", "kategori_peserta", "instansi_type"),
    "Layanan": ("layanan", "service", "service_name", "product", "offering", "service_type", "nama_layanan", "program", "course", "training_name", "kelas", "judul_pelatihan"),
    "Lokasi": ("lokasi", "location", "training_location", "city", "kota", "venue_location", "tempat", "cabang", "venue"),
    "Tipe Instruktur": ("tipe instruktur", "instructor_type", "trainer_type", "coach_type", "internal_ol", "internal_or_ol", "trainer_origin", "jenis_instruktur", "tipe_trainer", "pengajar_type"),
    "Rentang Waktu": ("rentang waktu", "timeframe", "periode", "period", "reporting_period", "bulan", "semester", "tahun", "periode_laporan"),
    "Rating": ("rating", "score", "csat", "sentiment_score", "nilai", "skor", "bintang", "kepuasan", "satisfaction_score"),
    "Komentar": ("komentar", "comment", "feedback", "feedback_text", "review", "notes", "complaint_text", "customer_comment", "ulasan", "saran", "kritik", "pesan", "testimoni", "isi_feedback"),
    "Customer Journey Hint": ("customer_journey_hint", "journey_hint", "journey_stage", "customer_journey_stage", "touchpoint_stage", "tahap_journey", "fase_layanan"),
}

DATE_COLUMN_ALIASES = (
    "tanggal feedback", "tanggal", "date", "created_at", "submitted_at", "feedback_date",
)

class InternalDataProvider:
    source_name = "internal"

    @staticmethod
    def _normalize_token(value):
        return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")

    @classmethod
    def _column_variants(cls, column_name):
        raw_name = str(column_name).strip()
        normalized = cls._normalize_token(raw_name)
        variants = {normalized}

        for part in re.split(r"[.\[\]/\\]+", raw_name):
            token = cls._normalize_token(part)
            if token:
                variants.add(token)

        normalized_parts = [part for part in normalized.split("_") if part]
        if normalized_parts:
            variants.add(normalized_parts[-1])
            if len(normalized_parts) >= 2:
                variants.add("_".join(normalized_parts[-2:]))

        return variants

    @classmethod
    def _find_matching_column(cls, columns, aliases):
        alias_tokens = [cls._normalize_token(alias) for alias in aliases if cls._normalize_token(alias)]
        best_match = None
        best_score = -1

        for column_name in columns:
            variants = cls._column_variants(column_name)
            score = 0
            for alias_token in alias_tokens:
                if alias_token in variants:
                    score += 100
                elif any(
                    variant.endswith(f"_{alias_token}")
                    or variant.startswith(f"{alias_token}_")
                    or (len(alias_token) >= 4 and alias_token in variant)
                    for variant in variants
                ):
                    score += 60
            if score > best_score:
                best_match = column_name
                best_score = score

        return best_match if best_score > 0 else None

    @classmethod
    def normalize_dataframe(cls, raw_df):
        if raw_df is None:
            return pd.DataFrame(columns=list(CANONICAL_INTERNAL_COLUMNS))

        dataframe = raw_df.copy()
        dataframe.columns = [str(column).strip() for column in dataframe.columns]

        rename_map = {}
        for canonical_name, aliases in COLUMN_ALIASES.items():
            if canonical_name in dataframe.columns:
                continue
            matched_column = cls._find_matching_column(dataframe.columns, aliases)
            if matched_column:
                rename_map[matched_column] = canonical_name

        dataframe = dataframe.rename(columns=rename_map)

        feedback_dates = pd.Series(dtype="datetime64[ns]")
        for alias in DATE_COLUMN_ALIASES:
            matched_column = cls._find_matching_column(dataframe.columns, (alias,))
            if matched_column:
                feedback_dates = pd.to_datetime(
                    dataframe[matched_column],
                    errors="coerce",
                )
                if feedback_dates.notna().any():
                    dataframe["Tanggal Feedback"] = feedback_dates.dt.strftime("%Y-%m-%d")
                    break

        if "Rentang Waktu" not in dataframe.columns:
            for alias in DATE_COLUMN_ALIASES:
                matched_column = cls._find_matching_column(dataframe.columns, (alias,))
                if not matched_column:
                    continue
                parsed_dates = pd.to_datetime(
                    dataframe[matched_column],
                    errors="coerce",
                )
                if parsed_dates.notna().any():
                    dataframe["Rentang Waktu"] = parsed_dates.dt.to_period("M").astype(str)
                    break

        for column_name in CANONICAL_INTERNAL_COLUMNS:
            if column_name not in dataframe.columns:
                dataframe[column_name] = pd.NA

        if dataframe["Record ID"].isna().all():
            dataframe["Record ID"] = [f"FB-{index + 1:05d}" for index in range(len(dataframe))]

        dataframe["Rating"] = pd.to_numeric(dataframe["Rating"], errors="coerce")
        for column_name in CANONICAL_INTERNAL_COLUMNS:
            dataframe[column_name] = dataframe[column_name].fillna("").astype(str).str.strip()

        return dataframe

    def load_feedback_data(self):
        raise NotImplementedError

class DemoCsvProvider(InternalDataProvider):
    source_name = "demo_csv"

    def load_feedback_data(self):
        if not os.path.exists(CSV_PATH):
            raise FileNotFoundError(f"CSV file not found: {CSV_PATH}")
        raw_df = pd.read_csv(CSV_PATH)
        return self.normalize_dataframe(raw_df)

class InternalApiProvider(InternalDataProvider):
    source_name = "company_api"

    def __init__(self):
        self.client = InternalApiClient()
        self.dataset_name = "feedback"
        self.connector = load_internal_connector()

    def _reload_connector(self):
        self.connector = load_internal_connector()

    def _load_via_connector(self):
        if not self.connector or not self.connector.enabled:
            return None

        interpreted = self.client.interpret_payload(self.connector.to_endpoint_spec())
        raw_df = pd.DataFrame(interpreted["records"])
        if raw_df.empty:
            raise ValueError(
                f"Internal connector '{self.connector.name}' returned no records."
            )

        mapped_df = self.connector.apply_field_map(raw_df)
        normalized_df = self.normalize_dataframe(mapped_df)
        return normalized_df

    def load_dataset(self, dataset_name, extra_params=None):
        raw_df = pd.DataFrame(
            self.client.fetch_records(dataset_name, extra_params=extra_params)
        )
        if raw_df.empty:
            raise ValueError(
                f"Internal API returned no records for endpoint '{dataset_name}'."
            )
        return self.normalize_dataframe(raw_df)

    def load_feedback_data(self):
        self._reload_connector()
        connector_df = self._load_via_connector()
        if connector_df is not None:
            return connector_df
        return self.load_dataset(self.dataset_name)

class KnowledgeBase:
    def __init__(self, db_uri):
        os.makedirs(DATA_DIR, exist_ok=True)
        self.engine = create_engine(db_uri)
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

    def _load_cached_dataframe(self):
        try:
            cached_df = pd.read_sql("SELECT * FROM feedback", self.engine)
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
            latest_df.to_sql("feedback", self.engine, index=False, if_exists="replace")
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

    def query(self, timeframe, context_keywords=""):
        if not self.enable_vector_index or self.collection is None:
            filtered_df = self.df if self.df is not None else pd.DataFrame()
            if timeframe and not filtered_df.empty:
                filtered_df = filtered_df[filtered_df["Rentang Waktu"] == timeframe]
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
        query_payload = {"query_texts": [query_text], "n_results": 25}
        if timeframe:
            query_payload["where"] = {"Rentang Waktu": timeframe}

        try:
            result = self.collection.query(**query_payload)
            documents = result.get("documents", [[]])
            if documents and documents[0]:
                return "\n---\n".join(documents[0])
        except Exception as exc:
            logger.error("Query error: %s", exc)

        return "Tidak ada data feedback internal untuk periode ini."


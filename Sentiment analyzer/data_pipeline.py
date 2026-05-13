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

APIDOG_TIMELESS_TIMEFRAME = "Semua Data APIDog (tanggal tidak tersedia)"
APIDOG_UNKNOWN_DATE = "Tanggal tidak tersedia"
APIDOG_CLASS_REPORT_CHANNEL = "Evaluasi Kelas Internal"

CLASS_REPORT_METADATA_COLUMNS = (
    "Raw Response Count",
    "Rating Response Count",
    "Text Response Count",
    "Rating Distribution",
    "Representative Why",
)

CLASS_REPORT_LABEL_OVERRIDES = {
    "KESESUAIAN MATERIAL BAHAN AJAR": "Kesesuaian materi bahan ajar",
    "KUALITAS PENYAMPAIAN INSTRUKTUR": "Kualitas penyampaian instruktur",
    "PENGUASAAN MATERI INSTRUKTUR": "Penguasaan materi instruktur",
    "KEMAMPUAN INSTRUKTUR MENJAWAB PERTANYAAN": "Kemampuan instruktur menjawab pertanyaan",
    "FASILITAS RUANG KELAS": "Fasilitas ruang kelas",
    "KUALITAS KONSUMSI": "Kualitas konsumsi",
    "KOMENTAR INSTRUKTUR": "Komentar instruktur",
    "SARAN": "Saran peserta",
}

CLASS_REPORT_JOURNEY_RULES = (
    (("materi", "bahan ajar", "kurikulum", "modul"), "Materi dan kurikulum", "Pelaksanaan Layanan"),
    (("instruktur", "trainer", "pengajar", "penyampaian"), "Kinerja instruktur", "Pelaksanaan Layanan"),
    (("fasilitas", "ruang", "kelas", "lab", "lokasi"), "Fasilitas pelatihan", "Persiapan dan Kesiapan Delivery"),
    (("konsumsi", "makan", "snack", "coffee"), "Hospitality pelatihan", "Persiapan dan Kesiapan Delivery"),
    (("pendaftaran", "administrasi", "sertifikat"), "Administrasi pelatihan", "Tindak Lanjut dan Outcome"),
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

    def _apply_connector_auth_mode(self):
        auth_mode = str(getattr(self.connector, "auth_mode", "") or "").strip().lower()
        if auth_mode == "basic_env":
            self.client.auth_mode = "basic"
        elif auth_mode == "none":
            self.client.auth_mode = "none"
        elif auth_mode == "bearer_env":
            self.client.auth_mode = "api_key"
            self.client.auth_prefix = self.client.auth_prefix or "Bearer"

    @staticmethod
    def _looks_like_class_report(dataframe):
        columns = {str(column).strip() for column in dataframe.columns}
        return {"response_id", "response_name", "response_answer"}.issubset(columns)

    @staticmethod
    def _looks_like_reference_class_report(dataframe):
        columns = {str(column).strip() for column in dataframe.columns}
        return (
            {"response_id", "response_name", "response_type"}.issubset(columns)
            and "response_answer" not in columns
        )

    @staticmethod
    def _clean_class_report_label(value):
        raw_label = re.sub(r"\s+", " ", str(value or "")).strip(" :-")
        if not raw_label:
            return "Evaluasi kelas"
        override = CLASS_REPORT_LABEL_OVERRIDES.get(raw_label.upper())
        if override:
            return override
        if raw_label.isupper():
            return raw_label.lower().capitalize()
        return raw_label[:1].upper() + raw_label[1:]

    @classmethod
    def _class_report_question_lookup(cls, dataframe):
        lookup = {}
        if dataframe is None or dataframe.empty:
            return lookup
        for _, row in dataframe.iterrows():
            response_id = str(row.get("response_id") or "").strip()
            if not response_id:
                continue
            lookup[response_id] = {
                "label": cls._clean_class_report_label(row.get("response_name")),
                "type": str(row.get("response_type") or "").strip(),
                "parent_id": str(row.get("response_parent_id") or "").strip(),
            }
        return lookup

    @staticmethod
    def _class_report_semantics(clean_label):
        lowered = str(clean_label or "").lower()
        for keywords, service_label, journey_hint in CLASS_REPORT_JOURNEY_RULES:
            if any(keyword in lowered for keyword in keywords):
                return service_label, journey_hint
        return clean_label or "Evaluasi kelas", "Pelaksanaan Layanan"

    @staticmethod
    def _is_rating_response(row):
        response_type = str(row.get("response_type") or "").strip().lower()
        if response_type.startswith("rating"):
            return True
        answer = str(row.get("response_answer") or "").strip()
        return bool(answer) and pd.notna(pd.to_numeric(answer, errors="coerce"))

    @staticmethod
    def _is_text_response(row):
        response_type = str(row.get("response_type") or "").strip().lower()
        return response_type == "text"

    @staticmethod
    def _format_rating_value(value):
        if pd.isna(value):
            return ""
        rounded = round(float(value), 2)
        if rounded.is_integer():
            return str(int(rounded))
        return str(rounded).rstrip("0").rstrip(".")

    @classmethod
    def _class_report_question_label(cls, row, reference_lookup=None):
        reference_lookup = reference_lookup or {}
        response_id = str(row.get("response_id") or "").strip()
        reference = reference_lookup.get(response_id, {})
        return cls._clean_class_report_label(reference.get("label") or row.get("response_name"))

    @classmethod
    def _class_report_parent_label(cls, parent_id, fallback_row, reference_lookup=None):
        reference_lookup = reference_lookup or {}
        reference = reference_lookup.get(str(parent_id or "").strip(), {})
        return cls._clean_class_report_label(reference.get("label") or fallback_row.get("response_name"))

    @classmethod
    def _build_class_report_row(cls, endpoint_name, row_index, question, rating_value, explanation_texts):
        service_label, journey_hint = cls._class_report_semantics(question)
        average_text = f"Rata-rata rating {question}: {round(float(rating_value), 2)} dari 5" if pd.notna(rating_value) else question
        explanations = cls._dedupe_texts(explanation_texts, limit=5)
        why_text = f"Mengapa: {'; '.join(explanations)}" if explanations else "Mengapa: belum ada komentar teks yang terhubung ke rating ini."
        return {
            "Record ID": f"{endpoint_name}-{row_index + 1:05d}",
            "Sumber Feedback": endpoint_name,
            "Kanal Feedback": APIDOG_CLASS_REPORT_CHANNEL,
            "Tanggal Feedback": APIDOG_UNKNOWN_DATE,
            "Tipe Stakeholder": "Peserta Kelas",
            "Layanan": service_label,
            "Lokasi": "",
            "Tipe Instruktur": "",
            "Rentang Waktu": APIDOG_TIMELESS_TIMEFRAME,
            "Rating": cls._format_rating_value(rating_value),
            "Komentar": f"{average_text}. {why_text}",
            "Customer Journey Hint": journey_hint,
            "Raw Response Count": "",
            "Rating Response Count": "",
            "Text Response Count": "",
            "Rating Distribution": "",
            "Representative Why": "; ".join(explanations),
        }

    @staticmethod
    def _dedupe_texts(values, limit=5):
        deduped = []
        seen = set()
        for value in values:
            cleaned = re.sub(r"\s+", " ", str(value or "").strip(" .:-"))
            key = cleaned.lower()
            if not cleaned or key in seen:
                continue
            seen.add(key)
            deduped.append(cleaned)
            if len(deduped) >= limit:
                break
        return deduped

    @classmethod
    def _normalize_class_report_dataframe(cls, dataframe, endpoint_name, reference_lookup=None):
        reference_lookup = reference_lookup or {}
        rating_groups = {}
        text_by_parent = {}
        orphan_text_rows = []
        for index, row in dataframe.iterrows():
            response_id = str(row.get("response_id") or "").strip()
            answer = str(row.get("response_answer") or "").strip()
            if not response_id and not answer:
                continue

            if cls._is_rating_response(row):
                rating = pd.to_numeric(answer, errors="coerce")
                if pd.isna(rating):
                    continue
                group = rating_groups.setdefault(
                    response_id,
                    {
                        "first_index": index,
                        "question": cls._class_report_question_label(row, reference_lookup),
                        "ratings": [],
                    },
                )
                group["ratings"].append(float(rating))
                continue

            if cls._is_text_response(row):
                parent_id = str(row.get("response_parent_id") or "").strip()
                if parent_id:
                    text_by_parent.setdefault(parent_id, []).append(answer)
                else:
                    orphan_text_rows.append((index, cls._class_report_question_label(row, reference_lookup), answer))

        rows = []
        for response_id, group in sorted(rating_groups.items(), key=lambda item: item[1]["first_index"]):
            ratings = group["ratings"]
            average_rating = sum(ratings) / len(ratings) if ratings else float("nan")
            explanations = text_by_parent.get(response_id, [])
            distribution = {}
            for value in ratings:
                label = cls._format_rating_value(value)
                distribution[label] = distribution.get(label, 0) + 1
            row_payload = cls._build_class_report_row(
                endpoint_name,
                len(rows),
                group["question"],
                average_rating,
                explanations,
            )
            row_payload.update(
                {
                    "Raw Response Count": str(len(ratings) + len(explanations)),
                    "Rating Response Count": str(len(ratings)),
                    "Text Response Count": str(len(explanations)),
                    "Rating Distribution": "; ".join(
                        f"{key}: {distribution[key]}" for key in sorted(distribution, key=lambda item: float(item))
                    ),
                }
            )
            rows.append(
                row_payload
            )

        for _, question, answer in orphan_text_rows:
            service_label, journey_hint = cls._class_report_semantics(question)
            rows.append(
                {
                    "Record ID": f"{endpoint_name}-{len(rows) + 1:05d}",
                    "Sumber Feedback": endpoint_name,
                    "Kanal Feedback": APIDOG_CLASS_REPORT_CHANNEL,
                    "Tanggal Feedback": APIDOG_UNKNOWN_DATE,
                    "Tipe Stakeholder": "Peserta Kelas",
                    "Layanan": service_label,
                    "Lokasi": "",
                    "Tipe Instruktur": "",
                    "Rentang Waktu": APIDOG_TIMELESS_TIMEFRAME,
                    "Rating": "",
                    "Komentar": f"{question}: {answer}".strip(": "),
                    "Customer Journey Hint": journey_hint,
                    "Raw Response Count": "1",
                    "Rating Response Count": "0",
                    "Text Response Count": "1",
                    "Rating Distribution": "",
                    "Representative Why": answer,
                }
            )

        output = pd.DataFrame(rows)
        for column_name in CLASS_REPORT_METADATA_COLUMNS:
            if column_name not in output.columns:
                output[column_name] = ""
            output[column_name] = output[column_name].fillna("").astype(str).str.strip()
        return output

    def _load_via_connector(self):
        if not self.connector or not self.connector.enabled:
            return None

        self._apply_connector_auth_mode()
        endpoint_payloads = []
        reference_lookup = {}
        for endpoint in self.connector.active_endpoints():
            interpreted = self.client.interpret_payload(endpoint.to_endpoint_spec())
            raw_df = pd.DataFrame(interpreted["records"])
            if raw_df.empty:
                logger.warning(
                    "Internal connector endpoint '%s' returned no records.",
                    endpoint.endpoint_name,
                )
                continue

            if self._looks_like_reference_class_report(raw_df):
                reference_lookup.update(self._class_report_question_lookup(raw_df))
                continue
            endpoint_payloads.append((endpoint, raw_df))

        normalized_frames = []
        for endpoint, raw_df in endpoint_payloads:
            if self._looks_like_class_report(raw_df):
                mapped_df = self._normalize_class_report_dataframe(
                    raw_df,
                    endpoint.endpoint_name,
                    reference_lookup=reference_lookup,
                )
            else:
                mapped_df = endpoint.apply_field_map(raw_df)
            normalized_df = self.normalize_dataframe(mapped_df)
            if "Sumber Feedback" in normalized_df.columns:
                empty_source = normalized_df["Sumber Feedback"].astype(str).str.strip() == ""
                normalized_df.loc[empty_source, "Sumber Feedback"] = endpoint.endpoint_name
            normalized_frames.append(normalized_df)

        if not normalized_frames:
            raise ValueError(
                f"Internal connector '{self.connector.name}' returned no records."
            )

        combined_df = pd.concat(normalized_frames, ignore_index=True)
        if "Record ID" in combined_df.columns:
            combined_df = combined_df.drop_duplicates(subset=["Record ID"], keep="last")
        return combined_df.reset_index(drop=True)

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

    def activate_internal_api_provider(self):
        self.internal_data_mode = "api"
        self.provider = InternalApiProvider()

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

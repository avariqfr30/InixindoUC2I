import concurrent.futures
import io
import logging
import os
import re
import textwrap
from datetime import datetime
from urllib.parse import urlparse

import chromadb
import markdown
import matplotlib
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import pandas as pd
import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from chromadb.config import Settings
from chromadb.utils import embedding_functions
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from sqlalchemy import create_engine

from config import (
    APP_MODE,
    CX_SENTIMENT_STRUCTURE,
    CSV_PATH,
    DATA_DIR,
    DEFAULT_COLOR,
    EMBED_MODEL,
    ENABLE_VECTOR_INDEX,
    EXTERNAL_DATA_MODE,
    INTERNAL_API_BASE_URL,
    INTERNAL_API_FEEDBACK_ENDPOINT,
    INTERNAL_API_KEY,
    INTERNAL_API_TIMEOUT_SECONDS,
    INTERNAL_DATA_MODE,
    OLLAMA_HOST,
    OSINT_BASE_QUERIES,
    OSINT_MAX_SIGNALS,
    OSINT_RECENCY,
    OSINT_RESULTS_PER_QUERY,
    OSINT_SEARCH_LANGUAGE,
    OSINT_SEARCH_REGION,
    SERPER_API_KEY,
    WRITER_FIRM_NAME,
)

matplotlib.use("Agg")

logger = logging.getLogger(__name__)

CANONICAL_INTERNAL_COLUMNS = (
    "Record ID",
    "Sumber Feedback",
    "Kanal Feedback",
    "Tanggal Feedback",
    "Tipe Stakeholder",
    "Layanan",
    "Rentang Waktu",
    "Rating",
    "Komentar",
)

COLUMN_ALIASES = {
    "Record ID": (
        "record_id",
        "id",
        "feedback_id",
        "ticket_id",
        "case_id",
    ),
    "Sumber Feedback": (
        "sumber feedback",
        "source",
        "feedback_source",
        "origin",
        "source_name",
    ),
    "Kanal Feedback": (
        "kanal feedback",
        "channel",
        "feedback_channel",
        "touchpoint",
        "platform",
        "kanal",
    ),
    "Tanggal Feedback": (
        "tanggal feedback",
        "feedback_date",
        "created_at",
        "submitted_at",
        "date",
        "tanggal",
    ),
    "Tipe Stakeholder": (
        "tipe stakeholder",
        "stakeholder_type",
        "stakeholder",
        "customer_segment",
        "customer_type",
        "segment",
        "segmen",
    ),
    "Layanan": (
        "layanan",
        "service",
        "service_name",
        "product",
        "offering",
        "service_type",
    ),
    "Rentang Waktu": (
        "rentang waktu",
        "timeframe",
        "periode",
        "period",
        "reporting_period",
    ),
    "Rating": (
        "rating",
        "score",
        "csat",
        "sentiment_score",
        "nilai",
    ),
    "Komentar": (
        "komentar",
        "comment",
        "feedback",
        "feedback_text",
        "review",
        "notes",
        "complaint_text",
        "customer_comment",
    ),
}

DATE_COLUMN_ALIASES = (
    "tanggal feedback",
    "tanggal",
    "date",
    "created_at",
    "submitted_at",
    "feedback_date",
)


def append_field(paragraph, instruction):
    run = paragraph.add_run()

    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")

    field_instruction = OxmlElement("w:instrText")
    field_instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    field_instruction.text = instruction

    field_separator = OxmlElement("w:fldChar")
    field_separator.set(qn("w:fldCharType"), "separate")

    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")

    run._r.extend([field_begin, field_instruction, field_separator, field_end])


class InternalDataProvider:
    source_name = "internal"

    @staticmethod
    def _normalize_token(value):
        return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")

    @classmethod
    def normalize_dataframe(cls, raw_df):
        if raw_df is None:
            return pd.DataFrame(columns=list(CANONICAL_INTERNAL_COLUMNS))

        dataframe = raw_df.copy()
        dataframe.columns = [str(column).strip() for column in dataframe.columns]
        normalized_lookup = {
            cls._normalize_token(column): column for column in dataframe.columns
        }

        rename_map = {}
        for canonical_name, aliases in COLUMN_ALIASES.items():
            if canonical_name in dataframe.columns:
                continue
            for alias in aliases:
                matched_column = normalized_lookup.get(cls._normalize_token(alias))
                if matched_column:
                    rename_map[matched_column] = canonical_name
                    break

        dataframe = dataframe.rename(columns=rename_map)

        feedback_dates = pd.Series(dtype="datetime64[ns]")
        for alias in DATE_COLUMN_ALIASES:
            matched_column = normalized_lookup.get(cls._normalize_token(alias))
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
                matched_column = normalized_lookup.get(cls._normalize_token(alias))
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
        self.base_url = INTERNAL_API_BASE_URL
        self.feedback_endpoint = INTERNAL_API_FEEDBACK_ENDPOINT
        self.timeout_seconds = INTERNAL_API_TIMEOUT_SECONDS
        self.api_key = INTERNAL_API_KEY

    def _headers(self):
        headers = {"Accept": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
            headers["X-API-Key"] = self.api_key
        return headers

    def _extract_records(self, payload):
        if isinstance(payload, list):
            return payload
        if not isinstance(payload, dict):
            return None

        for key in ("items", "data", "results", "records", "feedback"):
            value = payload.get(key)
            records = self._extract_records(value)
            if records is not None:
                return records

        return None

    def load_feedback_data(self):
        if not self.base_url:
            raise RuntimeError("INTERNAL_API_BASE_URL is not configured.")

        endpoint = self.feedback_endpoint
        if not endpoint.startswith("/"):
            endpoint = f"/{endpoint}"

        response = requests.get(
            f"{self.base_url}{endpoint}",
            headers=self._headers(),
            timeout=self.timeout_seconds,
        )
        response.raise_for_status()

        payload = response.json()
        records = self._extract_records(payload)
        if records is None:
            raise ValueError("Unsupported internal API payload format.")

        raw_df = pd.DataFrame(records)
        if raw_df.empty:
            raise ValueError("Internal API returned no feedback records.")

        return self.normalize_dataframe(raw_df)


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


class Researcher:
    SERPER_ENDPOINT = "https://google.serper.dev/search"
    INVALID_API_KEYS = {
        "",
        "YOUR_SERPER_API_KEY",
        "masukkan_api_key_serper_anda_disini",
    }

    @staticmethod
    def _is_enabled():
        return (SERPER_API_KEY or "").strip() not in Researcher.INVALID_API_KEYS

    @staticmethod
    def _search_serper(query, max_results=OSINT_RESULTS_PER_QUERY):
        payload = {
            "q": query,
            "num": max_results,
            "gl": OSINT_SEARCH_REGION,
            "hl": OSINT_SEARCH_LANGUAGE,
        }
        if OSINT_RECENCY:
            payload["tbs"] = OSINT_RECENCY

        headers = {
            "X-API-KEY": SERPER_API_KEY,
            "Content-Type": "application/json",
        }

        response = requests.post(
            Researcher.SERPER_ENDPOINT,
            headers=headers,
            json=payload,
            timeout=10,
        )
        response.raise_for_status()
        return response.json()

    @staticmethod
    def _extract_items(query, payload):
        items = []

        for position, entry in enumerate(payload.get("organic", []), start=1):
            items.append(
                {
                    "query": query,
                    "title": entry.get("title", "Tanpa Judul"),
                    "snippet": entry.get("snippet", "").strip(),
                    "url": entry.get("link", "").strip(),
                    "date": entry.get("date", "").strip(),
                    "source_type": "organic",
                    "position": position,
                }
            )

        for position, entry in enumerate(payload.get("news", []), start=1):
            items.append(
                {
                    "query": query,
                    "title": entry.get("title", "Tanpa Judul"),
                    "snippet": entry.get("snippet", "").strip(),
                    "url": entry.get("link", "").strip(),
                    "date": entry.get("date", "").strip(),
                    "source_type": "news",
                    "position": position,
                }
            )

        return [item for item in items if item["url"]]

    @staticmethod
    def _deduplicate_items(items):
        seen_keys = set()
        unique_items = []

        for item in items:
            key = item["url"] or f"{item['title']}::{item['snippet']}"
            if key in seen_keys:
                continue
            seen_keys.add(key)
            unique_items.append(item)

        return unique_items

    @staticmethod
    def _score_items(items, context_text):
        keywords = {
            token.lower()
            for token in re.findall(r"[A-Za-z]{4,}", context_text)
            if token.lower() not in {"yang", "dengan", "untuk", "pada", "from", "this"}
        }

        for item in items:
            corpus = f"{item['title']} {item['snippet']}".lower()
            coverage_score = sum(1 for keyword in keywords if keyword in corpus)
            freshness_bonus = 1 if item["source_type"] == "news" else 0
            ranking_penalty = (item["position"] - 1) * 0.05
            item["score"] = coverage_score + freshness_bonus - ranking_penalty

        return sorted(items, key=lambda value: value["score"], reverse=True)

    @staticmethod
    def _source_domain(url):
        try:
            netloc = urlparse(url).netloc.lower()
            return netloc.replace("www.", "") if netloc else "unknown"
        except Exception:
            return "unknown"

    @staticmethod
    def _format_osint_brief(items, title):
        if not items:
            return "Tidak ada sinyal OSINT eksternal yang dapat digunakan untuk benchmark periode ini."

        lines = [f"{title}:"]
        for index, item in enumerate(items, start=1):
            date_part = f" | tanggal={item['date']}" if item["date"] else ""
            lines.append(
                (
                    f"{index}. {item['title']} | {item['snippet']} "
                    f"| sumber={Researcher._source_domain(item['url'])}"
                    f"{date_part} | url={item['url']}"
                ).strip()
            )

        return "\n".join(lines)

    @staticmethod
    def _run_query_batch(queries, max_signals=OSINT_MAX_SIGNALS):
        collected = []

        with concurrent.futures.ThreadPoolExecutor(
            max_workers=min(4, max(1, len(queries)))
        ) as pool:
            future_map = {
                pool.submit(Researcher._search_serper, query): query for query in queries
            }
            for future in concurrent.futures.as_completed(future_map):
                query = future_map[future]
                try:
                    payload = future.result()
                    collected.extend(Researcher._extract_items(query, payload))
                except Exception as exc:
                    logger.warning("OSINT query gagal (%s): %s", query, exc)

        deduplicated = Researcher._deduplicate_items(collected)
        ranked = Researcher._score_items(deduplicated, " ".join(queries))
        return ranked[:max_signals]

    @staticmethod
    def get_macro_trends(timeframe, notes=""):
        if not Researcher._is_enabled():
            return "Data OSINT eksternal tidak tersedia (SERPER_API_KEY belum diatur)."

        scope = timeframe or "periode terbaru"
        compact_notes = re.sub(r"\s+", " ", notes).strip()
        contextual_query = (
            f"benchmark sentimen pelanggan pelatihan dan konsultasi IT Indonesia {scope}"
        )
        if compact_notes:
            contextual_query += f" {compact_notes[:140]}"

        queries = [
            f"{query} {scope}" for query in OSINT_BASE_QUERIES
        ] + [contextual_query]

        findings = Researcher._run_query_batch(queries)
        return Researcher._format_osint_brief(
            findings,
            "Sinyal OSINT Makro (Indonesia)",
        )

class FeedbackAnalyticsEngine:
    THEME_LIBRARY = (
        {
            "id": "responsiveness",
            "label": "Respons dan SLA",
            "keywords": ("lambat", "respon", "response", "sla", "timeline", "delay", "mundur", "follow up"),
            "prescription": "Tetapkan SLA respon, dashboard aging, dan owner follow-up per tiket/permintaan.",
        },
        {
            "id": "schedule",
            "label": "Jadwal dan beban sesi",
            "keywords": ("jadwal", "padat", "jeda", "durasi", "sesi", "waktu"),
            "prescription": "Kalibrasi durasi sesi, sediakan jeda terstruktur, dan review desain agenda per layanan.",
        },
        {
            "id": "facility",
            "label": "Fasilitas dan infrastruktur",
            "keywords": ("fasilitas", "lab", "ruang", "wifi", "jaringan", "network", "kelas"),
            "prescription": "Audit kesiapan fasilitas sebelum delivery dan tetapkan checklist operasional harian.",
        },
        {
            "id": "instructor",
            "label": "Kualitas instruktur atau konsultan",
            "keywords": ("instruktur", "trainer", "konsultan", "mentor", "pengajar", "narasumber"),
            "prescription": "Perkuat coaching instruktur, review kompetensi domain, dan standardisasi evaluasi fasilitator.",
        },
        {
            "id": "material",
            "label": "Materi dan relevansi konten",
            "keywords": ("materi", "kurikulum", "modul", "silabus", "relevan", "contoh"),
            "prescription": "Review kurikulum per segmen, tambahkan contoh kontekstual, dan perbarui modul prioritas.",
        },
        {
            "id": "communication",
            "label": "Komunikasi dan koordinasi",
            "keywords": ("komunikasi", "informasi", "koordinasi", "brief", "update"),
            "prescription": "Rapikan alur komunikasi pra-delivery dan pastikan semua stakeholder menerima update status yang sama.",
        },
        {
            "id": "outcome",
            "label": "Dampak hasil layanan",
            "keywords": ("actionable", "implementasi", "hasil", "manfaat", "membantu", "sertifikasi"),
            "prescription": "Pertahankan praktik outcome review dan ubah testimoni hasil menjadi playbook layanan.",
        },
    )

    def __init__(self, dataframe):
        self.full_df = dataframe.copy() if dataframe is not None else pd.DataFrame()
        self.full_df = self.full_df.fillna("")
        if not self.full_df.empty:
            self.full_df["Rating Numeric"] = pd.to_numeric(
                self.full_df["Rating"],
                errors="coerce",
            )
            self.full_df["Sentiment Label"] = self.full_df["Rating Numeric"].apply(
                self._sentiment_label
            )
            self.full_df["Komentar Lower"] = self.full_df["Komentar"].astype(str).str.lower()

    @staticmethod
    def _sentiment_label(value):
        if pd.isna(value):
            return "unknown"
        if value >= 4:
            return "positive"
        if value <= 2:
            return "negative"
        return "neutral"

    @staticmethod
    def _safe_percentage(numerator, denominator):
        if not denominator:
            return 0.0
        return round((numerator / denominator) * 100, 1)

    @staticmethod
    def _truncate_text(text, max_length=180):
        clean_text = re.sub(r"\s+", " ", str(text)).strip()
        if len(clean_text) <= max_length:
            return clean_text
        return f"{clean_text[:max_length - 3]}..."

    @staticmethod
    def _series_counts(series, limit=5):
        filtered = series.fillna("").astype(str).str.strip()
        filtered = filtered[filtered != ""]
        return filtered.value_counts().head(limit)

    def _filter_timeframe(self, timeframe):
        if self.full_df.empty:
            return self.full_df.copy()
        filtered = self.full_df[self.full_df["Rentang Waktu"] == timeframe].copy()
        return filtered

    def _theme_hits(self, dataframe):
        theme_stats = []
        if dataframe.empty:
            return theme_stats

        comment_series = dataframe["Komentar Lower"].astype(str)
        for theme in self.THEME_LIBRARY:
            match_mask = comment_series.apply(
                lambda text: any(keyword in text for keyword in theme["keywords"])
            )
            matched = dataframe[match_mask]
            if matched.empty:
                continue
            positive_hits = int((matched["Sentiment Label"] == "positive").sum())
            negative_hits = int((matched["Sentiment Label"] == "negative").sum())
            neutral_hits = int((matched["Sentiment Label"] == "neutral").sum())
            theme_stats.append(
                {
                    "id": theme["id"],
                    "label": theme["label"],
                    "prescription": theme["prescription"],
                    "total_hits": int(len(matched)),
                    "positive_hits": positive_hits,
                    "negative_hits": negative_hits,
                    "neutral_hits": neutral_hits,
                    "matched_df": matched,
                }
            )

        return sorted(
            theme_stats,
            key=lambda item: (item["negative_hits"], item["total_hits"]),
            reverse=True,
        )

    def _quote_lines(self, dataframe, limit=3):
        if dataframe.empty:
            return ["- Tidak ada kutipan yang cukup untuk periode ini."]

        lines = []
        seen_comments = set()
        for _, row in dataframe.iterrows():
            comment = self._truncate_text(row.get("Komentar", ""))
            if not comment or comment in seen_comments:
                continue
            seen_comments.add(comment)
            stakeholder = row.get("Tipe Stakeholder", "Stakeholder")
            service = row.get("Layanan", "Layanan")
            rating = row.get("Rating", "-")
            lines.append(
                f'- "{comment}" ({stakeholder} | {service} | rating {rating})'
            )
            if len(lines) >= limit:
                break

        return lines or ["- Tidak ada kutipan yang cukup untuk periode ini."]

    def _group_risk(self, dataframe, column_name, limit=3):
        if dataframe.empty:
            return []

        rows = []
        grouped = dataframe.groupby(column_name, dropna=False)
        for label, group in grouped:
            clean_label = str(label).strip() or "Tidak terklasifikasi"
            rating_avg = group["Rating Numeric"].mean()
            negative_ratio = (group["Sentiment Label"] == "negative").mean()
            volume = len(group)
            safe_avg_rating = round(rating_avg, 2) if pd.notna(rating_avg) else 0.0
            risk_score = round((negative_ratio * 70) + ((5 - safe_avg_rating) * 6) + min(volume, 10), 1)
            rows.append(
                {
                    "label": clean_label,
                    "volume": volume,
                    "average_rating": safe_avg_rating,
                    "negative_ratio": round(negative_ratio * 100, 1),
                    "risk_score": risk_score,
                }
            )

        rows.sort(key=lambda item: item["risk_score"], reverse=True)
        return rows[:limit]

    def _governance_summary(self, timeframe_df):
        total_rows = len(timeframe_df)
        if total_rows == 0:
            return {
                "total_rows": 0,
                "completeness_pct": 0.0,
                "source_count": 0,
                "channel_count": 0,
            }

        mandatory_fields = ["Tipe Stakeholder", "Layanan", "Rentang Waktu", "Komentar"]
        completeness_scores = []
        for field in mandatory_fields:
            populated = timeframe_df[field].astype(str).str.strip() != ""
            completeness_scores.append(populated.mean())

        source_count = self._series_counts(timeframe_df["Sumber Feedback"], limit=20).shape[0]
        channel_count = self._series_counts(timeframe_df["Kanal Feedback"], limit=20).shape[0]
        if source_count == 0:
            source_count = 1

        return {
            "total_rows": total_rows,
            "completeness_pct": round(sum(completeness_scores) / len(completeness_scores) * 100, 1),
            "source_count": source_count,
            "channel_count": channel_count,
        }

    @staticmethod
    def _rating_assessment(avg_rating):
        if pd.isna(avg_rating):
            return "belum dapat dinilai secara memadai"
        if avg_rating >= 4.3:
            return "sangat baik dan relatif konsisten"
        if avg_rating >= 3.75:
            return "baik, tetapi masih menyisakan beberapa titik perbaikan"
        if avg_rating >= 3.0:
            return "cukup, namun belum cukup stabil untuk dianggap kuat"
        return "masih lemah dan memerlukan perhatian manajemen segera"

    @staticmethod
    def _negative_share_assessment(negative_share):
        if negative_share >= 35:
            return "cukup tinggi dan berpotensi mengganggu persepsi layanan jika tidak segera ditangani"
        if negative_share >= 20:
            return "perlu diawasi karena dapat berkembang menjadi isu yang lebih luas"
        if negative_share > 0:
            return "masih dalam batas terkendali, namun tetap membutuhkan pemantauan"
        return "belum menunjukkan sinyal keluhan yang berarti"

    @staticmethod
    def _risk_severity(risk_score):
        if risk_score >= 55:
            return "tinggi"
        if risk_score >= 40:
            return "menengah"
        return "terkendali"

    @staticmethod
    def _primary_label(series_counts, fallback):
        return series_counts.index[0] if not series_counts.empty else fallback

    @staticmethod
    def _format_count_summary(series_counts, unit="feedback", limit=3):
        if series_counts.empty:
            return "belum terpetakan"
        return ", ".join(
            f"{label} ({count} {unit})"
            for label, count in series_counts.head(limit).items()
        )

    @staticmethod
    def _escape_table_cell(value):
        return str(value).replace("|", "\\|").replace("\n", " ").strip()

    @classmethod
    def _markdown_table(cls, headers, rows):
        if not rows:
            return ""

        header_line = "| " + " | ".join(cls._escape_table_cell(item) for item in headers) + " |"
        separator_line = "| " + " | ".join("---" for _ in headers) + " |"
        row_lines = [
            "| " + " | ".join(cls._escape_table_cell(cell) for cell in row) + " |"
            for row in rows
        ]
        return "\n".join([header_line, separator_line, *row_lines])

    def _distribution_rows(self, series_counts, total_rows, limit=5):
        rows = []
        for label, count in series_counts.head(limit).items():
            rows.append([label, count, f"{self._safe_percentage(count, total_rows)}%"])
        return rows

    def _extract_osint_signals(self, macro_trends, limit=3):
        signals = []
        for line in str(macro_trends).splitlines():
            cleaned = line.strip()
            if not re.match(r"^\d+\.", cleaned):
                continue

            cleaned = re.sub(r"^\d+\.\s*", "", cleaned)
            parts = [part.strip() for part in cleaned.split(" | ") if part.strip()]
            if not parts:
                continue

            title = parts[0]
            snippet = parts[1] if len(parts) > 1 else ""
            source = "Tidak diketahui"
            date = "-"

            for part in parts[2:]:
                if part.startswith("sumber="):
                    source = part.split("=", maxsplit=1)[1] or source
                elif part.startswith("tanggal="):
                    date = part.split("=", maxsplit=1)[1] or date

            signals.append(
                {
                    "title": title,
                    "snippet": snippet,
                    "source": source,
                    "date": date,
                }
            )
            if len(signals) >= limit:
                break

        return signals

    @staticmethod
    def _theme_owner(theme_id):
        owner_map = {
            "responsiveness": "Customer Service / Account Management",
            "schedule": "Operations / Delivery Management",
            "facility": "Operations / General Affairs",
            "instructor": "Academic Lead / Service Quality",
            "material": "Academic Lead / Product Owner",
            "communication": "Customer Service / Project Coordinator",
            "outcome": "Service Owner / Quality Assurance",
        }
        return owner_map.get(theme_id, "Service Owner")

    @staticmethod
    def _theme_outcome(theme_id):
        outcome_map = {
            "responsiveness": "Waktu respons lebih konsisten dan penutupan isu lebih cepat.",
            "schedule": "Pengalaman delivery lebih tertata dan beban sesi lebih seimbang.",
            "facility": "Gangguan operasional di kelas atau sesi layanan dapat ditekan.",
            "instructor": "Konsistensi kualitas fasilitator meningkat di berbagai layanan.",
            "material": "Materi lebih relevan dengan kebutuhan peserta dan konteks klien.",
            "communication": "Ekspektasi stakeholder lebih selaras sejak pra-delivery hingga pasca-delivery.",
            "outcome": "Nilai manfaat layanan lebih mudah dirasakan dan dibuktikan.",
        }
        return outcome_map.get(theme_id, "Persepsi kualitas layanan membaik secara terukur.")

    def _descriptive_markdown(self, timeframe_df, timeframe, notes):
        governance = self._governance_summary(timeframe_df)
        total_rows = governance["total_rows"]
        if total_rows == 0:
            return (
                "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola\n"
                "Tidak ada feedback internal yang tersedia untuk periode ini.\n"
            )

        avg_rating = timeframe_df["Rating Numeric"].mean()
        positive_count = int((timeframe_df["Sentiment Label"] == "positive").sum())
        neutral_count = int((timeframe_df["Sentiment Label"] == "neutral").sum())
        negative_count = int((timeframe_df["Sentiment Label"] == "negative").sum())

        stakeholder_counts = self._series_counts(timeframe_df["Tipe Stakeholder"])
        service_counts = self._series_counts(timeframe_df["Layanan"])
        source_counts = self._series_counts(timeframe_df["Sumber Feedback"])
        channel_counts = self._series_counts(timeframe_df["Kanal Feedback"])

        top_sources = source_counts.index.tolist() if not source_counts.empty else [
            "Sumber internal terstandar"
        ]
        top_channels = channel_counts.index.tolist() if not channel_counts.empty else [
            "Belum terpetakan"
        ]
        positive_share = self._safe_percentage(positive_count, total_rows)
        neutral_share = self._safe_percentage(neutral_count, total_rows)
        negative_share = self._safe_percentage(negative_count, total_rows)

        cleaned_notes = notes.strip().rstrip(".!?")
        focus_line = (
            f"Fokus tambahan dari pengguna pada periode ini adalah: {cleaned_notes}."
            if notes and notes.strip()
            else "Tidak ada fokus tambahan dari pengguna, sehingga analisis dilakukan terhadap seluruh sinyal yang tersedia."
        )
        governance_note = (
            "Cakupan sumber sudah mulai terpetakan, tetapi pemetaan kanal masih perlu diperkuat."
            if governance["channel_count"] == 0
            else "Pemetaan sumber dan kanal sudah tersedia sehingga jalur asal feedback lebih mudah diaudit."
        )
        descriptive_intro = (
            f"Bagian ini menjelaskan kualitas dasar portofolio feedback yang menjadi fondasi laporan. "
            f"Pada periode {timeframe}, sistem memproses {total_rows} feedback tervalidasi dengan "
            f"rata-rata rating {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5, yang menunjukkan kinerja "
            f"layanan berada pada kategori {self._rating_assessment(avg_rating)}. "
            f"Komposisi sentimen memperlihatkan {positive_share}% sinyal positif, {neutral_share}% sinyal netral, "
            f"dan {negative_share}% sinyal negatif."
        )
        governance_intro = (
            f"Dari sisi tata kelola, kelengkapan field inti mencapai {governance['completeness_pct']}%. "
            f"Data berasal dari {governance['source_count']} sumber feedback dan {governance['channel_count']} kanal yang terpetakan. "
            f"{governance_note} {focus_line}"
        )
        indicator_table = self._markdown_table(
            ["Indikator", "Nilai"],
            [
                ["Periode analisis", timeframe],
                ["Total feedback tervalidasi", f"{total_rows} record"],
                ["Rata-rata rating", f"{round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5"],
                ["Kelengkapan field inti", f"{governance['completeness_pct']}%"],
                ["Jumlah sumber feedback", governance["source_count"]],
                ["Jumlah kanal feedback", governance["channel_count"]],
            ],
        )

        chart_line = (
            "[[CHART: Distribusi Sentimen Feedback | Persentase | "
            f"Positif,{positive_share}; "
            f"Netral,{neutral_share}; "
            f"Negatif,{negative_share}]]"
        )
        sentiment_table = self._markdown_table(
            ["Kategori Sentimen", "Jumlah", "Persentase"],
            [
                ["Positif", positive_count, f"{positive_share}%"],
                ["Netral", neutral_count, f"{neutral_share}%"],
                ["Negatif", negative_count, f"{negative_share}%"],
            ],
        )
        stakeholder_table = self._markdown_table(
            ["Segmen Stakeholder", "Jumlah Feedback", "Persentase"],
            self._distribution_rows(stakeholder_counts, total_rows, limit=5),
        )
        service_table = self._markdown_table(
            ["Layanan", "Jumlah Feedback", "Persentase"],
            self._distribution_rows(service_counts, total_rows, limit=5),
        )
        source_lines = [
            f"- Sumber utama: {', '.join(top_sources[:3])}"
        ]
        source_lines.append(f"- Kanal utama: {', '.join(top_channels[:3])}")
        distribution_paragraph = (
            f"Sebaran volume feedback menunjukkan bahwa konsentrasi terbesar berasal dari segmen "
            f"{self._format_count_summary(stakeholder_counts, limit=3)}. Dari sisi layanan, perhatian pengguna paling banyak "
            f"tercurah pada {self._format_count_summary(service_counts, limit=3)}. Pola ini penting untuk dibaca secara hati-hati, "
            f"karena volume tinggi belum otomatis berarti performa buruk, tetapi menandakan area yang paling banyak terekspos kepada pelanggan."
        )
        source_paragraph = (
            f"Dari sisi asal data, sumber yang paling dominan saat ini adalah {', '.join(top_sources[:3])}. "
            f"Pada saat yang sama, kanal yang tercatat masih didominasi oleh {', '.join(top_channels[:3])}. "
            f"Informasi ini perlu dibaca sebagai indikator awal representativitas data: semakin luas sumber dan kanal, "
            f"semakin kuat dasar analisis untuk pengambilan keputusan lintas fungsi."
        )

        return "\n".join(
            [
                "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola",
                descriptive_intro,
                "",
                governance_intro,
                "",
                indicator_table,
                "",
                "## 1.2 Distribusi Sentimen, Rating, dan Volume",
                (
                    f"Distribusi sentimen menunjukkan bahwa proporsi sentimen negatif sebesar {negative_share}% "
                    f"{self._negative_share_assessment(negative_share)}. Sentimen positif tetap menjadi penopang utama "
                    f"pengalaman pelanggan, tetapi keberadaan sentimen netral yang cukup material mengindikasikan masih ada "
                    f"ruang untuk memperkuat pengalaman agar tidak berhenti pada persepsi 'cukup'."
                ),
                "",
                sentiment_table,
                "",
                chart_line,
                "",
                "## 1.3 Distribusi Stakeholder, Layanan, dan Kanal/Sumber",
                distribution_paragraph,
                "",
                "### Stakeholder dengan volume feedback terbesar",
                stakeholder_table,
                "",
                "### Layanan dengan volume feedback terbesar",
                service_table,
                "",
                "### Cakupan sumber dan kanal",
                source_paragraph,
                "",
                *source_lines,
            ]
        )

    def _diagnostic_markdown(self, timeframe_df):
        if timeframe_df.empty:
            return (
                "## 2.1 Akar Masalah Utama dan Pain Point Dominan\n"
                "Tidak ada feedback internal yang tersedia untuk periode ini.\n"
            )

        theme_hits = self._theme_hits(timeframe_df)
        negative_themes = [theme for theme in theme_hits if theme["negative_hits"] > 0][:3]
        positive_themes = sorted(
            theme_hits,
            key=lambda item: (item["positive_hits"], item["total_hits"]),
            reverse=True,
        )[:3]

        if not negative_themes:
            negative_lines = ["- Belum ada pola keluhan dominan yang menonjol; mayoritas feedback berada pada area stabil."]
        else:
            negative_lines = []
            for theme in negative_themes:
                impacted_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
                impacted_segments = self._series_counts(theme["matched_df"]["Tipe Stakeholder"], limit=2)
                negative_lines.append(
                    f"- {theme['label']}: {theme['negative_hits']} sinyal negatif. "
                    f"Layanan terdampak: {', '.join(impacted_services.index.tolist()) or 'belum terpetakan'}. "
                    f"Segmen terdampak: {', '.join(impacted_segments.index.tolist()) or 'belum terpetakan'}."
                )

        positive_lines = []
        for theme in positive_themes:
            if theme["positive_hits"] <= 0:
                continue
            strongest_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
            positive_lines.append(
                f"- {theme['label']}: {theme['positive_hits']} sinyal positif. "
                f"Paling banyak muncul pada layanan {', '.join(strongest_services.index.tolist()) or 'belum terpetakan'}."
            )
        if not positive_lines:
            positive_lines = ["- Belum ada kekuatan yang cukup konsisten untuk dikonfirmasi pada periode ini."]

        negative_quotes = self._quote_lines(timeframe_df[timeframe_df["Sentiment Label"] == "negative"], limit=3)
        positive_quotes = self._quote_lines(timeframe_df[timeframe_df["Sentiment Label"] == "positive"], limit=2)

        service_risks = self._group_risk(timeframe_df, "Layanan", limit=5)
        process_gap_lines = [
            f"- {item['label']}: rata-rata rating {item['average_rating']}, "
            f"proporsi negatif {item['negative_ratio']}%, volume {item['volume']}."
            for item in service_risks
        ] or ["- Belum ada gap proses yang dapat dipetakan."]
        top_issue = negative_themes[0] if negative_themes else None
        top_strength = next(
            (
                theme
                for theme in positive_themes
                if theme["positive_hits"] > 0
                and (not top_issue or theme["id"] != top_issue["id"])
            ),
            None,
        )
        if not top_strength:
            top_strength = next(
                (theme for theme in positive_themes if theme["positive_hits"] > 0),
                None,
            )

        if top_issue and top_strength and top_issue["id"] == top_strength["id"]:
            strength_context = (
                f"Menariknya, tema {top_strength['label']} muncul sebagai area yang terpolarisasi: "
                "sebagian pelanggan menilai sangat baik, sementara sebagian lain masih mengalami hambatan."
            )
        elif top_strength:
            strength_context = (
                f"Di sisi lain, kekuatan yang paling konsisten terlihat pada {top_strength['label']}."
            )
        else:
            strength_context = (
                "Kekuatan layanan belum muncul secara cukup konsisten untuk dijadikan diferensiasi yang kuat."
            )

        diagnostic_intro = (
            f"Analisis diagnostik bertujuan menjawab mengapa pola feedback pada periode ini muncul. "
            f"{'Tema keluhan paling dominan saat ini adalah ' + top_issue['label'] + ', yang berulang pada beberapa komentar pelanggan.' if top_issue else 'Belum ada tema keluhan yang sangat dominan, sehingga pola masalah masih relatif tersebar.'} "
            f"{strength_context}"
        )
        root_cause_table_rows = []
        for theme in negative_themes:
            impacted_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
            impacted_segments = self._series_counts(theme["matched_df"]["Tipe Stakeholder"], limit=2)
            root_cause_table_rows.append(
                [
                    theme["label"],
                    theme["negative_hits"],
                    ", ".join(impacted_services.index.tolist()) or "Belum terpetakan",
                    ", ".join(impacted_segments.index.tolist()) or "Belum terpetakan",
                ]
            )
        root_cause_table = self._markdown_table(
            ["Tema Prioritas", "Sinyal Negatif", "Layanan Dominan", "Segmen Dominan"],
            root_cause_table_rows,
        )
        strength_table_rows = []
        for theme in positive_themes:
            if theme["positive_hits"] <= 0:
                continue
            strongest_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
            strength_table_rows.append(
                [
                    theme["label"],
                    theme["positive_hits"],
                    ", ".join(strongest_services.index.tolist()) or "Belum terpetakan",
                ]
            )
        strength_table = self._markdown_table(
            ["Kekuatan", "Sinyal Positif", "Layanan Dominan"],
            strength_table_rows,
        )
        service_risk_table = self._markdown_table(
            ["Layanan", "Rata-rata Rating", "Proporsi Negatif", "Volume", "Skor Risiko"],
            [
                [
                    item["label"],
                    item["average_rating"],
                    f"{item['negative_ratio']}%",
                    item["volume"],
                    item["risk_score"],
                ]
                for item in service_risks
            ],
        )

        return "\n".join(
            [
                "## 2.1 Akar Masalah Utama dan Pain Point Dominan",
                diagnostic_intro,
                "",
                (
                    "Pembacaan akar masalah dilakukan dengan melihat pengulangan tema, dampaknya pada layanan, "
                    "dan segmen pelanggan yang paling sering menyinggung isu serupa. Dengan pendekatan ini, "
                    "tim manajemen dapat membedakan antara keluhan yang bersifat insidental dan keluhan yang "
                    "sudah layak dibaca sebagai pola struktural."
                ),
                "",
                root_cause_table,
                "",
                *negative_lines,
                "",
                "## 2.2 Kekuatan yang Konsisten dan Area yang Perlu Dijaga",
                (
                    "Selain keluhan, periode ini juga memperlihatkan area yang secara berulang diapresiasi oleh pelanggan. "
                    "Bagian ini penting karena kekuatan yang konsisten dapat dijadikan acuan untuk standardisasi layanan, "
                    "replikasi praktik baik, dan bahan komunikasi nilai kepada klien."
                ),
                "",
                strength_table,
                "",
                *positive_lines,
                "",
                "## 2.3 Bukti Verbatim, Kesenjangan Proses, dan Segmentasi Masalah",
                (
                    "Bukti verbatim di bawah ini digunakan untuk menjaga agar interpretasi manajerial tetap berpijak pada suara pelanggan. "
                    "Ringkasan kesenjangan proses membantu menerjemahkan komentar individual ke dalam area operasional yang dapat ditindaklanjuti."
                ),
                "",
                "### Kutipan keluhan representatif",
                *negative_quotes,
                "### Kutipan apresiasi representatif",
                *positive_quotes,
                "### Kesenjangan proses yang paling terlihat",
                service_risk_table,
                "",
                *process_gap_lines,
            ]
        )

    def _predictive_markdown(self, timeframe_df, macro_trends):
        if timeframe_df.empty:
            return (
                "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut\n"
                "Tidak ada feedback internal yang tersedia untuk periode ini.\n"
            )

        service_risks = self._group_risk(timeframe_df, "Layanan", limit=5)
        stakeholder_risks = self._group_risk(timeframe_df, "Tipe Stakeholder", limit=5)

        risk_lines = []
        for item in service_risks:
            severity = self._risk_severity(item["risk_score"])
            risk_lines.append(
                f"- {item['label']} berisiko {severity} mengalami penurunan kepuasan lanjutan "
                f"karena proporsi sinyal negatif {item['negative_ratio']}% dengan rata-rata rating {item['average_rating']}."
            )
        if not risk_lines:
            risk_lines = ["- Tidak ada risiko layanan yang cukup kuat untuk diproyeksikan pada periode ini."]

        segment_lines = []
        for item in stakeholder_risks:
            segment_lines.append(
                f"- Segmen {item['label']} perlu dipantau karena volume {item['volume']} feedback "
                f"dengan proporsi negatif {item['negative_ratio']}%."
            )
        if not segment_lines:
            segment_lines = ["- Tidak ada segmen pelanggan yang cukup dominan untuk diproyeksikan."]

        osint_signals = self._extract_osint_signals(macro_trends, limit=4)
        osint_lines = []
        for signal in osint_signals:
            osint_lines.append(
                f"- {signal['title']} ({signal['source']}, {signal['date']}): {signal['snippet']}"
            )
        if not osint_lines:
            osint_lines = [
                "- Tren eksternal belum tersedia; prediksi saat ini sepenuhnya didasarkan pada data internal."
            ]

        top_service_risk = service_risks[0] if service_risks else None
        top_segment_risk = stakeholder_risks[0] if stakeholder_risks else None
        predictive_intro = (
            f"Analisis prediktif membaca risiko yang kemungkinan berkembang apabila pola feedback saat ini berlanjut dalam jangka pendek. "
            f"{'Layanan yang paling layak diprioritaskan untuk pengawasan adalah ' + top_service_risk['label'] + '.' if top_service_risk else 'Belum ada layanan dengan pola risiko yang cukup kuat untuk diprioritaskan.'} "
            f"{'Segmen yang paling perlu dipantau adalah ' + top_segment_risk['label'] + '.' if top_segment_risk else 'Belum ada segmen dengan paparan risiko yang dominan.'}"
        )
        service_risk_table = self._markdown_table(
            ["Layanan", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif", "Volume"],
            [
                [
                    item["label"],
                    self._risk_severity(item["risk_score"]).title(),
                    item["average_rating"],
                    f"{item['negative_ratio']}%",
                    item["volume"],
                ]
                for item in service_risks
            ],
        )
        stakeholder_risk_table = self._markdown_table(
            ["Segmen", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif", "Volume"],
            [
                [
                    item["label"],
                    self._risk_severity(item["risk_score"]).title(),
                    item["average_rating"],
                    f"{item['negative_ratio']}%",
                    item["volume"],
                ]
                for item in stakeholder_risks
            ],
        )
        osint_table = self._markdown_table(
            ["Sinyal Eksternal", "Sumber", "Tanggal"],
            [
                [signal["title"], signal["source"], signal["date"]]
                for signal in osint_signals
            ],
        )

        return "\n".join(
            [
                "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut",
                predictive_intro,
                "",
                (
                    "Prediksi pada dokumen ini tidak dimaksudkan sebagai forecast statistik jangka panjang, "
                    "melainkan sebagai early warning berbasis pola rating, proporsi sentimen negatif, dan konsentrasi volume feedback. "
                    "Dengan pendekatan ini, manajemen dapat lebih cepat memutuskan layanan mana yang perlu ditangani lebih dahulu."
                ),
                "",
                service_risk_table,
                "",
                *risk_lines,
                "",
                "## 3.2 Prediksi Segmen dan Layanan yang Paling Rentan",
                (
                    "Selain layanan, pemantauan juga perlu diarahkan pada segmen pelanggan yang memperlihatkan kombinasi "
                    "antara volume feedback tinggi dan kualitas pengalaman yang menurun. Segmen seperti ini biasanya "
                    "lebih cepat mempengaruhi reputasi, retensi, dan peluang repeat engagement."
                ),
                "",
                stakeholder_risk_table,
                "",
                *segment_lines,
                "",
                "## 3.3 Tren Eksternal yang Berpotensi Memperbesar Risiko",
                (
                    "Sinyal eksternal digunakan sebagai benchmark untuk membaca apakah tantangan yang muncul berasal "
                    "murni dari kondisi internal atau juga diperkuat oleh perubahan ekspektasi pasar. "
                    "Bila tren eksternal bergerak ke arah yang sama dengan keluhan pelanggan internal, maka urgensi intervensi meningkat."
                ),
                "",
                osint_table,
                "",
                *osint_lines[:6],
            ]
        )

    def _prescriptive_markdown(self, timeframe_df):
        if timeframe_df.empty:
            return (
                "## 4.1 Intervensi Prioritas 30 Hari\n"
                "Tidak ada feedback internal yang tersedia untuk periode ini.\n"
            )

        theme_hits = self._theme_hits(timeframe_df)
        prioritized_actions = []
        prioritized_rows = []
        for theme in theme_hits:
            if theme["negative_hits"] <= 0:
                continue
            action_index = len(prioritized_actions) + 1
            prioritized_actions.append(
                f"{action_index}. {theme['label']}: {theme['prescription']}"
            )
            prioritized_rows.append(
                [
                    action_index,
                    theme["label"],
                    theme["prescription"],
                    self._theme_owner(theme["id"]),
                    self._theme_outcome(theme["id"]),
                ]
            )
            if len(prioritized_actions) >= 4:
                break

        if not prioritized_actions:
            prioritized_actions = [
                "1. Pertahankan monitoring mingguan karena belum ada pain point dominan yang membutuhkan intervensi besar."
            ]
            prioritized_rows = [
                [
                    1,
                    "Monitoring berkala",
                    "Pertahankan pemantauan mingguan dan lakukan review tren secara berkala.",
                    "Quality Assurance / CX",
                    "Risiko laten tetap termonitor meskipun belum ada isu dominan.",
                ]
            ]

        governance_actions = [
            "1. Wajibkan field sumber feedback, kanal, stakeholder, layanan, tanggal, dan rating pada setiap record yang masuk.",
            "2. Satukan kontrak data antar sistem supaya analisis lintas sumber tetap konsisten dan dapat diaudit.",
            "3. Tetapkan SLA respon dan eskalasi untuk feedback negatif berprioritas tinggi.",
        ]

        roadmap_actions = [
            "1. Minggu 1: validasi kualitas data, pemetaan owner layanan, dan review pain point dominan.",
            "2. Minggu 2: jalankan quick wins pada layanan berisiko tertinggi serta aktifkan dashboard monitoring.",
            "3. Minggu 3-4: evaluasi dampak perbaikan, tutup feedback loop ke stakeholder, dan siapkan iterasi berikutnya.",
            "[[FLOW: Kumpulkan Feedback Multi-Sumber -> Normalisasi dan Audit Data -> Diagnosa Prioritas -> Jalankan Intervensi -> Evaluasi Dampak]]",
        ]
        action_matrix = self._markdown_table(
            ["Prioritas", "Fokus", "Tindakan", "Owner Utama", "Hasil yang Diharapkan"],
            prioritized_rows,
        )
        roadmap_table = self._markdown_table(
            ["Tahap", "Fokus Kerja", "Output yang Diharapkan"],
            [
                ["Minggu 1", "Validasi kualitas data dan pemetaan owner layanan", "Daftar isu prioritas dan penanggung jawab yang disepakati."],
                ["Minggu 2", "Eksekusi quick wins pada layanan berisiko tertinggi", "Perbaikan cepat berjalan dan dashboard monitoring aktif."],
                ["Minggu 3-4", "Evaluasi dampak, penutupan feedback loop, dan iterasi", "Status dampak awal terdokumentasi dan rencana lanjutan tersusun."],
            ],
        )
        prescriptive_intro = (
            "Bagian preskriptif menerjemahkan temuan sebelumnya ke dalam tindakan yang dapat dibahas dan diputuskan dalam forum internal. "
            "Urutan prioritas disusun berdasarkan intensitas sinyal negatif, potensi dampak ke pengalaman pelanggan, dan kebutuhan koordinasi lintas fungsi."
        )

        return "\n".join(
            [
                "## 4.1 Intervensi Prioritas 30 Hari",
                prescriptive_intro,
                "",
                action_matrix,
                "",
                *prioritized_actions,
                "",
                "## 4.2 Penguatan Tata Kelola Feedback dan Eskalasi",
                (
                    "Selain quick wins layanan, perusahaan juga perlu memperkuat tata kelola feedback agar keputusan perbaikan "
                    "berikutnya tidak selalu dimulai dari data yang parsial. Penguatan tata kelola akan menentukan kualitas "
                    "diagnosis, kecepatan eskalasi, dan akuntabilitas tindak lanjut."
                ),
                "",
                *governance_actions,
                "",
                "## 4.3 Rencana Tindak Lanjut Lintas Fungsi",
                (
                    "Rencana tindak lanjut di bawah ini disusun agar forum internal tidak berhenti pada pembacaan laporan, "
                    "tetapi langsung bergerak ke tahap eksekusi. Timeline dapat disesuaikan, namun disiplin implementasi "
                    "antar fungsi tetap menjadi faktor penentu keberhasilan."
                ),
                "",
                roadmap_table,
                "",
                *roadmap_actions,
            ]
        )

    def build_report_sections(self, timeframe, notes, macro_trends):
        timeframe_df = self._filter_timeframe(timeframe)
        section_map = {
            "cx_chap_1": self._descriptive_markdown(timeframe_df, timeframe, notes),
            "cx_chap_2": self._diagnostic_markdown(timeframe_df),
            "cx_chap_3": self._predictive_markdown(timeframe_df, macro_trends),
            "cx_chap_4": self._prescriptive_markdown(timeframe_df),
        }

        sections = []
        for chapter in CX_SENTIMENT_STRUCTURE:
            sections.append(
                {
                    "id": chapter["id"],
                    "title": chapter["title"],
                    "content": section_map.get(chapter["id"], ""),
                }
            )
        return sections

    def build_executive_snapshot(self, timeframe, notes=""):
        timeframe_df = self._filter_timeframe(timeframe)
        if timeframe_df.empty:
            return (
                "## Ringkasan Eksekutif\n"
                "- Tidak ada data internal yang cukup untuk menyusun snapshot eksekutif.\n"
            )

        total_rows = len(timeframe_df)
        avg_rating = timeframe_df["Rating Numeric"].mean()
        negative_count = int((timeframe_df["Sentiment Label"] == "negative").sum())
        negative_share = self._safe_percentage(
            int((timeframe_df["Sentiment Label"] == "negative").sum()),
            total_rows,
        )
        top_service = self._series_counts(timeframe_df["Layanan"], limit=1)
        top_stakeholder = self._series_counts(timeframe_df["Tipe Stakeholder"], limit=1)
        top_risk = self._group_risk(timeframe_df, "Layanan", limit=1)
        governance = self._governance_summary(timeframe_df)
        theme_hits = self._theme_hits(timeframe_df)
        top_issue = next((theme for theme in theme_hits if theme["negative_hits"] > 0), None)
        focus_text = notes.strip() if notes and notes.strip() else "Tidak ada fokus tambahan dari pengguna."

        risk_statement = (
            f"- Risiko teratas saat ini ada pada layanan {top_risk[0]['label']} "
            f"dengan proporsi sinyal negatif {top_risk[0]['negative_ratio']}%."
            if top_risk
            else "- Belum ada layanan dengan risiko dominan yang teridentifikasi."
        )
        executive_intro = (
            f"Laporan ini merangkum kondisi pengalaman pelanggan untuk periode {timeframe} berdasarkan {total_rows} feedback tervalidasi. "
            f"Secara umum, rata-rata rating berada pada level {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5, "
            f"yang menunjukkan kualitas layanan {self._rating_assessment(avg_rating)}. "
            f"Proporsi sentimen negatif tercatat sebesar {negative_share}% ({negative_count} feedback), sehingga kondisi ini "
            f"{self._negative_share_assessment(negative_share)}."
        )
        meeting_context = (
            f"Untuk kebutuhan rapat internal, perhatian utama sebaiknya diarahkan pada layanan "
            f"{top_risk[0]['label'] if top_risk else self._primary_label(top_service, 'yang memiliki volume feedback terbesar')} "
            f"serta pada isu {top_issue['label'] if top_issue else 'konsistensi kualitas layanan'}. "
            f"Fokus tambahan yang diminta pengguna: {focus_text}"
        )
        snapshot_table = self._markdown_table(
            ["Indikator Kunci", "Nilai"],
            [
                ["Total feedback dianalisis", f"{total_rows} record"],
                ["Rata-rata rating", f"{round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5"],
                ["Proporsi sentimen negatif", f"{negative_share}%"],
                ["Layanan dengan volume terbesar", self._primary_label(top_service, "Belum terpetakan")],
                ["Segmen dengan volume terbesar", self._primary_label(top_stakeholder, "Belum terpetakan")],
                ["Kelengkapan field inti", f"{governance['completeness_pct']}%"],
            ],
        )
        meeting_agenda = [
            (
                f"- Apakah layanan {top_risk[0]['label']} memerlukan intervensi prioritas lintas fungsi pada 30 hari ke depan?"
                if top_risk
                else "- Apakah perusahaan perlu memperluas pengumpulan feedback agar risiko layanan lebih mudah dibaca?"
            ),
            (
                f"- Bagaimana tindak lanjut yang paling tepat untuk tema {top_issue['label']} agar tidak berkembang menjadi keluhan berulang?"
                if top_issue
                else "- Kekuatan layanan mana yang paling layak distandardisasi dan direplikasi?"
            ),
            "- Apakah tata kelola sumber, kanal, dan owner tindak lanjut sudah cukup jelas untuk mendukung evaluasi periodik berikutnya?",
        ]

        return "\n".join(
            [
                "## Ringkasan Eksekutif",
                executive_intro,
                "",
                meeting_context,
                "",
                snapshot_table,
                "",
                "### Agenda Diskusi Prioritas",
                *meeting_agenda,
                "",
                "### Poin Utama untuk Pembacaan Cepat",
                f"- Total feedback yang dianalisis: {total_rows} record.",
                f"- Rata-rata rating periode ini: {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5.",
                f"- Volume layanan terbesar: {top_service.index[0] if not top_service.empty else 'Belum terpetakan'}.",
                f"- Segmen dengan volume terbesar: {top_stakeholder.index[0] if not top_stakeholder.empty else 'Belum terpetakan'}.",
                f"- Proporsi sentimen negatif: {negative_share}%.",
                risk_statement,
                "- Struktur laporan ini disusun untuk mendukung analisis Descriptive, Diagnostic, Predictive, dan Prescriptive secara konsisten.",
            ]
        )


class StyleEngine:
    @staticmethod
    def _configure_text_style(style, font_name, font_size, color, bold=False):
        style.font.name = font_name
        style.font.size = Pt(font_size)
        style.font.color.rgb = RGBColor(*color)
        style.font.bold = bold

    @staticmethod
    def apply_document_styles(doc, theme_color):
        for section in doc.sections:
            section.top_margin = Cm(2.54)
            section.bottom_margin = Cm(2.54)
            section.left_margin = Cm(2.54)
            section.right_margin = Cm(2.54)

            footer = section.footer
            footer_paragraph = (
                footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            )
            footer_paragraph.clear()
            footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            footer_run = footer_paragraph.add_run(
                "STRICTLY CONFIDENTIAL | Inixindo Jogja Executive Report | Page "
            )
            footer_run.font.name = "Calibri"
            footer_run.font.size = Pt(9)
            footer_run.font.color.rgb = RGBColor(128, 128, 128)
            append_field(footer_paragraph, "PAGE")

        normal_style = doc.styles["Normal"]
        StyleEngine._configure_text_style(
            normal_style,
            font_name="Calibri",
            font_size=11,
            color=(33, 37, 41),
        )
        normal_format = normal_style.paragraph_format
        normal_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        normal_format.line_spacing = 1.15
        normal_format.space_after = Pt(8)

        heading_1 = doc.styles["Heading 1"]
        StyleEngine._configure_text_style(
            heading_1,
            font_name="Calibri",
            font_size=16,
            color=theme_color,
            bold=True,
        )
        heading_1.paragraph_format.space_before = Pt(18)
        heading_1.paragraph_format.space_after = Pt(8)

        heading_2 = doc.styles["Heading 2"]
        StyleEngine._configure_text_style(
            heading_2,
            font_name="Calibri",
            font_size=13,
            color=(0, 0, 0),
            bold=True,
        )
        heading_2.paragraph_format.space_before = Pt(14)
        heading_2.paragraph_format.space_after = Pt(6)

        heading_3 = doc.styles["Heading 3"]
        StyleEngine._configure_text_style(
            heading_3,
            font_name="Calibri",
            font_size=12,
            color=(54, 54, 54),
            bold=True,
        )
        heading_3.paragraph_format.space_before = Pt(10)
        heading_3.paragraph_format.space_after = Pt(4)

        for list_style_name in (
            "List Bullet",
            "List Bullet 2",
            "List Bullet 3",
            "List Number",
            "List Number 2",
            "List Number 3",
        ):
            if list_style_name not in doc.styles:
                continue
            list_style = doc.styles[list_style_name]
            StyleEngine._configure_text_style(
                list_style,
                font_name="Calibri",
                font_size=11,
                color=(33, 37, 41),
            )
            list_style.paragraph_format.space_after = Pt(4)


class ChartEngine:
    @staticmethod
    def _get_plt_color(theme_color):
        return tuple(channel / 255 for channel in theme_color)

    @staticmethod
    def create_bar_chart(data_str, theme_color):
        try:
            parts = data_str.split("|")
            if len(parts) == 3:
                title_str, ylabel_str, raw_data = [part.strip() for part in parts]
            else:
                title_str, ylabel_str, raw_data = (
                    "Sentimen Makro",
                    "Persentase",
                    data_str,
                )

            labels, values = [], []
            for pair in raw_data.split(";"):
                if "," not in pair:
                    continue
                label, value = pair.split(",", maxsplit=1)
                cleaned_value = re.sub(r"[^\d.]", "", value)
                if not cleaned_value:
                    continue
                labels.append(label.strip())
                values.append(float(cleaned_value))

            if not labels:
                return None

            fig, axis = plt.subplots(figsize=(7, 4.5))
            axis.bar(
                labels,
                values,
                color=ChartEngine._get_plt_color(theme_color),
                alpha=0.9,
                width=0.5,
            )
            axis.set_title(title_str, fontsize=12, fontweight="bold", pad=20)
            axis.set_ylabel(ylabel_str, fontsize=10)
            axis.spines["top"].set_visible(False)
            axis.spines["right"].set_visible(False)

            image_stream = io.BytesIO()
            plt.savefig(image_stream, format="png", bbox_inches="tight", dpi=150)
            plt.close(fig)
            image_stream.seek(0)
            return image_stream
        except Exception as exc:
            logger.warning("Gagal membuat bar chart: %s", exc)
            return None

    @staticmethod
    def create_flowchart(data_str, theme_color):
        try:
            steps = [
                "\n".join(textwrap.wrap(step.strip(), width=18))
                for step in data_str.split("->")
                if step.strip()
            ]
            if len(steps) < 2:
                return None

            fig, axis = plt.subplots(figsize=(8, 3))
            axis.axis("off")
            x_positions = [index * 2.5 for index in range(len(steps))]

            for index in range(len(steps) - 1):
                axis.annotate(
                    "",
                    xy=(x_positions[index + 1] - 1.0, 0.5),
                    xytext=(x_positions[index] + 1.0, 0.5),
                    arrowprops={"arrowstyle": "-|>", "lw": 1.5},
                )

            for index, step in enumerate(steps):
                box = patches.FancyBboxPatch(
                    (x_positions[index] - 1.0, 0.1),
                    2.0,
                    0.8,
                    boxstyle="round,pad=0.1",
                    fc=ChartEngine._get_plt_color(theme_color),
                    alpha=0.9,
                )
                axis.add_patch(box)
                axis.text(
                    x_positions[index],
                    0.5,
                    step,
                    ha="center",
                    va="center",
                    size=9,
                    color="white",
                    fontweight="bold",
                )

            axis.set_xlim(-1.2, (len(steps) - 1) * 2.5 + 1.2)
            axis.set_ylim(0, 1)

            image_stream = io.BytesIO()
            plt.savefig(
                image_stream,
                format="png",
                bbox_inches="tight",
                dpi=200,
                transparent=True,
            )
            plt.close(fig)
            image_stream.seek(0)
            return image_stream
        except Exception as exc:
            logger.warning("Gagal membuat flowchart: %s", exc)
            return None


class DocumentBuilder:
    LIST_STYLES = {
        "ul": ["List Bullet", "List Bullet 2", "List Bullet 3"],
        "ol": ["List Number", "List Number 2", "List Number 3"],
    }

    @staticmethod
    def _resolve_list_style(doc, list_tag, level):
        style_candidates = DocumentBuilder.LIST_STYLES.get(list_tag, ["List Bullet"])
        style_name = style_candidates[min(level, len(style_candidates) - 1)]
        fallback_style = style_candidates[0]
        return style_name if style_name in doc.styles else fallback_style

    @staticmethod
    def _append_inline_runs(paragraph, node, bold=False, italic=False):
        if isinstance(node, NavigableString):
            text = str(node)
            if not text:
                return
            run = paragraph.add_run(text)
            run.bold = bold
            run.italic = italic
            return

        if not isinstance(node, Tag):
            return

        if node.name == "br":
            paragraph.add_run("\n")
            return

        next_bold = bold or node.name in {"strong", "b"}
        next_italic = italic or node.name in {"em", "i"}

        for child in node.children:
            DocumentBuilder._append_inline_runs(
                paragraph,
                child,
                bold=next_bold,
                italic=next_italic,
            )

    @staticmethod
    def _render_list(doc, list_node, level=0):
        style_name = DocumentBuilder._resolve_list_style(doc, list_node.name, level)

        for list_item in list_node.find_all("li", recursive=False):
            paragraph = doc.add_paragraph(style=style_name)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            for child in list_item.contents:
                if isinstance(child, Tag) and child.name in {"ul", "ol"}:
                    continue
                DocumentBuilder._append_inline_runs(paragraph, child)

            nested_lists = [
                child
                for child in list_item.contents
                if isinstance(child, Tag) and child.name in {"ul", "ol"}
            ]
            for nested in nested_lists:
                DocumentBuilder._render_list(doc, nested, level=level + 1)

    @staticmethod
    def _render_table(doc, table_node):
        rows = table_node.find_all("tr")
        if not rows:
            return

        column_count = max(
            len(row.find_all(["th", "td"], recursive=False))
            for row in rows
        )
        table = doc.add_table(rows=0, cols=column_count)
        table.style = "Table Grid"

        for row_index, row_node in enumerate(rows):
            cells = row_node.find_all(["th", "td"], recursive=False)
            table_row = table.add_row().cells

            for col_index in range(column_count):
                text = ""
                if col_index < len(cells):
                    text = cells[col_index].get_text(" ", strip=True)
                table_row[col_index].text = text

                if row_index == 0 and col_index < len(cells) and cells[col_index].name == "th":
                    for run in table_row[col_index].paragraphs[0].runs:
                        run.bold = True

    @staticmethod
    def parse_html_to_docx(doc, html_content):
        soup = BeautifulSoup(html_content, "html.parser")
        for element in soup.contents:
            if not isinstance(element, Tag):
                continue

            if element.name in {"h1", "h2", "h3"}:
                level = int(element.name[1])
                doc.add_heading(element.get_text(" ", strip=True), level=level)
                continue

            if element.name == "p":
                paragraph = doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for child in element.children:
                    DocumentBuilder._append_inline_runs(paragraph, child)
                continue

            if element.name in {"ul", "ol"}:
                DocumentBuilder._render_list(doc, element, level=0)
                continue

            if element.name == "table":
                DocumentBuilder._render_table(doc, element)

    @staticmethod
    def process_content(doc, raw_text, theme_color=DEFAULT_COLOR):
        clean_lines = []
        for line in raw_text.split("\n"):
            stripped_line = line.strip()

            if stripped_line.startswith("[[CHART:") and stripped_line.endswith("]]"):
                chart_data = stripped_line.replace("[[CHART:", "").replace("]]", "").strip()
                image = ChartEngine.create_bar_chart(chart_data, theme_color)
                if image:
                    doc.add_paragraph().add_run().add_picture(image, width=Inches(5.5))
                continue

            if stripped_line.startswith("[[FLOW:") and stripped_line.endswith("]]"):
                flow_data = stripped_line.replace("[[FLOW:", "").replace("]]", "").strip()
                image = ChartEngine.create_flowchart(flow_data, theme_color)
                if image:
                    doc.add_paragraph().add_run().add_picture(image, width=Inches(6.5))
                continue

            clean_lines.append(line)

        html = markdown.markdown("\n".join(clean_lines), extensions=["tables"])
        DocumentBuilder.parse_html_to_docx(doc, html)

    @staticmethod
    def add_table_of_contents(doc):
        doc.add_heading("DAFTAR ISI", level=1)
        toc_paragraph = doc.add_paragraph()
        append_field(toc_paragraph, 'TOC \\o "1-3" \\h \\z \\u')

        note = doc.add_paragraph(
            "Perbarui field di Microsoft Word agar daftar isi otomatis terisi."
        )
        note.runs[0].italic = True
        note.runs[0].font.size = Pt(9)
        note.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_page_break()

    @staticmethod
    def create_cover(doc, timeframe, theme_color=DEFAULT_COLOR):
        StyleEngine.apply_document_styles(doc, theme_color)

        for _ in range(5):
            doc.add_paragraph()

        confidentiality = doc.add_paragraph("S T R I C T L Y   C O N F I D E N T I A L")
        confidentiality.alignment = WD_ALIGN_PARAGRAPH.CENTER
        confidentiality.runs[0].font.size = Pt(10)
        confidentiality.runs[0].font.color.rgb = RGBColor(128, 128, 128)
        confidentiality.runs[0].font.bold = True

        doc.add_paragraph()

        report_title = doc.add_paragraph("HOLISTIC CUSTOMER EXPERIENCE REPORT")
        report_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        report_title.runs[0].font.name = "Calibri"
        report_title.runs[0].font.size = Pt(20)

        company_name = doc.add_paragraph("INIXINDO JOGJA")
        company_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        company_name.runs[0].font.name = "Calibri"
        company_name.runs[0].font.bold = True
        company_name.runs[0].font.size = Pt(32)
        company_name.runs[0].font.color.rgb = RGBColor(*theme_color)

        doc.add_paragraph()

        period_text = doc.add_paragraph(f"Periode Evaluasi Laporan: {timeframe}")
        period_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
        period_text.runs[0].font.size = Pt(13)

        generated_on = datetime.now().strftime("%d %B %Y")
        generated_text = doc.add_paragraph(f"Tanggal Generasi: {generated_on}")
        generated_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
        generated_text.runs[0].font.size = Pt(11)
        generated_text.runs[0].font.color.rgb = RGBColor(128, 128, 128)

        for _ in range(8):
            doc.add_paragraph()

        prepared_for = doc.add_paragraph(
            f"Prepared for Executive Board by:\n{WRITER_FIRM_NAME}"
        )
        prepared_for.alignment = WD_ALIGN_PARAGRAPH.CENTER
        prepared_for.runs[0].font.bold = True

        doc.add_page_break()
        DocumentBuilder.add_table_of_contents(doc)


class ReportGenerator:
    def __init__(self, kb_instance):
        self.kb = kb_instance
        self.research_pool = concurrent.futures.ThreadPoolExecutor(max_workers=4)

    def run(self, timeframe, notes=""):
        logger.info("Starting feedback intelligence report generation for timeframe: %s", timeframe)

        macro_future = self.research_pool.submit(
            Researcher.get_macro_trends,
            timeframe,
            notes,
        )

        try:
            macro_trends = macro_future.result(timeout=25)
        except Exception:
            macro_trends = "Tidak ada tren eksternal yang berhasil dimuat."

        analytics = FeedbackAnalyticsEngine(self.kb.df)
        executive_snapshot = analytics.build_executive_snapshot(timeframe, notes)
        report_sections = analytics.build_report_sections(timeframe, notes, macro_trends)

        document = Document()
        DocumentBuilder.create_cover(document, timeframe, DEFAULT_COLOR)
        document.add_heading("EXECUTIVE SNAPSHOT", level=1)
        DocumentBuilder.process_content(
            document,
            executive_snapshot,
            DEFAULT_COLOR,
        )
        document.add_page_break()

        for index, section in enumerate(report_sections):
            document.add_heading(section["title"], level=1)
            DocumentBuilder.process_content(
                document,
                section["content"],
                DEFAULT_COLOR,
            )
            if index < len(report_sections) - 1:
                document.add_page_break()

        filename = f"Inixindo_Feedback_Intelligence_Report_{timeframe}".replace(" ", "_")
        return document, filename

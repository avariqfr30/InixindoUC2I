import re
from dataclasses import dataclass
from datetime import datetime, timedelta

import pandas as pd

from config import (
    ADOPTION_READINESS_PILLARS,
    CUSTOMER_JOURNEY_STAGES,
    CX_SENTIMENT_STRUCTURE,
    DEFAULT_SCORE_ENGINE,
    SCORE_ENGINE_PROFILES,
    SENTIMENT_OPTIONS,
)
import config
from data_contract import CANONICAL_INTERNAL_COLUMNS
from report_narratives import ReportNarrativeBuilderMixin
from timeframe_filters import filter_by_timeframe, readable_timeframe_label


def _confidence_tier(n: int) -> str:
    """Return confidence label based on sample size."""
    if n >= 100:
        return "tinggi"
    if n >= 30:
        return "sedang"
    if n >= 10:
        return "rendah"
    return "sangat rendah"


@dataclass(frozen=True)
class PreparedReportAnalysis:
    request_key: tuple
    scoped_dataframe: pd.DataFrame
    analysis_context: dict
    governance_summary: dict
    contradiction_review: dict
    top_service: pd.Series
    top_risk: list
    top_issue: dict | None


class FeedbackAnalyticsEngine(ReportNarrativeBuilderMixin):
    PLACEHOLDER_DIMENSION_PATTERNS = (
        "brand equity",
        "mengapa inixindo",
        "menjadi pilihan",
        "alasan memilih inixindo",
        "reputasi dan alasan memilih inixindo",
    )
    CONSTRUCTIVE_CRITIQUE_PATTERNS = (
        r"\bperlu\b",
        r"\bkurang\b",
        r"\btidak\s+sesuai\b",
        r"\bbelum\b",
        r"\bupdate\b",
        r"\bdiupdate\b",
        r"\bversi\s+lama\b",
        r"\bout\s+of\s+date\b",
        r"\bterlalu\b",
        r"\bbeda\b|\bberbeda\b",
        r"\bmepet\b",
        r"\bmasih\b",
    )
    LOW_DETAIL_FEEDBACK = {
        "",
        "tidak",
        "ya",
        "iya",
        "ada",
        "ok",
        "okay",
        "sesuai",
        "benar",
        "mungkin",
    }
    THEME_LIBRARY = (
        {
            "id": "responsiveness", "label": "Respons dan SLA",
            "keywords": ("lambat", "respon", "response", "sla", "timeline", "delay", "mundur", "follow up"),
            "prescription": "Tetapkan SLA respon, dashboard aging, dan owner follow-up per tiket/permintaan.",
        },
        {
            "id": "schedule", "label": "Jadwal dan beban sesi",
            "keywords": ("jadwal", "padat", "jeda", "durasi", "sesi", "waktu"),
            "prescription": "Kalibrasi durasi sesi, sediakan jeda terstruktur, dan review desain agenda per layanan.",
        },
        {
            "id": "facility", "label": "Fasilitas dan infrastruktur",
            "keywords": ("fasilitas", "lab", "ruang", "wifi", "jaringan", "network", "kelas"),
            "prescription": "Audit kesiapan fasilitas sebelum delivery dan tetapkan checklist operasional harian.",
        },
        {
            "id": "instructor", "label": "Kualitas instruktur atau konsultan",
            "keywords": ("instruktur", "trainer", "konsultan", "mentor", "pengajar", "narasumber"),
            "prescription": "Perkuat coaching instruktur, review kompetensi domain, dan standardisasi evaluasi fasilitator.",
        },
        {
            "id": "material", "label": "Materi dan relevansi konten",
            "keywords": ("materi", "kurikulum", "modul", "silabus", "relevan", "contoh"),
            "prescription": "Review kurikulum per segmen, tambahkan contoh kontekstual, dan perbarui modul prioritas.",
        },
        {
            "id": "communication", "label": "Komunikasi dan koordinasi",
            "keywords": ("komunikasi", "informasi", "koordinasi", "brief", "update"),
            "prescription": "Rapikan alur komunikasi pra-delivery dan pastikan semua stakeholder menerima update status yang sama.",
        },
        {
            "id": "outcome", "label": "Dampak hasil layanan",
            "keywords": ("actionable", "implementasi", "hasil", "manfaat", "membantu", "sertifikasi"),
            "prescription": "Pertahankan praktik outcome review dan ubah testimoni hasil menjadi playbook layanan.",
        },
    )

    def __init__(self, dataframe):
        self._prepared_report_analysis = None
        self._prepared_analysis_dataframe = None
        self._prepared_helper_cache = {}
        self.full_df = dataframe.copy() if dataframe is not None else pd.DataFrame()
        for column_name in CANONICAL_INTERNAL_COLUMNS:
            if column_name not in self.full_df.columns:
                self.full_df[column_name] = ""
        self.full_df = self.full_df.fillna("")
        if not self.full_df.empty:
            rating_numeric = pd.to_numeric(self.full_df["Rating"], errors="coerce")
            if "Rating Numeric" in dataframe.columns:
                explicit_rating_numeric = pd.to_numeric(dataframe["Rating Numeric"], errors="coerce")
                rating_numeric = rating_numeric.fillna(explicit_rating_numeric)
            self.full_df["Rating Numeric"] = rating_numeric.apply(self._normalize_rating_value)
            self.full_df["Rating"] = self.full_df["Rating Numeric"].apply(self._format_rating_for_display)
            self.full_df["Layanan"] = self.full_df["Layanan"].apply(self._reader_safe_dimension_label)
            self.full_df["Komentar"] = self.full_df["Komentar"].apply(self._reader_safe_text_label)
            self.full_df["Sentiment Label"] = self.full_df.apply(
                lambda row: self._sentiment_label(row.get("Rating Numeric"), row.get("Komentar")),
                axis=1,
            )
            self.full_df["Komentar Lower"] = self.full_df["Komentar"].astype(str).str.lower()
            self.full_df["Reportable Analysis Row"] = ~self.full_df["Layanan"].apply(self._is_placeholder_dimension)

    @classmethod
    def from_records(cls, records):
        return cls(pd.DataFrame.from_records(records))

    @classmethod
    def _sentiment_label(cls, value, comment=""):
        if pd.isna(value):
            return "unknown"
        low_detail = cls._is_low_detail_comment(comment)
        has_critique = cls._has_constructive_critique(comment)
        if value >= 4:
            if has_critique and not low_detail:
                return "mixed"
            return "positive"
        if value <= 2:
            if low_detail:
                return "weak_negative"
            return "negative"
        if has_critique and not low_detail:
            return "mixed"
        return "neutral"

    @classmethod
    def _is_placeholder_dimension(cls, value):
        text = re.sub(r"\s+", " ", str(value or "").strip().lower())
        return any(pattern in text for pattern in cls.PLACEHOLDER_DIMENSION_PATTERNS)

    @classmethod
    def _has_constructive_critique(cls, value):
        text = str(value or "").lower()
        return any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in cls.CONSTRUCTIVE_CRITIQUE_PATTERNS)

    @classmethod
    def _is_low_detail_comment(cls, value):
        text = re.sub(r"rata-rata (?:rating|penilaian).*?mengapa:", " ", str(value or ""), flags=re.IGNORECASE)
        text = re.sub(r"belum ada komentar teks yang terhubung ke (?:rating|penilaian) ini", " ", text, flags=re.IGNORECASE)
        text = re.sub(r"[^a-zA-Z0-9\s]+", " ", text).lower()
        tokens = [token for token in text.split() if token]
        if not tokens:
            return True
        if len(tokens) <= 2 and " ".join(tokens) in cls.LOW_DETAIL_FEEDBACK:
            return True
        return len(tokens) <= 1 and tokens[0] in cls.LOW_DETAIL_FEEDBACK

    @classmethod
    def _sentiment_issue_weight(cls, value):
        return {
            "negative": 1.0,
            "mixed": 0.55,
            "weak_negative": 0.2,
        }.get(str(value or ""), 0.0)

    @classmethod
    def _sentiment_summary(cls, dataframe):
        total = len(dataframe) if dataframe is not None else 0
        labels = dataframe["Sentiment Label"] if total and "Sentiment Label" in dataframe.columns else pd.Series(dtype="object")
        counts = {label: int((labels == label).sum()) for label in ("positive", "mixed", "neutral", "negative", "weak_negative")}
        issue_weight = float(labels.apply(cls._sentiment_issue_weight).sum()) if total else 0.0
        counts.update(
            {
                "total": total,
                "issue_weight": issue_weight,
                "positive_share": cls._safe_percentage(counts["positive"], total),
                "mixed_share": cls._safe_percentage(counts["mixed"], total),
                "neutral_share": cls._safe_percentage(counts["neutral"], total),
                "negative_share": cls._safe_percentage(counts["negative"], total),
                "weak_negative_share": cls._safe_percentage(counts["weak_negative"], total),
                "issue_share": round((issue_weight / total) * 100, 1) if total else 0.0,
            }
        )
        return counts

    @staticmethod
    def _normalize_rating_value(value):
        if pd.isna(value):
            return value
        numeric_value = float(value)
        if 5 < numeric_value <= 50:
            return numeric_value / 10
        if 50 < numeric_value <= 100:
            return numeric_value / 20
        return numeric_value

    @staticmethod
    def _format_rating_for_display(value):
        if pd.isna(value):
            return ""
        rounded = round(float(value), 2)
        if rounded.is_integer():
            return str(int(rounded))
        return str(rounded).rstrip("0").rstrip(".")

    @staticmethod
    def _reader_safe_dimension_label(value):
        text = re.sub(r"\s+", " ", str(value or "")).strip(" :-")
        lowered = text.lower()
        if any(token in lowered for token in ("brand equity", "mengapa inixindo", "menjadi pilihan")):
            return "Reputasi dan alasan memilih Inixindo"
        if re.search(r"\bpilih\s+\d+\s+bintang\b|\buntuk mengisi\b", text, flags=re.IGNORECASE):
            return "Evaluasi umum kelas"
        if len(text) > 72 and any(mark in text for mark in ("?", "(", ")")):
            return "Evaluasi umum kelas"
        return text or "Tidak terklasifikasi"

    @classmethod
    def _reader_safe_text_label(cls, value):
        text = re.sub(r"\s+", " ", str(value or "")).strip()
        text = re.sub(
            r"BRAND EQUITY\s*\(Mengapa Inixindo Jogja menjadi pilihan\?\)\s*Pilih\s+\d+\s+Bintang\s+untuk\s+mengisi",
            "Reputasi dan alasan memilih Inixindo",
            text,
            flags=re.IGNORECASE,
        )
        return text

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

    @staticmethod
    def _sum_numeric_column(dataframe, column_name):
        if dataframe is None or column_name not in dataframe.columns:
            return 0
        numeric_values = pd.to_numeric(dataframe[column_name], errors="coerce").fillna(0)
        return int(numeric_values.sum())

    def _raw_response_count(self, dataframe):
        raw_count = self._sum_numeric_column(dataframe, "Raw Response Count")
        return raw_count if raw_count > 0 else len(dataframe)

    def _rating_response_count(self, dataframe):
        rating_count = self._sum_numeric_column(dataframe, "Rating Response Count")
        if rating_count > 0:
            return rating_count
        if dataframe is None or dataframe.empty or "Rating Numeric" not in dataframe.columns:
            return 0
        return int(dataframe["Rating Numeric"].notna().sum())

    def _text_response_count(self, dataframe):
        text_count = self._sum_numeric_column(dataframe, "Text Response Count")
        if text_count > 0:
            return text_count
        if dataframe is None or dataframe.empty or "Komentar" not in dataframe.columns:
            return 0
        return int((dataframe["Komentar"].astype(str).str.strip() != "").sum())

    @staticmethod
    def _dimension_count(dataframe):
        return len(dataframe) if dataframe is not None else 0

    @staticmethod
    def _column_series(dataframe, column_name):
        if dataframe is None or column_name not in dataframe.columns:
            return pd.Series(dtype="object")
        return dataframe[column_name]

    def _series_counts_for_column(self, dataframe, column_name, limit=5):
        return self._series_counts(self._column_series(dataframe, column_name), limit=limit)

    @staticmethod
    def _label_from_options(options, option_id, fallback):
        for item in options:
            if item["id"] == option_id:
                return item["label"]
        return fallback

    @staticmethod
    def _clamp(value, minimum=0.0, maximum=100.0):
        return max(minimum, min(maximum, value))

    def _normalize_sentiment_filter(self, sentiment):
        valid_ids = {item["id"] for item in SENTIMENT_OPTIONS}
        return sentiment if sentiment in valid_ids else "all"

    def _normalize_score_engine(self, score_engine):
        return score_engine if score_engine in SCORE_ENGINE_PROFILES else DEFAULT_SCORE_ENGINE

    def _normalize_segment_filter(self, segment):
        cleaned = str(segment or "").strip()
        if not cleaned or cleaned.lower() == "all":
            return "all"
        available_segments = set(self.full_df["Tipe Stakeholder"].fillna("").astype(str).str.strip().tolist())
        return cleaned if cleaned in available_segments else "all"

    def _score_engine_profile(self, score_engine):
        normalized_engine = self._normalize_score_engine(score_engine)
        return SCORE_ENGINE_PROFILES.get(normalized_engine, SCORE_ENGINE_PROFILES[DEFAULT_SCORE_ENGINE])

    def _analysis_scope_text(self, timeframe, sentiment, segment, score_engine):
        sentiment_label = self._label_from_options(SENTIMENT_OPTIONS, sentiment, "Semua Sentimen")
        profile = self._score_engine_profile(score_engine)
        scope_parts = [f"periode {readable_timeframe_label(timeframe)}", f"perspektif {profile['label']}"]
        if sentiment != "all":
            scope_parts.append(f"filter sentimen {sentiment_label.lower()}")
        if segment != "all":
            scope_parts.append(f"segmen {segment}")
        return ", ".join(scope_parts)

    def _forecast_horizon(self, timeframe):
        normalized = str(timeframe or "").lower()
        if "minggu" in normalized or "weekly" in normalized:
            return "1-2 minggu ke depan"
        if "semester" in normalized or "6 bulan" in normalized:
            return "semester berikutnya"
        if "tahun" in normalized or "year" in normalized:
            return "periode tahun berikutnya"
        if "bulan" in normalized or "monthly" in normalized:
            return "1-2 bulan ke depan"
        return "1-2 periode evaluasi berikutnya"

    @staticmethod
    def _format_month_year(value):
        month_names = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        return f"{month_names[value.month - 1]} {value.year}"

    def _forecast_calendar_reference(self, timeframe):
        now = datetime.now()
        normalized = str(timeframe or "").lower()
        if "minggu" in normalized or "weekly" in normalized:
            start, end = now + timedelta(days=7), now + timedelta(days=14)
            return f"sekitar {self._format_month_year(start)} sampai {self._format_month_year(end)}"
        if "bulan" in normalized or "monthly" in normalized:
            start, end = now + timedelta(days=30), now + timedelta(days=60)
            if start.year == end.year and start.month == end.month:
                return f"sekitar {self._format_month_year(start)}"
            return f"sekitar {self._format_month_year(start)} sampai {self._format_month_year(end)}"
        if "semester" in normalized or "6 bulan" in normalized or "biannual" in normalized:
            future = now + timedelta(days=180)
            return f"sekitar semester berikutnya pada {future.year}"
        if "tahun" in normalized or "yearly" in normalized:
            return f"pada tahun {now.year + 1}"
        future = now + timedelta(days=60)
        return f"sekitar {self._format_month_year(future)}"

    def _chart_pairs(self, series_counts, total_rows=None, limit=5, use_percentage=False):
        pairs = []
        for label, count in series_counts.head(limit).items():
            value = self._safe_percentage(count, total_rows) if use_percentage and total_rows else count
            pairs.append(f"{label},{value}")
        return "; ".join(pairs)

    @staticmethod
    def _format_score(value):
        return round(float(value or 0.0), 1)

    @staticmethod
    def _date_series(dataframe):
        if dataframe is None or dataframe.empty or "Tanggal Feedback" not in dataframe.columns:
            return pd.Series(dtype="datetime64[ns]")
        return pd.to_datetime(dataframe["Tanggal Feedback"], errors="coerce", dayfirst=False)

    def _weekly_forecast_source(self, timeframe_df, sentiment, segment):
        dates = self._date_series(timeframe_df)
        if dates.notna().sum() >= 8 and dates.dt.normalize().nunique() >= 2:
            return timeframe_df.copy(), "filtered"

        source = self.full_df.copy()
        if "Reportable Analysis Row" in source.columns:
            source = source[source["Reportable Analysis Row"]].copy()
        if sentiment != "all":
            source = source[source["Sentiment Label"] == sentiment].copy()
        if segment != "all":
            source = source[source["Tipe Stakeholder"].astype(str).str.strip() == segment].copy()
        return source, "full_history"

    def _score_timeseries_for_weeks(self, source_df, score_engine):
        if source_df.empty:
            return []

        dated = source_df.copy()
        dated["_feedback_date"] = self._date_series(dated)
        dated = dated[dated["_feedback_date"].notna()].copy()
        if dated.empty:
            return []

        anchor_date = dated["_feedback_date"].max().normalize()
        rows = []
        for index in range(4, 0, -1):
            end_date = anchor_date - timedelta(days=(index - 1) * 7)
            start_date = end_date - timedelta(days=6)
            bucket = dated[
                (dated["_feedback_date"] >= start_date)
                & (dated["_feedback_date"] <= end_date + timedelta(days=1))
            ].drop(columns=["_feedback_date"])
            if bucket.empty:
                continue
            metrics = self._score_engine_metrics(bucket, score_engine)
            rows.append(
                {
                    "label": f"H-{index}",
                    "start": start_date.date().isoformat(),
                    "end": end_date.date().isoformat(),
                    "score": self._format_score(metrics.get("current_score")),
                    "volume": len(bucket),
                }
            )
        return rows

    @staticmethod
    def _weekly_pattern_label(delta):
        if delta >= 1.0:
            return "menguat"
        if delta <= -1.0:
            return "melemah"
        return "stabil"

    def _build_30d_forecast(self, timeframe_df, score_engine, score_metrics, sentiment, segment):
        source_df, source_scope = self._weekly_forecast_source(timeframe_df, sentiment, segment)
        historical_rows = self._score_timeseries_for_weeks(source_df, score_engine)
        current_score = self._format_score(score_metrics.get("current_score"))
        projected_score = self._format_score(score_metrics.get("projected_score"))
        model_weekly_delta = (projected_score - current_score) / 4.0

        historical_delta = None
        if len(historical_rows) >= 2:
            step_deltas = [
                historical_rows[index]["score"] - historical_rows[index - 1]["score"]
                for index in range(1, len(historical_rows))
            ]
            historical_delta = sum(step_deltas) / len(step_deltas)
            weekly_delta = (historical_delta * 0.6) + (model_weekly_delta * 0.4)
            method = "rolling_7_day_history"
            confidence = _confidence_tier(sum(row["volume"] for row in historical_rows))
        else:
            weekly_delta = model_weekly_delta
            method = "deterministic_fallback"
            confidence = "rendah" if len(timeframe_df) >= 10 else "sangat rendah"

        if (projected_score - current_score) * weekly_delta < 0:
            weekly_delta = model_weekly_delta

        weekly_rows = []
        running_score = current_score
        for week_number in range(1, 5):
            target_score = self._clamp(running_score + weekly_delta)
            if week_number == 4:
                target_score = projected_score
            delta = self._format_score(target_score - running_score)
            weekly_rows.append(
                {
                    "week": f"Minggu {week_number}",
                    "score": self._format_score(target_score),
                    "delta": delta,
                    "pattern": self._weekly_pattern_label(delta),
                    "reading": (
                        "momentum pengalaman menguat"
                        if delta > 0
                        else "perlu mitigasi agar pengalaman tidak melemah"
                        if delta < 0
                        else "pola pengalaman relatif stabil"
                    ),
                }
            )
            running_score = target_score

        component_rows = []
        for component in score_metrics.get("component_breakdown", []):
            current = self._format_score(component.get("current_score"))
            projected = self._format_score(component.get("projected_score"))
            component_rows.append(
                {
                    "component_id": component.get("component_id"),
                    "label": component.get("label"),
                    "weight_pct": self._format_score(float(component.get("weight", 0.0)) * 100),
                    "current_score": current,
                    "projected_score": projected,
                    "delta": self._format_score(projected - current),
                }
            )

        if score_engine == "experience_index":
            component_rows.append(
                {
                    "component_id": "experience_index",
                    "label": score_metrics.get("label", "Experience Index"),
                    "weight_pct": 100.0,
                    "current_score": current_score,
                    "projected_score": projected_score,
                    "delta": self._format_score(projected_score - current_score),
                }
            )

        return {
            "method": method,
            "confidence": confidence,
            "source_scope": source_scope,
            "historical_rows": historical_rows,
            "weekly_rows": weekly_rows,
            "component_rows": component_rows,
            "score_chart": "; ".join(
                f"{row['label']},{row['current_score']}" for row in component_rows
            ),
            "weekly_chart": "; ".join(
                f"{row['week']},{row['score']}" for row in weekly_rows
            ),
            "source_note": (
                "Pola mingguan memakai histori rolling 7 hari pada data bertanggal."
                if method == "rolling_7_day_history"
                else "Data bertanggal belum cukup rapat; pola mingguan memakai fallback deterministik dari proyeksi early-warning."
            ),
        }

    def _filter_timeframe(self, timeframe):
        return self._filter_view(timeframe)

    def _prepared_analysis_key(self, timeframe, sentiment, segment, score_engine):
        return (
            str(timeframe or ""),
            self._normalize_sentiment_filter(sentiment),
            self._normalize_segment_filter(segment),
            self._normalize_score_engine(score_engine),
        )

    def get_prepared_report_analysis(
        self,
        timeframe,
        sentiment="all",
        segment="all",
        score_engine=DEFAULT_SCORE_ENGINE,
    ):
        expected_key = self._prepared_analysis_key(timeframe, sentiment, segment, score_engine)
        prepared = self._prepared_report_analysis
        return (
            prepared
            if prepared is not None and prepared.request_key == expected_key
            else None
        )

    def resolve_prepared_report_analysis(
        self,
        prepared_analysis,
        timeframe,
        sentiment="all",
        segment="all",
        score_engine=DEFAULT_SCORE_ENGINE,
    ):
        expected_key = self._prepared_analysis_key(timeframe, sentiment, segment, score_engine)
        registered = self._prepared_report_analysis
        return (
            registered
            if prepared_analysis is registered
            and registered is not None
            and registered.request_key == expected_key
            else None
        )

    def prepare_report_analysis(
        self,
        timeframe,
        sentiment="all",
        segment="all",
        score_engine=DEFAULT_SCORE_ENGINE,
    ):
        existing = self.get_prepared_report_analysis(
            timeframe,
            sentiment=sentiment,
            segment=segment,
            score_engine=score_engine,
        )
        if existing is not None:
            return existing

        key = self._prepared_analysis_key(timeframe, sentiment, segment, score_engine)
        scoped_dataframe = self._filter_view(
            timeframe,
            sentiment=sentiment,
            segment=segment,
        )
        self._prepared_analysis_dataframe = scoped_dataframe
        self._prepared_helper_cache = {}
        analysis_context = self._build_analysis_context(
            scoped_dataframe,
            timeframe,
            sentiment,
            segment,
            score_engine,
        )
        governance_summary = self._governance_summary(scoped_dataframe)
        theme_hits = self._theme_hits(scoped_dataframe)
        top_service = self._series_counts(scoped_dataframe["Layanan"], limit=1)
        top_risk = self._group_risk(scoped_dataframe, "Layanan", limit=1)
        top_issue = next(
            (theme for theme in theme_hits if theme["negative_hits"] > 0),
            None,
        )
        from report_agents import FeedbackProposalTeam

        prepared = PreparedReportAnalysis(
            request_key=key,
            scoped_dataframe=scoped_dataframe,
            analysis_context=analysis_context,
            governance_summary=governance_summary,
            contradiction_review=FeedbackProposalTeam._contradiction_review(scoped_dataframe),
            top_service=top_service,
            top_risk=top_risk,
            top_issue=top_issue,
        )
        self._prepared_report_analysis = prepared
        return prepared

    def _filter_view(self, timeframe, sentiment="all", segment="all"):
        if self.full_df.empty: return self.full_df.copy()
        filtered = filter_by_timeframe(self.full_df, timeframe)
        if "Reportable Analysis Row" in filtered.columns:
            filtered = filtered[filtered["Reportable Analysis Row"]].copy()
        normalized_sentiment = self._normalize_sentiment_filter(sentiment)
        normalized_segment = self._normalize_segment_filter(segment)
        if normalized_sentiment != "all": filtered = filtered[filtered["Sentiment Label"] == normalized_sentiment]
        if normalized_segment != "all": filtered = filtered[filtered["Tipe Stakeholder"].astype(str).str.strip() == normalized_segment]
        return filtered

    def _customer_journey_keywords(self):
        theme_lookup = {theme["id"]: theme for theme in self.THEME_LIBRARY}
        keyword_map = {}
        for stage in CUSTOMER_JOURNEY_STAGES:
            stage_keywords = []
            for theme_id in stage["theme_ids"]:
                stage_keywords.extend(theme_lookup.get(theme_id, {}).get("keywords", ()))
            keyword_map[stage["label"]] = tuple(dict.fromkeys(stage_keywords))
        return keyword_map

    def _attach_customer_journey(self, dataframe):
        if dataframe.empty:
            enriched = dataframe.copy()
            enriched["Customer Journey Stage"] = pd.Series(dtype="object")
            return enriched

        keyword_map = self._customer_journey_keywords()
        valid_stage_labels = {stage["label"] for stage in CUSTOMER_JOURNEY_STAGES}
        default_stage = next((stage["label"] for stage in CUSTOMER_JOURNEY_STAGES if stage["id"] == "delivery_experience"), CUSTOMER_JOURNEY_STAGES[0]["label"])

        def classify_stage(text):
            lowered = str(text or "").lower()
            best_stage, best_score = default_stage, 0
            for stage_label, keywords in keyword_map.items():
                score = sum(1 for keyword in keywords if keyword in lowered)
                if score > best_score: best_score, best_stage = score, stage_label
            return best_stage

        enriched = dataframe.copy()
        fallback_stages = enriched["Komentar Lower"].apply(classify_stage)
        if "Customer Journey Hint" in enriched.columns:
            hinted_stage = enriched["Customer Journey Hint"].fillna("").astype(str).str.strip()
            enriched["Customer Journey Stage"] = hinted_stage.where(hinted_stage.isin(valid_stage_labels), fallback_stages)
        else:
            enriched["Customer Journey Stage"] = fallback_stages
        return enriched

    def _customer_journey_rows(self, dataframe):
        if dataframe.empty: return []
        enriched = self._attach_customer_journey(dataframe)
        rows = []
        for stage in CUSTOMER_JOURNEY_STAGES:
            label = stage["label"]
            stage_df = enriched[enriched["Customer Journey Stage"] == label]
            if stage_df.empty: continue

            total = len(stage_df)
            sentiment_summary = self._sentiment_summary(stage_df)
            stage_theme_hits = self._theme_hits(stage_df)
            dominant_theme = next((theme["label"] for theme in stage_theme_hits if theme["negative_hits"] > 0 or theme["positive_hits"] > 0), "Sinyal umum customer journey")

            rows.append({
                "stage_id": stage["id"], "stage_label": label, "description": stage["description"], "volume": total,
                "average_rating": round(stage_df["Rating Numeric"].mean(), 2) if stage_df["Rating Numeric"].notna().any() else 0.0,
                "positive_share": sentiment_summary["positive_share"], "neutral_share": sentiment_summary["neutral_share"],
                "negative_share": sentiment_summary["issue_share"], "dominant_theme": dominant_theme,
            })
        rows.sort(key=lambda item: (item["negative_share"], item["volume"]), reverse=True)
        return rows

    def _score_engine_metrics(self, dataframe, score_engine):
        normalized_engine = self._normalize_score_engine(score_engine)
        if normalized_engine == "experience_index":
            return self._score_engine_metrics_experience_index(dataframe)
        return self._score_engine_metrics_single(dataframe, normalized_engine)

    def _score_engine_metrics_single(self, dataframe, score_engine):
        profile = self._score_engine_profile(score_engine)
        if dataframe.empty:
            return {
                "label": profile["label"],
                "current_score": 0.0,
                "projected_score": 0.0,
                "delta": 0.0,
                "direction": "stabil",
                "theme_rows": [],
                "confidence": "sangat rendah",
                "sample_caveat": "Berdasarkan 0 respons (sampel kecil — interpretasi harus hati-hati)",
            }

        avg_rating = dataframe["Rating Numeric"].mean()
        base_score = ((avg_rating / 5) * 100) if pd.notna(avg_rating) else 0.0
        total_rows = len(dataframe)
        sentiment_summary = self._sentiment_summary(dataframe)
        positive_share = sentiment_summary["positive_share"]
        negative_share = sentiment_summary["issue_share"]

        weighted_balance, weighted_positive_ratio, weighted_negative_ratio, total_weight = 0.0, 0.0, 0.0, 0.0
        theme_rows = []

        for theme in self._theme_hits(dataframe):
            weight = profile["theme_weights"].get(theme["id"], 0.35)
            total_hits = max(theme["total_hits"], 1)
            positive_ratio = theme["positive_hits"] / total_hits
            negative_ratio = theme["negative_hits"] / total_hits
            balance = positive_ratio - negative_ratio
            priority_score = round(((theme["negative_hits"] * 1.8) + theme["total_hits"]) * weight, 1)

            weighted_balance += balance * weight
            weighted_positive_ratio += positive_ratio * weight
            weighted_negative_ratio += negative_ratio * weight
            total_weight += weight

            theme_rows.append({
                "theme_id": theme["id"], "label": theme["label"], "weight": round(weight, 2), "total_hits": theme["total_hits"],
                "positive_hits": theme["positive_hits"], "negative_hits": theme["negative_hits"], "priority_score": priority_score, "prescription": theme["prescription"],
            })

        if total_weight > 0:
            weighted_balance /= total_weight; weighted_positive_ratio /= total_weight; weighted_negative_ratio /= total_weight

        current_score = self._clamp((base_score * config.SCORE_BASE_WEIGHT) + ((50 + (weighted_balance * 50)) * config.SCORE_BALANCE_WEIGHT))
        top_weighted_risk = max(((row["negative_hits"] / max(row["total_hits"], 1)) * row["weight"] for row in theme_rows), default=0.0)
        delta = round((((positive_share - negative_share) / 100) * config.SCORE_POS_FACTOR) - (weighted_negative_ratio * config.SCORE_NEG_FACTOR) + (weighted_positive_ratio * config.SCORE_RISK_PENALTY_SCALE) - (top_weighted_risk * config.SCORE_RISK_PENALTY_MAX), 1)
        if abs(delta) < config.SCORE_DIRECTION_THRESHOLD: delta = 0.0
        projected_score = self._clamp(current_score + delta)
        direction = "naik" if delta > 0 else "turun" if delta < 0 else "stabil"

        theme_rows.sort(key=lambda item: (item["priority_score"], item["negative_hits"]), reverse=True)
        return {
            "label": profile["label"],
            "current_score": round(float(current_score), 1),
            "projected_score": round(float(projected_score), 1),
            "delta": round(float(delta), 1),
            "direction": direction,
            "theme_rows": theme_rows,
            "confidence": _confidence_tier(total_rows),
            "sample_caveat": f"Berdasarkan {total_rows} respons" + (
                " (sampel kecil — interpretasi harus hati-hati)" if total_rows < 30 else ""
            ),
        }

    def _score_engine_metrics_experience_index(self, dataframe):
        profile = self._score_engine_profile("experience_index")
        if dataframe.empty:
            return {
                "label": profile["label"],
                "current_score": 0.0,
                "projected_score": 0.0,
                "delta": 0.0,
                "direction": "stabil",
                "theme_rows": [],
                "component_breakdown": [],
                "experience_lenses": [],
            }

        component_weights = profile.get(
            "component_weights",
            {"learning_score": 0.5, "service_score": 0.3, "facility_score": 0.2},
        )

        component_metrics = {}
        for component_id, weight in component_weights.items():
            normalized_component = self._normalize_score_engine(component_id)
            if normalized_component == "experience_index" or weight <= 0:
                continue
            component_metrics[normalized_component] = {
                "weight": float(weight),
                "metrics": self._score_engine_metrics_single(dataframe, normalized_component),
            }

        if not component_metrics:
            return self._score_engine_metrics_single(dataframe, "experience_index")

        total_component_weight = sum(item["weight"] for item in component_metrics.values()) or 1.0
        current_score = sum(
            item["metrics"]["current_score"] * item["weight"]
            for item in component_metrics.values()
        ) / total_component_weight
        projected_score = sum(
            item["metrics"]["projected_score"] * item["weight"]
            for item in component_metrics.values()
        ) / total_component_weight

        delta = round(projected_score - current_score, 1)
        if abs(delta) < config.SCORE_DIRECTION_THRESHOLD:
            delta = 0.0
        projected_score = self._clamp(current_score + delta)
        direction = "naik" if delta > 0 else "turun" if delta < 0 else "stabil"

        rolled_theme_rows = {}
        for component_id, component_payload in component_metrics.items():
            weight = component_payload["weight"]
            for row in component_payload["metrics"]["theme_rows"]:
                theme_id = row["theme_id"]
                aggregate = rolled_theme_rows.setdefault(
                    theme_id,
                    {
                        "theme_id": theme_id,
                        "label": row["label"],
                        "weight": profile["theme_weights"].get(theme_id, 0.0),
                        "total_hits": 0.0,
                        "positive_hits": 0.0,
                        "negative_hits": 0.0,
                        "priority_score": 0.0,
                        "prescription": row["prescription"],
                    },
                )
                aggregate["total_hits"] += row["total_hits"] * weight
                aggregate["positive_hits"] += row["positive_hits"] * weight
                aggregate["negative_hits"] += row["negative_hits"] * weight
                aggregate["priority_score"] += row["priority_score"] * weight

        theme_rows = []
        for item in rolled_theme_rows.values():
            effective_weight = item["weight"] if item["weight"] > 0 else 0.01
            item["priority_score"] = round(item["priority_score"] * effective_weight, 1)
            item["total_hits"] = int(round(item["total_hits"]))
            item["positive_hits"] = int(round(item["positive_hits"]))
            item["negative_hits"] = int(round(item["negative_hits"]))
            theme_rows.append(item)

        theme_rows.sort(
            key=lambda item: (item["priority_score"], item["negative_hits"]),
            reverse=True,
        )

        component_breakdown = []
        for component_id, item in component_metrics.items():
            component_breakdown.append(
                {
                    "component_id": component_id,
                    "label": item["metrics"]["label"],
                    "weight": round(item["weight"], 3),
                    "current_score": round(float(item["metrics"]["current_score"]), 1),
                    "projected_score": round(float(item["metrics"]["projected_score"]), 1),
                }
            )
        component_breakdown.sort(key=lambda item: item["weight"], reverse=True)

        result = {
            "label": profile["label"],
            "current_score": round(float(current_score), 1),
            "projected_score": round(float(projected_score), 1),
            "delta": round(float(delta), 1),
            "direction": direction,
            "theme_rows": theme_rows,
            "component_breakdown": component_breakdown,
        }
        result["experience_lenses"] = self._experience_lens_rows(dataframe, result)
        return result

    def _experience_lens_rows(self, dataframe, score_metrics):
        """Translate score components into the business meaning of Experience Index."""
        if dataframe.empty:
            return []
        component_by_id = {
            item["component_id"]: item
            for item in score_metrics.get("component_breakdown", [])
        }
        sentiment_summary = self._sentiment_summary(dataframe)
        journey_rows = self._customer_journey_rows(dataframe)
        dominant_journey = journey_rows[0] if journey_rows else None

        def component_score(*component_ids):
            values = [
                component_by_id[component_id]["current_score"]
                for component_id in component_ids
                if component_id in component_by_id
            ]
            if not values:
                return score_metrics.get("current_score", 0.0)
            return round(sum(values) / len(values), 1)

        top_touchpoint = next(
            (
                row["label"]
                for row in score_metrics.get("theme_rows", [])
                if row.get("theme_id") in {"responsiveness", "communication", "schedule", "facility"}
            ),
            "touchpoint layanan utama",
        )
        top_felt_theme = next(
            (
                row["label"]
                for row in score_metrics.get("theme_rows", [])
                if row.get("theme_id") in {"instructor", "material", "outcome"}
            ),
            "rasa manfaat layanan",
        )

        return [
            {
                "lens": "Touchpoint pelanggan",
                "score": component_score("service_score", "facility_score"),
                "reading": f"Interaksi pelanggan paling perlu dibaca pada {top_touchpoint}.",
                "evidence": f"Service Score dan Facility Score dipakai sebagai proksi mutu kontak layanan, kesiapan fasilitas, koordinasi, dan dukungan operasional.",
            },
            {
                "lens": "Pengalaman yang dirasakan",
                "score": score_metrics.get("current_score", 0.0),
                "reading": f"Pelanggan menunjukkan {sentiment_summary['positive_share']}% sinyal positif dan {sentiment_summary['issue_share']}% sinyal korektif tertimbang.",
                "evidence": f"Komentar, rating, kritik konstruktif, dan tema {top_felt_theme} dipakai untuk membaca rasa jelas, relevan, nyaman, atau friksi yang dialami.",
            },
            {
                "lens": "Perjalanan agenda pelanggan",
                "score": component_score("learning_score", "service_score"),
                "reading": f"Tahap paling menentukan saat ini adalah {dominant_journey['stage_label'] if dominant_journey else 'perjalanan agenda yang belum cukup terpetakan'}.",
                "evidence": "Customer journey dibaca dari pra-layanan, kesiapan pelaksanaan, delivery agenda, hingga tindak lanjut dan outcome pasca-layanan.",
            },
        ]

    def _build_analysis_context(self, timeframe_df, timeframe, sentiment, segment, score_engine):
        normalized_sentiment = self._normalize_sentiment_filter(sentiment)
        normalized_segment = self._normalize_segment_filter(segment)
        normalized_score_engine = self._normalize_score_engine(score_engine)
        score_profile = self._score_engine_profile(normalized_score_engine)
        journey_rows = self._customer_journey_rows(timeframe_df)
        score_metrics = self._score_engine_metrics(timeframe_df, normalized_score_engine)
        forecast_30d = self._build_30d_forecast(
            timeframe_df,
            normalized_score_engine,
            score_metrics,
            normalized_sentiment,
            normalized_segment,
        )
        dominant_journey = journey_rows[0] if journey_rows else None
        dominant_theme = score_metrics["theme_rows"][0] if score_metrics["theme_rows"] else None
        location_counts = self._series_counts_for_column(timeframe_df, "Lokasi", limit=5)
        instructor_type_counts = self._series_counts_for_column(timeframe_df, "Tipe Instruktur", limit=5)

        return {
            "timeframe": timeframe, "timeframe_label": readable_timeframe_label(timeframe), "sentiment": normalized_sentiment,
            "sentiment_label": self._label_from_options(SENTIMENT_OPTIONS, normalized_sentiment, "Semua Sentimen"),
            "segment": normalized_segment, "segment_label": normalized_segment if normalized_segment != "all" else "Semua Segmen",
            "score_engine": normalized_score_engine, "score_profile": score_profile, "score_metrics": score_metrics,
            "journey_rows": journey_rows, "dominant_journey": dominant_journey, "dominant_theme": dominant_theme,
            "location_counts": location_counts, "instructor_type_counts": instructor_type_counts,
            "scope_text": self._analysis_scope_text(timeframe, normalized_sentiment, normalized_segment, normalized_score_engine),
            "horizon_text": self._forecast_horizon(timeframe),
            "forecast_30d": forecast_30d,
        }

    def _theme_hits(self, dataframe):
        cache_key = "theme_hits"
        use_prepared_cache = dataframe is self._prepared_analysis_dataframe
        if use_prepared_cache and cache_key in self._prepared_helper_cache:
            return self._prepared_helper_cache[cache_key]
        theme_stats = self._compute_theme_hits(dataframe)
        if use_prepared_cache:
            self._prepared_helper_cache[cache_key] = theme_stats
        return theme_stats

    def _compute_theme_hits(self, dataframe):
        theme_stats = []
        if dataframe.empty: return theme_stats
        comment_series = dataframe["Komentar Lower"].astype(str)
        for theme in self.THEME_LIBRARY:
            match_mask = comment_series.apply(lambda text: any(keyword in text for keyword in theme["keywords"]))
            matched = dataframe[match_mask]
            if matched.empty: continue
            positive_hits = int((matched["Sentiment Label"] == "positive").sum())
            negative_hits = int(matched["Sentiment Label"].isin({"negative", "mixed"}).sum())
            neutral_hits = int((matched["Sentiment Label"] == "neutral").sum())
            theme_stats.append({
                "id": theme["id"], "label": theme["label"], "prescription": theme["prescription"], "total_hits": int(len(matched)),
                "positive_hits": positive_hits, "negative_hits": negative_hits, "neutral_hits": neutral_hits, "matched_df": matched,
            })
        return sorted(theme_stats, key=lambda item: (item["negative_hits"], item["total_hits"]), reverse=True)

    def _quote_lines(self, dataframe, limit=3):
        if dataframe.empty: return ["- Tidak ada kutipan yang cukup untuk periode ini."]
        lines, seen_comments = [], set()
        for _, row in dataframe.iterrows():
            comment = self._truncate_text(row.get("Komentar", ""))
            if not comment or comment in seen_comments: continue
            seen_comments.add(comment)
            lines.append(f'- "{comment}" ({row.get("Tipe Stakeholder", "Stakeholder")} | {row.get("Layanan", "Layanan")} | rating {row.get("Rating", "-")})')
            if len(lines) >= limit: break
        return lines or ["- Tidak ada kutipan yang cukup untuk periode ini."]

    def _group_risk(self, dataframe, column_name, limit=3):
        cache_key = ("group_risk", column_name)
        use_prepared_cache = dataframe is self._prepared_analysis_dataframe
        if use_prepared_cache:
            rows = self._prepared_helper_cache.get(cache_key)
            if rows is None:
                rows = self._compute_group_risk(dataframe, column_name)
                self._prepared_helper_cache[cache_key] = rows
        else:
            rows = self._compute_group_risk(dataframe, column_name)
        return rows[:limit]

    def _compute_group_risk(self, dataframe, column_name):
        if dataframe.empty or column_name not in dataframe.columns: return []
        rows = []
        grouped = dataframe.groupby(column_name, dropna=False)
        for label, group in grouped:
            clean_label = str(label).strip() or "Tidak terklasifikasi"
            rating_avg = group["Rating Numeric"].mean()
            issue_weight = group["Sentiment Label"].apply(self._sentiment_issue_weight).sum()
            negative_ratio = issue_weight / max(len(group), 1)
            volume = len(group)
            safe_avg_rating = round(rating_avg, 2) if pd.notna(rating_avg) else 0.0
            risk_score = round((negative_ratio * 70) + ((5 - safe_avg_rating) * 6) + min(volume, 10), 1)
            rows.append({
                "label": clean_label,
                "volume": volume,
                "average_rating": safe_avg_rating,
                "negative_ratio": round(negative_ratio * 100, 1),
                "risk_score": risk_score,
                "confidence": _confidence_tier(volume),
            })
        rows.sort(key=lambda item: item["risk_score"], reverse=True)
        return rows

    def _governance_summary(self, timeframe_df):
        cache_key = "governance_summary"
        use_prepared_cache = timeframe_df is self._prepared_analysis_dataframe
        if use_prepared_cache and cache_key in self._prepared_helper_cache:
            return self._prepared_helper_cache[cache_key]
        summary = self._compute_governance_summary(timeframe_df)
        if use_prepared_cache:
            self._prepared_helper_cache[cache_key] = summary
        return summary

    def _compute_governance_summary(self, timeframe_df):
        total_rows = self._raw_response_count(timeframe_df)
        if total_rows == 0: return {"total_rows": 0, "completeness_pct": 0.0, "source_count": 0, "channel_count": 0}
        completeness_scores = [(timeframe_df[field].astype(str).str.strip() != "").mean() for field in ["Tipe Stakeholder", "Layanan", "Rentang Waktu", "Komentar"]]
        source_count = max(self._series_counts(timeframe_df["Sumber Feedback"], limit=20).shape[0], 1)
        channel_count = self._series_counts(timeframe_df["Kanal Feedback"], limit=20).shape[0]
        return {
            "total_rows": total_rows,
            "dimension_count": self._dimension_count(timeframe_df),
            "rating_response_count": self._rating_response_count(timeframe_df),
            "text_response_count": self._text_response_count(timeframe_df),
            "completeness_pct": round(sum(completeness_scores) / len(completeness_scores) * 100, 1),
            "source_count": source_count,
            "channel_count": channel_count,
        }

    @staticmethod
    def _rating_assessment(avg_rating):
        if pd.isna(avg_rating): return "belum dapat dinilai secara memadai"
        if avg_rating >= 4.3: return "sangat baik dan relatif konsisten"
        if avg_rating >= 3.75: return "baik, tetapi masih menyisakan beberapa titik perbaikan"
        if avg_rating >= 3.0: return "cukup, namun belum cukup stabil untuk dianggap kuat"
        return "masih lemah dan memerlukan perhatian manajemen segera"

    @staticmethod
    def _negative_share_assessment(negative_share):
        if negative_share >= 35: return "cukup tinggi dan berpotensi mengganggu persepsi layanan jika tidak segera ditangani"
        if negative_share >= 20: return "perlu diawasi karena dapat berkembang menjadi isu yang lebih luas"
        if negative_share > 0: return "masih dalam batas terkendali, namun tetap membutuhkan pemantauan"
        return "belum menunjukkan sinyal keluhan yang berarti"

    @staticmethod
    def _risk_severity(risk_score):
        return "tinggi" if risk_score >= 55 else "menengah" if risk_score >= 40 else "terkendali"

    @staticmethod
    def _primary_label(series_counts, fallback):
        return series_counts.index[0] if not series_counts.empty else fallback

    @staticmethod
    def _format_count_summary(series_counts, unit="feedback", limit=3):
        if series_counts.empty: return "belum terpetakan"
        return ", ".join(f"{label} ({count} {unit})" for label, count in series_counts.head(limit).items())

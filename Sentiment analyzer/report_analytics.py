import re
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
from data_pipeline import CANONICAL_INTERNAL_COLUMNS
from report_narratives import ReportNarrativeBuilderMixin

class FeedbackAnalyticsEngine(ReportNarrativeBuilderMixin):
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
        self.full_df = dataframe.copy() if dataframe is not None else pd.DataFrame()
        for column_name in CANONICAL_INTERNAL_COLUMNS:
            if column_name not in self.full_df.columns:
                self.full_df[column_name] = ""
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
        scope_parts = [f"periode {timeframe}", f"perspektif {profile['label']}"]
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

    def _filter_timeframe(self, timeframe):
        return self._filter_view(timeframe)

    def _filter_view(self, timeframe, sentiment="all", segment="all"):
        if self.full_df.empty: return self.full_df.copy()
        filtered = self.full_df[self.full_df["Rentang Waktu"] == timeframe].copy()
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
            positive_count = int((stage_df["Sentiment Label"] == "positive").sum())
            neutral_count = int((stage_df["Sentiment Label"] == "neutral").sum())
            negative_count = int((stage_df["Sentiment Label"] == "negative").sum())
            stage_theme_hits = self._theme_hits(stage_df)
            dominant_theme = next((theme["label"] for theme in stage_theme_hits if theme["negative_hits"] > 0 or theme["positive_hits"] > 0), "Sinyal umum customer journey")

            rows.append({
                "stage_id": stage["id"], "stage_label": label, "description": stage["description"], "volume": total,
                "average_rating": round(stage_df["Rating Numeric"].mean(), 2) if stage_df["Rating Numeric"].notna().any() else 0.0,
                "positive_share": self._safe_percentage(positive_count, total), "neutral_share": self._safe_percentage(neutral_count, total),
                "negative_share": self._safe_percentage(negative_count, total), "dominant_theme": dominant_theme,
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
            }

        avg_rating = dataframe["Rating Numeric"].mean()
        base_score = ((avg_rating / 5) * 100) if pd.notna(avg_rating) else 0.0
        total_rows = len(dataframe)
        positive_share = self._safe_percentage(int((dataframe["Sentiment Label"] == "positive").sum()), total_rows)
        negative_share = self._safe_percentage(int((dataframe["Sentiment Label"] == "negative").sum()), total_rows)

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

        current_score = self._clamp((base_score * 0.72) + ((50 + (weighted_balance * 50)) * 0.28))
        top_weighted_risk = max(((row["negative_hits"] / max(row["total_hits"], 1)) * row["weight"] for row in theme_rows), default=0.0)
        delta = round((((positive_share - negative_share) / 100) * 6) - (weighted_negative_ratio * 11) + (weighted_positive_ratio * 4) - (top_weighted_risk * 3), 1)
        if abs(delta) < 0.6: delta = 0.0
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
        if abs(delta) < 0.6:
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

        return {
            "label": profile["label"],
            "current_score": round(float(current_score), 1),
            "projected_score": round(float(projected_score), 1),
            "delta": round(float(delta), 1),
            "direction": direction,
            "theme_rows": theme_rows,
            "component_breakdown": component_breakdown,
        }

    def _build_analysis_context(self, timeframe_df, timeframe, sentiment, segment, score_engine):
        normalized_sentiment = self._normalize_sentiment_filter(sentiment)
        normalized_segment = self._normalize_segment_filter(segment)
        normalized_score_engine = self._normalize_score_engine(score_engine)
        score_profile = self._score_engine_profile(normalized_score_engine)
        journey_rows = self._customer_journey_rows(timeframe_df)
        score_metrics = self._score_engine_metrics(timeframe_df, normalized_score_engine)
        dominant_journey = journey_rows[0] if journey_rows else None
        dominant_theme = score_metrics["theme_rows"][0] if score_metrics["theme_rows"] else None
        location_counts = self._series_counts_for_column(timeframe_df, "Lokasi", limit=5)
        instructor_type_counts = self._series_counts_for_column(timeframe_df, "Tipe Instruktur", limit=5)

        return {
            "timeframe": timeframe, "sentiment": normalized_sentiment,
            "sentiment_label": self._label_from_options(SENTIMENT_OPTIONS, normalized_sentiment, "Semua Sentimen"),
            "segment": normalized_segment, "segment_label": normalized_segment if normalized_segment != "all" else "Semua Segmen",
            "score_engine": normalized_score_engine, "score_profile": score_profile, "score_metrics": score_metrics,
            "journey_rows": journey_rows, "dominant_journey": dominant_journey, "dominant_theme": dominant_theme,
            "location_counts": location_counts, "instructor_type_counts": instructor_type_counts,
            "scope_text": self._analysis_scope_text(timeframe, normalized_sentiment, normalized_segment, normalized_score_engine),
            "horizon_text": self._forecast_horizon(timeframe),
        }

    def _theme_hits(self, dataframe):
        theme_stats = []
        if dataframe.empty: return theme_stats
        comment_series = dataframe["Komentar Lower"].astype(str)
        for theme in self.THEME_LIBRARY:
            match_mask = comment_series.apply(lambda text: any(keyword in text for keyword in theme["keywords"]))
            matched = dataframe[match_mask]
            if matched.empty: continue
            positive_hits = int((matched["Sentiment Label"] == "positive").sum())
            negative_hits = int((matched["Sentiment Label"] == "negative").sum())
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
        if dataframe.empty or column_name not in dataframe.columns: return []
        rows = []
        grouped = dataframe.groupby(column_name, dropna=False)
        for label, group in grouped:
            clean_label = str(label).strip() or "Tidak terklasifikasi"
            rating_avg = group["Rating Numeric"].mean()
            negative_ratio = (group["Sentiment Label"] == "negative").mean()
            volume = len(group)
            safe_avg_rating = round(rating_avg, 2) if pd.notna(rating_avg) else 0.0
            risk_score = round((negative_ratio * 70) + ((5 - safe_avg_rating) * 6) + min(volume, 10), 1)
            rows.append({"label": clean_label, "volume": volume, "average_rating": safe_avg_rating, "negative_ratio": round(negative_ratio * 100, 1), "risk_score": risk_score})
        rows.sort(key=lambda item: item["risk_score"], reverse=True)
        return rows[:limit]

    def _governance_summary(self, timeframe_df):
        total_rows = len(timeframe_df)
        if total_rows == 0: return {"total_rows": 0, "completeness_pct": 0.0, "source_count": 0, "channel_count": 0}
        completeness_scores = [(timeframe_df[field].astype(str).str.strip() != "").mean() for field in ["Tipe Stakeholder", "Layanan", "Rentang Waktu", "Komentar"]]
        source_count = max(self._series_counts(timeframe_df["Sumber Feedback"], limit=20).shape[0], 1)
        channel_count = self._series_counts(timeframe_df["Kanal Feedback"], limit=20).shape[0]
        return {"total_rows": total_rows, "completeness_pct": round(sum(completeness_scores) / len(completeness_scores) * 100, 1), "source_count": source_count, "channel_count": channel_count}

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

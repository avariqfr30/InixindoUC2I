import re
from datetime import datetime

import pandas as pd


class FeedbackSpecialistAgent:
    role = "Specialist"
    dataset_focus = ()

    @staticmethod
    def _clean_text(value):
        return re.sub(r"\s+", " ", str(value or "")).strip()

    @staticmethod
    def _primary_label(series_counts, fallback="Belum terpetakan"):
        if series_counts is None or getattr(series_counts, "empty", True):
            return fallback
        return str(series_counts.index[0])

    @staticmethod
    def _confidence_from_context(context, evidence_type):
        governance = context["governance"]
        raw_count = int(governance.get("total_rows", 0) or 0)
        completeness = float(governance.get("completeness_pct", 0.0) or 0.0)
        text_count = int(governance.get("text_response_count", 0) or 0)
        if evidence_type == "osint":
            return "Sedang" if context.get("has_osint_signal") else "Rendah"
        if evidence_type == "text":
            if text_count >= 20:
                return "Tinggi"
            if text_count >= 5:
                return "Sedang"
            return "Rendah"
        if raw_count >= 100 and completeness >= 80:
            return "Tinggi"
        if raw_count >= 20 and completeness >= 60:
            return "Sedang"
        return "Rendah"

    def run(self, engine, dataframe, context):
        raise NotImplementedError


class DataStewardAgent(FeedbackSpecialistAgent):
    role = "Data Steward"
    dataset_focus = ("raw_count", "field_completeness", "source_channel")

    def run(self, engine, dataframe, context):
        governance = context["governance"]
        finding = (
            f"Cakupan data terdiri dari {governance['total_rows']} respons mentah yang diringkas menjadi "
            f"{governance.get('dimension_count', len(dataframe))} dimensi evaluasi dengan kelengkapan field inti "
            f"{governance['completeness_pct']}%."
        )
        implication = (
            "Data layak dipakai sebagai dasar pembacaan manajemen, tetapi keputusan tetap perlu mencatat batas cakupan kanal "
            f"karena hanya {governance['source_count']} sumber dan {governance['channel_count']} kanal yang terpetakan."
        )
        return {
            "role": self.role,
            "dataset": "Operational evidence",
            "evidence_type": "operational_evidence",
            "confidence": self._confidence_from_context(context, "operational_evidence"),
            "finding": finding,
            "implication": implication,
        }


class RatingAnalystAgent(FeedbackSpecialistAgent):
    role = "Rating Analyst"
    dataset_focus = ("rating", "score_engine", "risk_service")

    def run(self, engine, dataframe, context):
        average_rating = dataframe["Rating Numeric"].mean() if "Rating Numeric" in dataframe.columns else pd.NA
        rating_text = round(average_rating, 2) if pd.notna(average_rating) else 0.0
        score_metrics = context["score_metrics"]
        top_risk = context["top_risk"][0]["label"] if context["top_risk"] else self._primary_label(context["top_service"])
        finding = (
            f"Rating rata-rata berada di {rating_text}/5, sementara {context['score_profile']['label']} berada pada "
            f"{score_metrics['current_score']} dan diproyeksikan {score_metrics['direction']} ke {score_metrics['projected_score']}."
        )
        implication = f"Prioritas evaluasi angka sebaiknya dimulai dari {top_risk}, bukan dari seluruh layanan secara merata."
        return {
            "role": self.role,
            "dataset": "Rating response",
            "evidence_type": "rating",
            "confidence": self._confidence_from_context(context, "rating"),
            "finding": finding,
            "implication": implication,
        }


class VoiceOfCustomerAgent(FeedbackSpecialistAgent):
    role = "Voice-of-Customer Analyst"
    dataset_focus = ("comment", "theme", "representative_why")

    def run(self, engine, dataframe, context):
        top_issue = context["top_issue"]
        issue_label = top_issue["label"] if top_issue else "konsistensi pengalaman layanan"
        comment_count = context["governance"].get("text_response_count", 0)
        quote = ""
        if "Komentar" in dataframe.columns:
            comments = dataframe["Komentar"].fillna("").astype(str).str.strip()
            comments = comments[comments != ""]
            if not comments.empty:
                quote = self._clean_text(comments.iloc[0])[:160]
        finding = f"{comment_count} komentar teks dipakai untuk membaca alasan di balik rating; tema utama yang perlu dijelaskan adalah {issue_label}."
        implication = (
            f"Bukti naratif terkuat saat ini: {quote}" if quote else
            "Bukti naratif masih terbatas, sehingga rekomendasi perlu lebih mengandalkan pola rating dan dimensi layanan."
        )
        return {
            "role": self.role,
            "dataset": "Text response / komentar",
            "evidence_type": "text",
            "confidence": self._confidence_from_context(context, "text"),
            "finding": finding,
            "implication": implication,
        }


class ExternalContextAgent(FeedbackSpecialistAgent):
    role = "External Context Analyst"
    dataset_focus = ("osint", "market_signal", "future_state")

    @staticmethod
    def _first_osint_signal(macro_trends):
        for line in str(macro_trends or "").splitlines():
            cleaned = line.strip()
            if re.match(r"^\d+\.", cleaned):
                return re.sub(r"^\d+\.\s*", "", cleaned).split(" | ")[0].strip()
        return ""

    def run(self, engine, dataframe, context):
        signal = self._first_osint_signal(context.get("macro_trends"))
        top_issue = context["top_issue"]["label"] if context["top_issue"] else "kualitas layanan"
        if signal:
            finding = f"Sinyal eksternal paling relevan untuk pembanding saat ini adalah {signal}."
        else:
            finding = "Sinyal eksternal belum cukup kuat; OSINT harus dipakai sebagai konteks, bukan sumber fakta operasional."
        implication = f"Kaitkan sinyal eksternal ke kondisi internal melalui risiko {top_issue}, agar rekomendasi tidak berhenti sebagai tren umum."
        return {
            "role": self.role,
            "dataset": "OSINT benchmark",
            "evidence_type": "osint",
            "confidence": self._confidence_from_context(context, "osint"),
            "finding": finding,
            "implication": implication,
        }


class ActionPlannerAgent(FeedbackSpecialistAgent):
    role = "Action Planner"
    dataset_focus = ("priority", "owner", "30_day_action")

    def run(self, engine, dataframe, context):
        top_risk = context["top_risk"][0]["label"] if context["top_risk"] else self._primary_label(context["top_service"])
        top_issue = context["top_issue"]["label"] if context["top_issue"] else "konsistensi kualitas layanan"
        journey = context["dominant_journey"]["stage_label"] if context["dominant_journey"] else "customer journey utama"
        finding = f"Tiga fokus eksekusi adalah {top_risk}, {top_issue}, dan titik journey {journey}."
        implication = "Tetapkan owner, target 30 hari, dan indikator bukti perbaikan sebelum rekomendasi diperluas ke area lain."
        return {
            "role": self.role,
            "dataset": "Prioritized analytics",
            "evidence_type": "recommendation",
            "confidence": self._confidence_from_context(context, "rating"),
            "finding": finding,
            "implication": implication,
        }


class FeedbackProposalTeam:
    """Deterministic specialist workflow that simulates an internal report team.

    These are not independent LLM calls. They are bounded analysis roles over
    prepared datasets, which keeps concurrency predictable and factual claims
    traceable to the existing analytics context.
    """

    specialists = (
        DataStewardAgent(),
        RatingAnalystAgent(),
        VoiceOfCustomerAgent(),
        ExternalContextAgent(),
        ActionPlannerAgent(),
    )

    @staticmethod
    def _sources_used(specialists):
        source_tokens = set()
        for item in specialists:
            dataset = str(item.get("dataset", "")).lower()
            if "rating" in dataset:
                source_tokens.add("rating")
            if "text" in dataset or "komentar" in dataset:
                source_tokens.add("komentar")
            if "osint" in dataset:
                source_tokens.add("osint")
            if "operational" in dataset:
                source_tokens.add("evaluasi_layanan")
        return sorted(source_tokens)

    @staticmethod
    def _manager_summary(context):
        governance = context["governance"]
        top_risk = context["top_risk"][0]["label"] if context["top_risk"] else "layanan prioritas"
        top_issue = context["top_issue"]["label"] if context["top_issue"] else "konsistensi kualitas layanan"
        return (
            f"Tim analis membaca {governance['total_rows']} respons mentah sebagai bahan keputusan: "
            f"validasi data, pola rating, alasan komentar, konteks eksternal, dan rencana aksi mengarah ke prioritas "
            f"{top_risk} dengan perhatian utama pada {top_issue}."
        )[:420]

    @staticmethod
    def _overall_confidence(context, specialist_outputs):
        confidence_values = [item.get("confidence", "Rendah") for item in specialist_outputs]
        if "Rendah" in confidence_values:
            return "Rendah"
        governance = context["governance"]
        if (
            int(governance.get("total_rows", 0) or 0) >= 100
            and float(governance.get("completeness_pct", 0.0) or 0.0) >= 80
            and int(governance.get("text_response_count", 0) or 0) >= 20
        ):
            return "Tinggi"
        return "Sedang"

    @staticmethod
    def _evidence_ledger(context, specialist_outputs):
        governance = context["governance"]
        ledger = [
            {
                "evidence_type": "operational_evidence",
                "source": "Basis evaluasi layanan",
                "detail": f"{governance['total_rows']} respons mentah; {governance.get('dimension_count', 0)} dimensi evaluasi.",
            },
            {
                "evidence_type": "rating",
                "source": "Rating response",
                "detail": f"{governance.get('rating_response_count', 0)} rating dipakai untuk skor dan risiko.",
            },
            {
                "evidence_type": "komentar",
                "source": "Text response / komentar",
                "detail": f"{governance.get('text_response_count', 0)} komentar dipakai untuk menjelaskan alasan rating.",
            },
        ]
        if any(item.get("evidence_type") == "osint" for item in specialist_outputs):
            ledger.append(
                {
                    "evidence_type": "osint",
                    "source": "Pembanding eksternal",
                    "detail": "Dipakai sebagai konteks eksternal, bukan pengganti bukti operasional.",
                }
            )
        return ledger

    @staticmethod
    def _qa_review(context, overall_confidence):
        governance = context["governance"]
        notes = [
            "temuan evaluasi dipisahkan dari konteks pembanding eksternal.",
            "temuan evaluasi didukung oleh rating, komentar, dan prioritas risiko sebelum menjadi rekomendasi.",
        ]
        if overall_confidence != "Tinggi":
            notes.append("temuan evaluasi memerlukan review manajemen sebelum menjadi arah kebijakan final.")
        if int(governance.get("text_response_count", 0) or 0) == 0:
            notes.append("temuan evaluasi memiliki bukti teks terbatas, sehingga penjelasan rating perlu dibaca hati-hati.")
        if not context.get("has_osint_signal"):
            notes.append("temuan evaluasi menjadi rujukan utama karena sinyal eksternal lemah atau belum tersedia.")
        return notes

    @staticmethod
    def _safe_rating_stats(dataframe):
        if dataframe is None or dataframe.empty or "Rating Numeric" not in dataframe.columns:
            return 0, 0.0, 0.0
        ratings = pd.to_numeric(dataframe["Rating Numeric"], errors="coerce")
        rated = ratings.dropna()
        if rated.empty:
            return 0, 0.0, 0.0
        negative_share = round((rated[rated <= 2].count() / rated.count()) * 100, 1)
        return int(rated.count()), round(float(rated.mean()), 2), negative_share

    @staticmethod
    def _audit_trail(context, timeframe, sentiment, segment, score_engine):
        governance = context["governance"]
        return {
            "generated_at_utc": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "timeframe": timeframe,
            "sentiment": sentiment,
            "segment": segment,
            "score_engine": score_engine,
            "raw_response_count": int(governance.get("total_rows", 0) or 0),
            "dimension_count": int(governance.get("dimension_count", 0) or 0),
            "rating_response_count": int(governance.get("rating_response_count", 0) or 0),
            "text_response_count": int(governance.get("text_response_count", 0) or 0),
            "field_completeness_pct": float(governance.get("completeness_pct", 0.0) or 0.0),
            "source_count": int(governance.get("source_count", 0) or 0),
            "channel_count": int(governance.get("channel_count", 0) or 0),
            "osint_signal_available": bool(context.get("has_osint_signal")),
        }

    @classmethod
    def _contradiction_review(cls, dataframe):
        rating_count, average_rating, negative_share = cls._safe_rating_stats(dataframe)
        comments = dataframe.get("Komentar", pd.Series(dtype="object")).fillna("").astype(str).str.lower()
        negative_terms = ("lambat", "kurang", "tidak", "belum", "sulit", "padat", "delay", "keluhan", "masalah")
        positive_terms = ("baik", "puas", "membantu", "jelas", "relevan", "cepat", "bagus")
        negative_text_hits = int(comments.apply(lambda text: any(term in text for term in negative_terms)).sum())
        positive_text_hits = int(comments.apply(lambda text: any(term in text for term in positive_terms)).sum())
        severity = "Rendah"
        alignment = "rating dan komentar relatif sejalan"
        if average_rating >= 4 and negative_text_hits >= max(3, int(rating_count * 0.2)):
            severity = "Tinggi"
            alignment = "rating tinggi tetapi komentar masih memuat banyak sinyal negatif"
        elif average_rating <= 2.5 and positive_text_hits >= max(3, int(rating_count * 0.2)):
            severity = "Sedang"
            alignment = "rating rendah tetapi komentar masih memuat sinyal positif yang perlu dipisahkan"
        elif negative_share >= 20 and negative_text_hits == 0:
            severity = "Sedang"
            alignment = "rating negatif tidak didukung banyak komentar negatif eksplisit"
        return {
            "rating_text_alignment": alignment,
            "severity": severity,
            "average_rating": average_rating,
            "negative_rating_share": negative_share,
            "negative_text_hits": negative_text_hits,
            "positive_text_hits": positive_text_hits,
        }

    @classmethod
    def _trend_review(cls, engine, scoped_df, timeframe):
        current_count, current_rating, current_negative_share = cls._safe_rating_stats(scoped_df)
        if engine.full_df.empty or "Rentang Waktu" not in engine.full_df.columns:
            return {
                "comparison_period": "Tidak tersedia",
                "rating_delta": 0.0,
                "negative_share_delta": 0.0,
                "reading": "Belum ada periode pembanding yang cukup untuk membaca tren historis.",
            }
        periods = [
            value
            for value in engine.full_df["Rentang Waktu"].dropna().astype(str).str.strip().unique().tolist()
            if value and value != timeframe
        ]
        if not periods:
            return {
                "comparison_period": "Tidak tersedia",
                "rating_delta": 0.0,
                "negative_share_delta": 0.0,
                "reading": "Belum ada periode pembanding yang cukup untuk membaca tren historis.",
            }
        comparison_period = sorted(periods)[-1]
        comparison_df = engine._filter_view(comparison_period)
        comparison_count, comparison_rating, comparison_negative_share = cls._safe_rating_stats(comparison_df)
        rating_delta = round(current_rating - comparison_rating, 2)
        negative_delta = round(current_negative_share - comparison_negative_share, 1)
        if current_count == 0 or comparison_count == 0:
            reading = "Periode pembanding tersedia, tetapi volume rating belum cukup untuk membaca arah tren secara kuat."
        elif rating_delta > 0:
            reading = f"Rating membaik {rating_delta} poin dibanding {comparison_period}; pantau apakah penurunan keluhan mengikuti."
        elif rating_delta < 0:
            reading = f"Rating melemah {abs(rating_delta)} poin dibanding {comparison_period}; perlu cek apakah isu berulang meningkat."
        else:
            reading = f"Rating relatif stabil dibanding {comparison_period}; fokus pada perubahan tema keluhan dan risiko layanan."
        return {
            "comparison_period": comparison_period,
            "rating_delta": rating_delta,
            "negative_share_delta": negative_delta,
            "reading": reading,
        }

    @staticmethod
    def _prediction_review(context):
        metrics = context["score_metrics"]
        return {
            "method": "Early warning projection berbasis rating, sentimen, tema risiko, dan score engine; bukan model statistik forecast.",
            "statistical_forecast": False,
            "direction": metrics.get("direction", "stabil"),
            "current_score": metrics.get("current_score"),
            "projected_score": metrics.get("projected_score"),
            "confidence_note": "Gunakan sebagai sinyal prioritas manajemen, lalu validasi dengan owner layanan sebelum menjadi target operasional final.",
        }

    def run(self, engine, dataframe, timeframe, macro_trends="", sentiment="all", segment="all", score_engine="experience_index"):
        scoped_df = engine._filter_view(timeframe, sentiment=sentiment, segment=segment)
        analysis_context = engine._build_analysis_context(scoped_df, timeframe, sentiment, segment, score_engine)
        context = {
            **analysis_context,
            "governance": engine._governance_summary(scoped_df),
            "top_service": engine._series_counts(scoped_df["Layanan"], limit=1),
            "top_risk": engine._group_risk(scoped_df, "Layanan", limit=1),
            "top_issue": next((theme for theme in engine._theme_hits(scoped_df) if theme["negative_hits"] > 0), None),
            "macro_trends": macro_trends,
            "has_osint_signal": bool(ExternalContextAgent._first_osint_signal(macro_trends)),
        }
        specialist_outputs = [specialist.run(engine, scoped_df, context) for specialist in self.specialists]
        overall_confidence = self._overall_confidence(context, specialist_outputs)
        audit_trail = self._audit_trail(context, timeframe, sentiment, segment, score_engine)
        return {
            "manager_summary": self._manager_summary(context),
            "sources_used": self._sources_used(specialist_outputs),
            "confidence": overall_confidence,
            "audit_trail": audit_trail,
            "contradiction_review": self._contradiction_review(scoped_df),
            "trend_review": self._trend_review(engine, scoped_df, timeframe),
            "prediction_review": self._prediction_review(context),
            "evidence_ledger": self._evidence_ledger(context, specialist_outputs),
            "qa_review": self._qa_review(context, overall_confidence),
            "specialists": specialist_outputs,
        }

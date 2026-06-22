import hashlib
import json
import os
import re
from datetime import datetime
from dataclasses import dataclass

import pandas as pd

from config import OLLAMA_HOST
from reasoning_policy import FeedbackHotsReasoningPolicy


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


@dataclass(frozen=True)
class AgentPassContract:
    role: str
    objective: str
    evidence_type: str
    required_context_keys: tuple
    output_fields: tuple = ("finding", "implication", "confidence")

    def prompt(self, context_packet):
        return "\n".join(
            [
                f"Role: {self.role}",
                f"Objective: {self.objective}",
                "Use only the supplied context packet.",
                FeedbackHotsReasoningPolicy.prompt_block(),
                "Do not reveal role names, agent status, source endpoints, or internal ledger labels in reader-facing prose.",
                "Return JSON with keys: " + ", ".join(self.output_fields) + ".",
                "Context packet:",
                json.dumps(context_packet, ensure_ascii=False, sort_keys=True, default=str),
            ]
        )


class HiddenAgentDesk:
    """Internal single-model, multi-pass desk for report quality evidence.

    The default path remains deterministic so report generation is stable in tests
    and offline deployments. Operators can enable live Ollama passes with
    REPORT_AGENT_DESK_MODE=ollama; every role then uses the same LLM_MODEL value
    while keeping role separation in the prompt contract.
    """

    contracts = (
        AgentPassContract(
            "Data Steward",
            "Validate coverage, completeness, and source-channel limits.",
            "operational_evidence",
            ("governance",),
        ),
        AgentPassContract(
            "Rating Analyst",
            "Read score direction, rating shape, and the highest-risk service.",
            "rating",
            ("score_metrics", "score_profile", "top_risk"),
        ),
        AgentPassContract(
            "Voice-of-Customer Analyst",
            "Explain the strongest customer voice behind ratings.",
            "text",
            ("governance", "top_issue"),
        ),
        AgentPassContract(
            "External Context Analyst",
            "Keep external research as context and never as a substitute for internal evidence.",
            "osint",
            ("macro_trends", "has_osint_signal"),
        ),
        AgentPassContract(
            "Action Planner",
            "Translate evidence into bounded management actions.",
            "recommendation",
            ("top_risk", "top_issue", "dominant_journey"),
        ),
    )

    forbidden_reader_terms = (
        "Agent",
        "Data Steward",
        "Rating Analyst",
        "Voice-of-Customer Analyst",
        "External Context Analyst",
        "Action Planner",
        "Evidence Ledger",
        "QA Guardrail",
        "Report Audit Trail",
        "endpoint",
        "source=",
        "APIDog",
        "Internal API",
    )

    def __init__(self, model_name=None, model_client=None, mode=None):
        self.model_name = model_name or os.getenv("LLM_MODEL", "gpt-oss:120b-cloud")
        self.model_client = model_client
        self.mode = (mode or os.getenv("REPORT_AGENT_DESK_MODE", "deterministic")).strip().lower()

    @staticmethod
    def _context_packet(context, contract):
        return {
            key: context.get(key)
            for key in contract.required_context_keys
            if key in context
        }

    @staticmethod
    def _stable_evidence_id(evidence_type, source, detail):
        payload = json.dumps(
            {"detail": detail, "evidence_type": evidence_type, "source": source},
            ensure_ascii=False,
            sort_keys=True,
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]

    @classmethod
    def evidence_card(cls, evidence_type, source, detail):
        return {
            "evidence_id": cls._stable_evidence_id(evidence_type, source, detail),
            "evidence_type": evidence_type,
            "source": source,
            "detail": detail,
        }

    def _client(self):
        if self.model_client is not None:
            return self.model_client
        if self.mode != "ollama":
            return None
        try:
            from ollama import Client
        except Exception:
            return None
        return Client(host=OLLAMA_HOST)

    @staticmethod
    def _parse_model_response(response):
        if not response:
            return {}
        if isinstance(response, dict):
            content = response.get("message", {}).get("content", "")
        else:
            content = getattr(getattr(response, "message", None), "content", "")
        if not content:
            return {}
        try:
            parsed = json.loads(content)
        except json.JSONDecodeError:
            return {"finding": str(content).strip()[:500]}
        return parsed if isinstance(parsed, dict) else {}

    def _run_model_pass(self, client, contract, context_packet):
        if client is None:
            return {"status": "skipped", "content": {}}
        prompt = contract.prompt(context_packet)
        try:
            response = client.chat(
                model=self.model_name,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a hidden report desk. Keep output concise, evidence-bound, and source-safe.",
                    },
                    {"role": "user", "content": prompt},
                ],
                format="json",
                options={"temperature": 0},
            )
        except Exception as exc:
            return {"status": "failed", "error": str(exc)[:240], "content": {}}
        return {"status": "completed", "content": self._parse_model_response(response)}

    def run_passes(self, context):
        client = self._client()
        passes = []
        for contract in self.contracts:
            packet = self._context_packet(context, contract)
            model_result = self._run_model_pass(client, contract, packet)
            passes.append(
                {
                    "role": contract.role,
                    "model": self.model_name,
                    "evidence_type": contract.evidence_type,
                    "required_context_keys": list(contract.required_context_keys),
                    "output_contract": list(contract.output_fields),
                    "prompt_contract": contract.prompt(packet),
                    "model_status": model_result["status"],
                    "model_output": model_result.get("content", {}),
                    **({"model_error": model_result["error"]} if "error" in model_result else {}),
                }
            )
        return passes

    def final_editor_review(self, manager_summary, specialist_outputs, evidence_ledger):
        combined = "\n".join(
            [
                manager_summary or "",
                *[item.get("finding", "") for item in specialist_outputs],
                *[item.get("implication", "") for item in specialist_outputs],
            ]
        )
        leaked_terms = [
            term
            for term in self.forbidden_reader_terms
            if re.search(rf"\b{re.escape(term)}\b", combined, flags=re.IGNORECASE)
        ]
        expected_evidence = {"operational_evidence", "rating", "komentar"}
        available_evidence = {item.get("evidence_type") for item in evidence_ledger}
        missing_evidence = sorted(expected_evidence - available_evidence)
        return {
            "reader_safe": not leaked_terms,
            "leaked_terms": sorted(set(leaked_terms)),
            "ledger_complete": not missing_evidence,
            "missing_evidence": missing_evidence,
            "evidence_item_count": len(evidence_ledger),
        }

    @staticmethod
    def final_quality_gate(editor_review, pass_reports):
        contract_failures = [
            item["role"]
            for item in pass_reports
            if not item.get("required_context_keys") or not item.get("output_contract")
        ]
        failed_model_passes = [
            item["role"]
            for item in pass_reports
            if item.get("model_status") == "failed"
        ]
        return {
            "passes": bool(editor_review.get("reader_safe")) and bool(editor_review.get("ledger_complete")) and not contract_failures,
            "contract_failures": contract_failures,
            "failed_model_passes": failed_model_passes,
            "model_failures_block_reader_output": False,
        }


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

    def __init__(self, agent_desk=None):
        self.agent_desk = agent_desk or HiddenAgentDesk()

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
            HiddenAgentDesk.evidence_card(
                "operational_evidence",
                "Basis evaluasi layanan",
                f"{governance['total_rows']} respons mentah; {governance.get('dimension_count', 0)} dimensi evaluasi.",
            ),
            HiddenAgentDesk.evidence_card(
                "rating",
                "Rating response",
                f"{governance.get('rating_response_count', 0)} rating dipakai untuk skor dan risiko.",
            ),
            HiddenAgentDesk.evidence_card(
                "komentar",
                "Text response / komentar",
                f"{governance.get('text_response_count', 0)} komentar dipakai untuk menjelaskan alasan rating.",
            ),
        ]
        if any(item.get("evidence_type") == "osint" for item in specialist_outputs):
            ledger.append(
                HiddenAgentDesk.evidence_card(
                    "osint",
                    "Pembanding eksternal",
                    "Dipakai sebagai konteks eksternal, bukan pengganti bukti operasional.",
                )
            )
        return sorted(ledger, key=lambda item: (item["evidence_type"], item["evidence_id"]))

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
        negative_terms = tuple(term for term in ("lambat", "kurang", "tidak", "belum", "sulit", "padat", "delay", "keluhan", "masalah") if term)
        positive_terms = tuple(term for term in ("baik", "puas", "membantu", "jelas", "relevan", "cepat", "bagus") if term)

        def count_text_hits(terms):
            return sum(1 for text in comments if any(bool(term) and term in text for term in terms))

        negative_text_hits = int(count_text_hits(negative_terms))
        positive_text_hits = int(count_text_hits(positive_terms))
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
        governance = context["governance"]
        top_issue = context["top_issue"]["label"] if context.get("top_issue") else "isu utama belum terpetakan"
        dominant_journey = (context.get("dominant_journey") or {}).get("stage_label", "tahap journey utama")
        return {
            "method": "Early-warning / early warning deterministic projection berbasis rating, sentimen, tema risiko, journey, dan score engine; bukan model statistik forecast.",
            "statistical_forecast": False,
            "direction": metrics.get("direction", "stabil"),
            "current_score": metrics.get("current_score"),
            "projected_score": metrics.get("projected_score"),
            "confidence_note": "Gunakan sebagai sinyal prioritas manajemen, lalu validasi dengan owner layanan sebelum menjadi target operasional final.",
            "confidence_drivers": [
                f"{governance.get('total_rows', 0)} respons dan {governance.get('text_response_count', 0)} komentar teks tersedia.",
                f"Kelengkapan field inti {governance.get('completeness_pct', 0.0)}%.",
                "Konteks eksternal tersedia." if context.get("has_osint_signal") else "Konteks eksternal lemah atau belum tersedia.",
            ],
            "challenge_checks": [
                f"Uji apakah tema {top_issue} benar-benar berulang, bukan keluhan insidental.",
                f"Uji apakah tahap {dominant_journey} memiliki cukup bukti komentar sebelum menjadi agenda intervensi.",
                "Belum diklaim sebagai backtesting statistik; validasi dilakukan dengan membandingkan proyeksi terhadap periode berikutnya.",
            ],
        }

    def run(self, engine, dataframe, timeframe, macro_trends="", sentiment="all", segment="all", score_engine="experience_index", prepared_analysis=None):
        prepared = engine.resolve_prepared_report_analysis(
            prepared_analysis,
            timeframe,
            sentiment=sentiment,
            segment=segment,
            score_engine=score_engine,
        )
        if prepared is None:
            scoped_df = engine._filter_view(timeframe, sentiment=sentiment, segment=segment)
            analysis_context = engine._build_analysis_context(scoped_df, timeframe, sentiment, segment, score_engine)
            governance = engine._governance_summary(scoped_df)
            top_service = engine._series_counts(scoped_df["Layanan"], limit=1)
            top_risk = engine._group_risk(scoped_df, "Layanan", limit=1)
            top_issue = next((theme for theme in engine._theme_hits(scoped_df) if theme["negative_hits"] > 0), None)
            contradiction_review = self._contradiction_review(scoped_df)
        else:
            scoped_df = prepared.scoped_dataframe
            analysis_context = prepared.analysis_context
            governance = prepared.governance_summary
            top_service = prepared.top_service
            top_risk = prepared.top_risk
            top_issue = prepared.top_issue
            contradiction_review = prepared.contradiction_review
        context = {
            **analysis_context,
            "governance": governance,
            "top_service": top_service,
            "top_risk": top_risk,
            "top_issue": top_issue,
            "macro_trends": macro_trends,
            "has_osint_signal": bool(ExternalContextAgent._first_osint_signal(macro_trends)),
        }
        specialist_outputs = [specialist.run(engine, scoped_df, context) for specialist in self.specialists]
        overall_confidence = self._overall_confidence(context, specialist_outputs)
        audit_trail = self._audit_trail(context, timeframe, sentiment, segment, score_engine)
        evidence_ledger = self._evidence_ledger(context, specialist_outputs)
        manager_summary = self._manager_summary(context)
        agent_passes = self.agent_desk.run_passes(context)
        editor_review = self.agent_desk.final_editor_review(
            manager_summary,
            specialist_outputs,
            evidence_ledger,
        )
        final_quality_gate = self.agent_desk.final_quality_gate(editor_review, agent_passes)
        return {
            "manager_summary": manager_summary,
            "sources_used": self._sources_used(specialist_outputs),
            "confidence": overall_confidence,
            "audit_trail": audit_trail,
            "contradiction_review": contradiction_review,
            "trend_review": self._trend_review(engine, scoped_df, timeframe),
            "prediction_review": self._prediction_review(context),
            "evidence_ledger": evidence_ledger,
            "qa_review": self._qa_review(context, overall_confidence),
            "agent_desk": {
                "mode": self.agent_desk.mode,
                "model": self.agent_desk.model_name,
                "passes": agent_passes,
                "editor_review": editor_review,
                "final_quality_gate": final_quality_gate,
            },
            "specialists": specialist_outputs,
        }

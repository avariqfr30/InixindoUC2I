import re
import pandas as pd

from config import ADOPTION_READINESS_PILLARS, CX_SENTIMENT_STRUCTURE, DEFAULT_SCORE_ENGINE
from document_builder import DocumentBuilder
from editorial_intelligence import compact_feedback_table_rows
from management_interpretation import FeedbackManagementInterpreter
from report_trust_sections import build_specialist_review_markdown


class BaseNarrativeMixin:
    @staticmethod
    def _escape_table_cell(value):
        return str(value).replace("|", "\\|").replace("\n", " ").strip()

    @classmethod
    def _markdown_table(cls, headers, rows):
        if not rows:
            return ""
        header_line = "| " + " | ".join(cls._escape_table_cell(item) for item in headers) + " |"
        separator_line = "| " + " | ".join("---" for _ in headers) + " |"
        safe_rows = compact_feedback_table_rows(rows)
        row_lines = ["| " + " | ".join(cls._escape_table_cell(cell) for cell in row) + " |" for row in safe_rows]
        return "\n".join([header_line, separator_line, *row_lines])

    def _distribution_rows(self, series_counts, total_rows, limit=5):
        return [[label, count, f"{self._safe_percentage(count, total_rows)}%"] for label, count in series_counts.head(limit).items()]

    def _extract_osint_signals(self, macro_trends, limit=3):
        signals = []
        for line in str(macro_trends).splitlines():
            cleaned = line.strip()
            if not re.match(r"^\d+\.", cleaned):
                continue
            cleaned = re.sub(r"^\d+\.\s*", "", cleaned)
            source_match = re.search(r"\(Sumber:\s*([^,)]+)(?:,\s*([^)]+))?\)", cleaned)
            if source_match:
                snippet = re.sub(r"\s*\(Sumber:[^)]+\)", "", cleaned).strip()
                signals.append({
                    "title": "Sinyal eksternal",
                    "snippet": snippet,
                    "source": source_match.group(1).strip() or "Tidak diketahui",
                    "date": (source_match.group(2) or "-").strip(),
                })
                if len(signals) >= limit:
                    break
                continue
            parts = [part.strip() for part in cleaned.split(" | ") if part.strip()]
            if not parts:
                continue
            title = parts[0]
            snippet = parts[1] if len(parts) > 1 else ""
            source, date = "Tidak diketahui", "-"
            for part in parts[2:]:
                if part.startswith("sumber="):
                    source = part.split("=", maxsplit=1)[1] or source
                elif part.startswith("tanggal="):
                    date = part.split("=", maxsplit=1)[1] or date
            signals.append({"title": title, "snippet": snippet, "source": source, "date": date})
            if len(signals) >= limit:
                break
        return signals

    @staticmethod
    def _extract_deep_insight(macro_trends):
        match = re.search(r"\*\*Insight Mendalam[^*]*\*\*\s*(.*)", str(macro_trends))
        if match:
            return match.group(0)
        return ""

    @staticmethod
    def _summarized_external_trend_lines(osint_signals, deep_insight):
        source_text = " ".join(
            [
                str(deep_insight or ""),
                *[str(signal.get("snippet") or "") for signal in osint_signals or []],
            ]
        ).lower()
        if not source_text.strip():
            return []
        lines = []
        if any(term in source_text for term in ("skill", "kompetensi", "training", "pelatihan", "cloud", "ai", "security", "siber")):
            lines.append("- Ringkasan tren eksternal: kebutuhan kompetensi digital bergerak cepat, sehingga materi dan contoh praktik perlu rutin diperbarui.")
        if any(term in source_text for term in ("customer", "experience", "loyal", "follow", "cem", "dampak")):
            lines.append("- Implikasi ke layanan: pelanggan semakin menilai pengalaman dari konsistensi tindak lanjut, bukti dampak, dan relevansi hasil setelah kelas.")
        if not lines:
            lines.append("- Ringkasan tren eksternal: konteks publik hanya dipakai sebagai pembanding umum; prioritas tetap ditentukan oleh bukti evaluasi internal.")
        return lines

    @staticmethod
    def _external_company_linkage(osint_signals, deep_insight, top_service_name, top_segment_name, top_issue_label, projection_sentence):
        if deep_insight or osint_signals:
            lead_signal = "Ringkasan tren eksternal menunjukkan bahwa ekspektasi pelanggan terhadap relevansi materi, bukti dampak, dan tindak lanjut layanan perlu dibaca bersama temuan internal."
        else:
            lead_signal = "Sinyal eksternal belum cukup kuat untuk dijadikan pembanding utama."

        current_context = (
            f"kondisi perusahaan saat ini perlu dibaca terhadap sinyal eksternal tersebut karena area internal yang paling rentan adalah "
            f"{top_service_name}, segmen yang paling perlu dipantau adalah {top_segment_name}, dan isu dominan yang harus dikendalikan adalah "
            f"{top_issue_label}."
        )
        future_context = (
            f"Untuk kondisi perusahaan ke depan, {projection_sentence} Jika ekspektasi pasar terhadap bukti dampak, follow-up, dan konsistensi delivery meningkat, "
            "maka keluhan internal yang terlihat kecil hari ini dapat berubah menjadi risiko reputasi, repeat order, dan kepercayaan stakeholder."
        )
        decision_context = (
            "Implikasi manajerialnya adalah OSINT tidak dipakai sebagai pengganti data internal, melainkan sebagai tekanan eksternal yang membantu menentukan "
            "apakah masalah internal perlu ditangani sebagai quick win operasional, perbaikan tata kelola, atau prioritas strategis."
        )
        return lead_signal, current_context, future_context, decision_context

    @staticmethod
    def _theme_owner(theme_id):
        owner_map = {"responsiveness": "Customer Service / Account Management", "schedule": "Operations / Delivery Management", "facility": "Operations / General Affairs", "instructor": "Academic Lead / Service Quality", "material": "Academic Lead / Product Owner", "communication": "Customer Service / Project Coordinator", "outcome": "Service Owner / Quality Assurance"}
        return owner_map.get(theme_id, "Service Owner")

    @staticmethod
    def _theme_outcome(theme_id):
        outcome_map = {"responsiveness": "Waktu respons lebih konsisten dan penutupan isu lebih cepat.", "schedule": "Pengalaman delivery lebih tertata dan beban sesi lebih seimbang.", "facility": "Gangguan operasional di kelas atau sesi layanan dapat ditekan.", "instructor": "Konsistensi kualitas fasilitator meningkat di berbagai layanan.", "material": "Materi lebih relevan dengan kebutuhan peserta dan konteks klien.", "communication": "Ekspektasi stakeholder lebih selaras sejak pra-delivery hingga pasca-delivery.", "outcome": "Nilai manfaat layanan lebih mudah dirasakan dan dibuktikan."}
        return outcome_map.get(theme_id, "Persepsi kualitas layanan membaik secara terukur.")

    @staticmethod
    def _readiness_label(score):
        return "Kuat" if score >= 80 else "Cukup Siap" if score >= 60 else "Perlu Diperkuat" if score >= 40 else "Prioritas Tinggi"

    def _projection_sentence(self, context):
        metrics = context["score_metrics"]
        score_label = context["score_profile"]["forecast_label"]
        horizon_text = context["horizon_text"]
        calendar_reference = self._forecast_calendar_reference(context["timeframe"])

        if metrics["direction"] == "turun":
            return f"{score_label} diproyeksikan turun dari {metrics['current_score']} menjadi sekitar {metrics['projected_score']} dalam {horizon_text}, atau {calendar_reference}, apabila pola saat ini berlanjut."
        if metrics["direction"] == "naik":
            return f"{score_label} diproyeksikan naik dari {metrics['current_score']} menjadi sekitar {metrics['projected_score']} dalam {horizon_text}, atau {calendar_reference}, jika momentum yang ada dapat dipertahankan."
        return f"{score_label} diproyeksikan relatif stabil di kisaran {metrics['projected_score']} dalam {horizon_text}, atau {calendar_reference}, namun tetap perlu dipantau agar tidak bergeser ketika volume feedback bertambah."

    @staticmethod
    def _score_component_formula(components, key_name):
        if not components:
            return 0.0, ""
        total = 0.0
        terms = []
        for item in components:
            weight = float(item.get("weight", 0.0))
            score_value = float(item.get(key_name, 0.0))
            total += weight * score_value
            terms.append(f"({round(weight * 100, 1)}% x {score_value})")
        return round(total, 2), " + ".join(terms)

    def _experience_formula_details(self, context):
        if context.get("score_engine") != "experience_index":
            return None
        components = context.get("score_metrics", {}).get("component_breakdown") or []
        if not components:
            return None

        current_calc, current_formula = self._score_component_formula(components, "current_score")
        projected_calc, projected_formula = self._score_component_formula(components, "projected_score")
        weight_summary = ", ".join(
            f"{item['label']} {round(float(item['weight']) * 100, 1)}%"
            for item in components
        )
        return {
            "weight_summary": weight_summary,
            "current_calc": current_calc,
            "current_formula": current_formula,
            "projected_calc": projected_calc,
            "projected_formula": projected_formula,
        }

    def _experience_lens_markdown(self, context):
        rows = []
        for item in context.get("score_metrics", {}).get("experience_lenses", []):
            rows.append([
                item.get("lens", "-"),
                item.get("score", "-"),
                item.get("reading", "-"),
                item.get("evidence", "-"),
            ])
        if not rows:
            return ""
        return self._markdown_table(
            ["Lensa Experience Index", "Skor", "Pembacaan", "Bukti Analitis"],
            rows,
        )

    def _prediction_defensibility_markdown(self, timeframe_df, context, external_ready):
        governance = self._governance_summary(timeframe_df)
        sentiment_summary = self._sentiment_summary(timeframe_df)
        score_metrics = context["score_metrics"]
        top_theme = (score_metrics.get("theme_rows") or [{}])[0]
        dominant_journey = context.get("dominant_journey") or {}
        text_count = int(governance.get("text_response_count", 0) or 0)
        total_rows = int(governance.get("total_rows", 0) or 0)
        completeness = float(governance.get("completeness_pct", 0.0) or 0.0)
        confidence_level = "Tinggi" if total_rows >= 100 and text_count >= 20 and completeness >= 80 else "Sedang" if total_rows >= 20 and completeness >= 60 else "Rendah"
        mixed_count = sentiment_summary.get("mixed", 0)
        weak_count = sentiment_summary.get("weak_negative", 0)
        high_rating_with_critique = "perlu dicek" if mixed_count else "tidak dominan"
        weak_evidence = "perlu dibaca hati-hati" if weak_count or text_count < 5 else "cukup terkendali"
        osint_check = "tersedia sebagai konteks pembanding" if external_ready else "lemah atau tidak tersedia, sehingga tidak menaikkan klaim prediksi"

        factor_lines = [
            f"- Penilaian dan sentimen: rata-rata penilaian dibaca bersama sinyal positif {sentiment_summary['positive_share']}% dan sinyal korektif tertimbang {sentiment_summary['issue_share']}%.",
            f"- Tema risiko: {top_theme.get('label', 'tema utama yang belum terpetakan')} dipakai untuk melihat konsentrasi isu, bukan sekadar jumlah komentar.",
            f"- Perjalanan agenda: {dominant_journey.get('stage_label', 'tahap agenda yang belum terpetakan')} menunjukkan titik sentuh dan rasa pengalaman yang paling perlu diuji.",
            "- Experience Index: Learning, Service, dan Facility Score hanya menjadi proksi untuk membaca titik sentuh, rasa pengalaman, dan perjalanan pelanggan secara terpadu.",
            "- Konteks eksternal: sinyal publik hanya menjadi tekanan pembanding; keputusan tetap bersandar pada bukti evaluasi internal.",
        ]
        confidence_lines = [
            f"- Volume data: {total_rows} respons dianalisis; tingkat keyakinan operasional {confidence_level}.",
            f"- Komentar teks: {text_count} komentar tersedia untuk menjelaskan alasan di balik penilaian.",
            f"- Kelengkapan konteks: {completeness}% kolom inti tersedia untuk audit konteks layanan.",
            f"- Konteks eksternal: {osint_check}.",
            "- Status proyeksi: belum diklaim sebagai backtesting statistik; dipakai sebagai sinyal prioritas sampai data periode berikutnya memvalidasi arah.",
        ]
        challenge_lines = [
            f"- Challenge check rating vs komentar: {high_rating_with_critique}; kritik konstruktif dipisahkan dari pujian penuh agar tidak terjadi false comfort.",
            f"- Challenge check bukti lemah: {weak_evidence}; rating rendah dengan komentar terlalu tipis tidak langsung dibaca sebagai keluhan kuat.",
            f"- Challenge check konsentrasi: layanan dan segmen berisiko dibaca bersama volume agar satu kelompok kecil tidak otomatis mewakili seluruh pengalaman pelanggan.",
            f"- Challenge check eksternal: {osint_check}.",
        ]
        sensitivity_lines = [
            f"- Sensitivitas terbesar datang dari perubahan proporsi sinyal korektif, terutama bila tema {top_theme.get('label', 'utama')} membaik atau memburuk pada periode berikutnya.",
            f"- Jika tindak lanjut pada {dominant_journey.get('stage_label', 'tahap agenda utama')} membaik, proyeksi {score_metrics.get('direction', 'stabil')} perlu dikalibrasi ulang sebelum menjadi target operasional.",
            "- Prospective validation: periode berikutnya harus membandingkan arah proyeksi ini dengan skor aktual agar model makin teruji secara historis.",
        ]
        return "\n".join([
            "### Model Early-Warning dan Batas Prediksi",
            "Proyeksi ini adalah peringatan dini deterministik untuk prioritas manajemen, bukan forecast statistik dan bukan keputusan otomatis. Nilainya dapat diperdebatkan karena faktor pembentuk, penggerak keyakinan, dan uji kewajaran ditampilkan secara eksplisit.",
            "",
            "### Faktor Pembentuk Proyeksi",
            *factor_lines,
            "",
            "### Penggerak Keyakinan Proyeksi",
            *confidence_lines,
            "",
            "### Challenge Check dan Risiko Salah Baca",
            *challenge_lines,
            "",
            "### Sensitivitas dan Validasi Berikutnya",
            *sensitivity_lines,
        ])

    @staticmethod
    def _section_brief(content, max_words=28):
        text = re.sub(r"```.*?```", " ", str(content or ""), flags=re.DOTALL)
        text = re.sub(r"^\s{0,3}#{1,6}\s*", "", text, flags=re.MULTILINE)
        text = re.sub(r"\|[^\n]*\|", " ", text)
        text = re.sub(r"(?i)\b(?:url|source|sumber)\s*=\s*[^|.\n]+", " ", text)
        text = re.sub(r"\bBukti yang Dipakai\b", " ", text, flags=re.IGNORECASE)
        text = re.sub(r"[*_`>|]+", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        if not text:
            return ""
        sentences = [part.strip(" -;,.") for part in re.split(r"(?<=[.!?])\s+|;\s+", text) if part.strip(" -;,.")]
        if sentences:
            sentences.sort(
                key=lambda sentence: (
                    bool(re.search(r"\b\d+(?:[,.]\d+)?%?\b", sentence)),
                    any(term in sentence.lower() for term in ("risiko", "feedback", "rating", "sentimen", "layanan", "rekomendasi")),
                    len(sentence),
                ),
                reverse=True,
            )
            text = sentences[0]
        words = text.split()
        if len(words) <= max_words:
            return text
        return " ".join(words[:max_words]).rstrip(" ,;:") + "."

import re
import pandas as pd

from config import DEFAULT_SCORE_ENGINE
from document_builder import DocumentBuilder
from management_interpretation import FeedbackManagementInterpreter
from report_trust_sections import build_specialist_review_markdown
from .base import BaseNarrativeMixin


class ExecutiveNarrativeMixin(BaseNarrativeMixin):
    def _executive_headlines(self, total_rows, dimension_count, avg_rating, negative_share, top_risk, top_issue, dominant_journey, score_metrics, context):
        rating_text = round(avg_rating, 2) if pd.notna(avg_rating) else 0.0
        risk_label = top_risk[0]["label"] if top_risk else "layanan prioritas yang belum terpetakan"
        issue_label = top_issue["label"] if top_issue else "konsistensi kualitas layanan"
        journey_label = dominant_journey["stage_label"] if dominant_journey else "customer journey yang masih perlu dipetakan"
        direction = str(score_metrics.get("direction") or "stabil").lower()
        if negative_share <= 0:
            return [
                f"- Rating rata-rata {rating_text}/5 tanpa sinyal negatif; gunakan {risk_label} sebagai praktik baik.",
                f"- Fokus pada {journey_label} agar layanan tetap konsisten.",
                f"- Lensa bergerak {direction} sebagai penguatan pengalaman.",
            ]
        return [
            f"- Rating rata-rata {rating_text}/5, negatif {negative_share}%; gunakan {risk_label} sebagai prioritas perbaikan.",
            f"- Isu {issue_label} pada {journey_label} perlu dikoreksi.",
            f"- Lensa bergerak {direction} sebagai peringatan dini.",
        ]

    def _executive_dashboard_rows(self, total_rows, dimension_count, rating_response_count, text_response_count, avg_rating, negative_share, top_risk, top_service, top_stakeholder, dominant_journey, score_metrics, context):
        rating_text = round(avg_rating, 2) if pd.notna(avg_rating) else 0.0
        top_risk_label = top_risk[0]["label"] if top_risk else self._primary_label(top_service, "Belum terpetakan")
        journey_label = dominant_journey["stage_label"] if dominant_journey else "Belum terpetakan"
        if negative_share <= 0:
            return [
                ["Kondisi utama", f"Rating {rating_text}/5, belum ada sentimen negatif."],
                ["Lensa pengalaman", f"arah {score_metrics['direction']} untuk titik sentuh, rasa layanan, dan perjalanan agenda."],
                ["Area yang dijaga", top_risk_label],
                ["Keputusan cepat", f"Replikasi praktik baik pada {top_risk_label}."],
            ]
        return [
            ["Kondisi utama", f"Rating {rating_text}/5, sentimen negatif {negative_share}%."],
            ["Lensa pengalaman", f"arah {score_metrics['direction']} untuk titik sentuh, rasa layanan, dan perjalanan agenda."],
            ["Area yang diprioritaskan", top_risk_label],
            ["Keputusan cepat", f"Tetapkan owner dan intervensi 30 hari untuk {top_risk_label}."],
        ]

    def _executive_action_rows(self, top_risk, top_issue, dominant_journey, negative_share=0):
        risk_label = top_risk[0]["label"] if top_risk else "layanan dengan volume feedback terbesar"
        issue_label = top_issue["label"] if top_issue else "konsistensi kualitas layanan"
        journey_label = dominant_journey["stage_label"] if dominant_journey else "customer journey utama"
        if negative_share <= 0:
            return [
                ["1", risk_label, "Dokumentasikan praktik baik dan jadikan baseline layanan berikutnya.", "Kekuatan pengalaman pelanggan tidak berhenti sebagai pujian sesaat."],
                ["2", issue_label, "Tetap pantau indikator awal agar standar positif tidak turun pada periode berikutnya.", "Manajemen menjaga mutu tanpa membuat intervensi berlebihan."],
                ["3", journey_label, "Replikasi pola positif pada titik pengalaman lain yang masih belum sekuat area ini.", "Kualitas layanan meningkat lebih merata di seluruh perjalanan pelanggan."],
            ]
        return [
            ["1", risk_label, "Tetapkan owner lintas fungsi dan target perbaikan 30 hari.", "Risiko layanan tidak melebar menjadi isu reputasi."],
            ["2", issue_label, "Pisahkan keluhan yang insidental dari pola berulang, lalu buat quick win operasional.", "Manajemen tahu mana yang perlu dieksekusi dulu."],
            ["3", journey_label, "Perbaiki satu titik gesekan utama sebelum memperluas program perbaikan.", "Perubahan terasa di pengalaman pelanggan, bukan hanya di laporan."],
        ]

    def _specialist_review_markdown(self, timeframe, macro_trends, sentiment, segment, score_engine):
        return build_specialist_review_markdown(
            self,
            self._markdown_table,
            timeframe,
            macro_trends,
            sentiment,
            segment,
            score_engine,
        )

    @classmethod
    def _executive_section_synthesis(cls, report_sections, limit=4):
        rows = []
        for section in report_sections or []:
            title = str((section or {}).get("title") or "").strip()
            content = str((section or {}).get("content") or "").strip()
            brief = cls._section_brief(content)
            if title and brief:
                rows.append(f"- {title}: {brief}")
            if len(rows) >= limit:
                break
        if not rows:
            return ""
        return "\n".join(rows)

    def build_executive_snapshot(self, timeframe, notes="", sentiment="all", segment="all", score_engine=DEFAULT_SCORE_ENGINE, macro_trends="", report_sections=None, section_context=None):
        timeframe_df = self._filter_view(timeframe, sentiment=sentiment, segment=segment)
        if timeframe_df.empty:
            return "- Belum ada bukti evaluasi yang cukup untuk menyusun ringkasan eksekutif pada kombinasi filter yang dipilih.\n"

        context = self._build_analysis_context(timeframe_df, timeframe, sentiment, segment, score_engine)
        governance = self._governance_summary(timeframe_df)
        total_rows = governance["total_rows"]
        dimension_count = governance.get("dimension_count", len(timeframe_df))
        rating_response_count = governance.get("rating_response_count", 0)
        text_response_count = governance.get("text_response_count", 0)
        avg_rating = timeframe_df["Rating Numeric"].mean()
        sentiment_summary = self._sentiment_summary(timeframe_df)
        negative_share = round((sentiment_summary["issue_weight"] / max(dimension_count, 1)) * 100, 1)
        top_service = self._series_counts(timeframe_df["Layanan"], limit=1)
        top_stakeholder = self._series_counts(timeframe_df["Tipe Stakeholder"], limit=1)
        top_risk = self._group_risk(timeframe_df, "Layanan", limit=1)
        top_issue = next((theme for theme in self._theme_hits(timeframe_df) if theme["negative_hits"] > 0), None)
        section_context = section_context or {}
        focus_text = str(section_context.get("focus_note") or notes or "").strip() or "Tidak ada fokus tambahan dari pengguna."
        dominant_journey, score_metrics = context["dominant_journey"], context["score_metrics"]
        top_location = self._series_counts_for_column(timeframe_df, "Lokasi", limit=1)
        top_instructor_type = self._series_counts_for_column(timeframe_df, "Tipe Instruktur", limit=1)

        headlines = self._executive_headlines(total_rows, dimension_count, avg_rating, negative_share, top_risk, top_issue, dominant_journey, score_metrics, context)
        dashboard_table = self._markdown_table(
            ["Pertanyaan Eksekutif", "Jawaban Singkat"],
            self._executive_dashboard_rows(total_rows, dimension_count, rating_response_count, text_response_count, avg_rating, negative_share, top_risk, top_service, top_stakeholder, dominant_journey, score_metrics, context),
        )
        action_table = self._markdown_table(
            ["Prioritas", "Fokus", "Tindakan Manajemen", "Dampak yang Diharapkan"],
            self._executive_action_rows(top_risk, top_issue, dominant_journey, negative_share),
        )
        context_table = self._markdown_table(
            ["Konteks Pendukung", "Nilai"],
            [
                ["Periode", context.get("timeframe_label") or ("Seluruh Periode Evaluasi" if "apidog" in timeframe.lower() else timeframe)],
                ["Cakupan analisis", context["scope_text"]],
                ["Respons evaluasi dianalisis", f"{total_rows} respons"],
                ["Dimensi evaluasi diringkas", f"{dimension_count} dimensi"],
                ["Komposisi jawaban", f"{rating_response_count} rating; {text_response_count} komentar teks"],
                ["Lokasi pelatihan dominan", self._primary_label(top_location, "Belum terpetakan")],
                ["Tipe instruktur dominan", self._primary_label(top_instructor_type, "Belum terpetakan")],
                ["Kelengkapan field inti", f"{governance['completeness_pct']}%"],
            ],
        )
        positive_only = negative_share <= 0
        meeting_agenda = [
            f"- Bagaimana praktik baik pada {top_risk[0]['label']} dapat distandardisasi dan direplikasi pada 30 hari ke depan?" if positive_only and top_risk else f"- Apakah layanan {top_risk[0]['label']} memerlukan intervensi prioritas lintas fungsi pada 30 hari ke depan?" if top_risk else "- Apakah perusahaan perlu memperluas pengumpulan feedback agar risiko layanan lebih mudah dibaca?",
            f"- Bagaimana tindak lanjut yang paling tepat untuk tema {top_issue['label']} agar tidak berkembang menjadi keluhan berulang?" if top_issue else "- Kekuatan layanan mana yang paling layak distandardisasi dan direplikasi?",
            f"- Tahap customer journey mana yang paling perlu dijadikan contoh praktik baik, mengingat kekuatan utama saat ini berada pada {dominant_journey['stage_label']}?" if positive_only and dominant_journey else f"- Tahap customer journey mana yang paling perlu dikoreksi lebih dulu, mengingat titik gesekan terbesar saat ini berada pada {dominant_journey['stage_label']}?" if dominant_journey else "- Tahap customer journey mana yang paling perlu dipetakan lebih rinci pada periode berikutnya?",
            "- Apakah tata kelola kanal, penanggung jawab, dan tindak lanjut sudah cukup jelas untuk mendukung evaluasi periodik berikutnya?",
        ]
        technical_note = (
            "Detail formula, distribusi kanal, bukti verbatim, dan konteks eksternal ditempatkan setelah ringkasan ini agar eksekutif memperoleh inti keputusan lebih dulu."
        )
        decision_sentence = (
            f"Manajemen perlu mempertahankan dan mereplikasi kekuatan pengalaman pelanggan pada {top_risk[0]['label'] if top_risk else 'area layanan paling kuat'} dalam 30 hari ke depan karena area ini menjadi contoh praktik baik paling jelas pada periode ini."
            if positive_only
            else f"Manajemen perlu memprioritaskan {top_risk[0]['label'] if top_risk else 'area layanan paling terekspos'} dalam 30 hari ke depan karena area ini menjadi sinyal risiko pengalaman pelanggan paling jelas pada periode ini."
        )
        recommendation_first = (
            f"- Tetapkan penanggung jawab lintas fungsi untuk menjaga standar {top_risk[0]['label'] if top_risk else 'prioritas layanan utama'} dan dokumentasikan praktik yang layak direplikasi."
            if positive_only
            else f"- Tetapkan penanggung jawab lintas fungsi untuk {top_risk[0]['label'] if top_risk else 'prioritas layanan utama'} dan minta rencana aksi 30 hari dengan indikator keberhasilan yang terukur."
        )
        risk_label = top_risk[0]["label"] if top_risk else self._primary_label(top_service, "area layanan utama")
        issue_label = top_issue["label"] if top_issue else "konsistensi kualitas layanan"
        journey_label = dominant_journey["stage_label"] if dominant_journey else "customer journey utama"
        conclusion_sentence = (
            f"Jawaban utama laporan ini: manajemen sebaiknya mempertahankan dan mereplikasi standar {risk_label} karena area tersebut menjadi bukti paling kuat bahwa pengalaman pelanggan dapat distandardisasi lintas agenda."
            if positive_only
            else f"Jawaban utama laporan ini: manajemen sebaiknya memprioritaskan {risk_label} sekarang karena area tersebut menjadi titik paling jelas yang dapat memengaruhi pengalaman pelanggan, tindak lanjut, dan kepercayaan pada agenda perusahaan."
        )
        reason_lines = [
            f"- **Dampak:** {risk_label} menentukan konsistensi layanan.",
            f"- **Konsentrasi:** Isu {issue_label} membatasi ruang lingkup perbaikan.",
            f"- **Perjalanan:** {journey_label} adalah prioritas pembahasan.",
            f"- **Lensa:** Tren {score_metrics['direction']} sebagai peringatan dini.",
        ]
        evidence_lines = [
            f"- Meringkas {total_rows} respons ({dimension_count} dimensi) pada {context['scope_text']}.",
            f"- Komposisi: {rating_response_count} rating dan {text_response_count} komentar (kelengkapan {governance['completeness_pct']}%).",
            f"- Lensa gabungan Experience Index untuk titik sentuh yang dirasakan sepanjang perjalanan.",
            *headlines,
        ]
        decision_lines = [
            recommendation_first,
            f"- Sepakati indikator keberhasilan, target 30 hari, dan ritme tinjauan mingguan untuk {risk_label}.",
            f"- Gunakan {journey_label} sebagai titik awal agar keputusan terasa langsung pada pengalaman pelanggan.",
        ]
        confidence_lines = [
            f"- Proyeksi dipakai sebagai peringatan dini deterministik berbasis penilaian, sentimen, tema risiko, perjalanan pelanggan, dan lensa pengalaman.",
            f"- Klaim belum diposisikan sebagai backtesting statistik; periode berikutnya harus membandingkan arah proyeksi dengan skor aktual.",
            f"- Catatan pengguna: {focus_text}",
            f"- {technical_note}",
        ]
        management_interpretation = FeedbackManagementInterpreter.build(
            avg_rating=avg_rating,
            negative_share=negative_share,
            top_risk_label=top_risk[0]["label"] if top_risk else self._primary_label(top_service, "area layanan utama"),
            issue_label=top_issue["label"] if top_issue else "konsistensi kualitas layanan",
            journey_label=dominant_journey["stage_label"] if dominant_journey else "customer journey utama",
            score_label=context["score_profile"]["label"],
            current_score=score_metrics["current_score"],
            projected_score=score_metrics["projected_score"],
        )
        interpretation_table = FeedbackManagementInterpreter.to_markdown_table(management_interpretation)

        forecast_30d = context.get("forecast_30d", {})
        component_table = self._markdown_table(
            ["Lensa Skor", "Bobot", "Saat Ini", "Proyeksi 30 Hari", "Arah"],
            [
                [
                    item.get("label", "-"),
                    f"{item.get('weight_pct', 0)}%",
                    item.get("current_score", 0),
                    item.get("projected_score", 0),
                    item.get("delta", 0),
                ]
                for item in forecast_30d.get("component_rows", [])
            ],
        )
        weekly_table = self._markdown_table(
            ["Minggu", "Proyeksi Skor", "Pola", "Pembacaan"],
            [
                [
                    item.get("week", "-"),
                    item.get("score", 0),
                    item.get("pattern", "-"),
                    item.get("reading", "-"),
                ]
                for item in forecast_30d.get("weekly_rows", [])
            ],
        )
        dashboard_visual = [
            "### Dasbor Visual Mingguan",
            "[[CHART: Komponen Pembentuk Experience Index | Skor Saat Ini | "
            f"{forecast_30d.get('score_chart', '')}]]" if forecast_30d.get("score_chart") else "",
            "",
            component_table,
            "",
            "### Proyeksi Mingguan",
            "[[LINE: Proyeksi Mingguan 30 Hari | Skor | "
            f"{forecast_30d.get('weekly_chart', '')}]]" if forecast_30d.get("weekly_chart") else "",
            "",
            weekly_table,
            "",
            f"Catatan pembacaan: {forecast_30d.get('source_note', 'Pola mingguan dibaca sebagai peringatan dini.')} "
            f"Tingkat keyakinan: {forecast_30d.get('confidence', 'rendah')}.",
            "",
        ]

        section_synthesis = self._executive_section_synthesis(report_sections)
        synthesis_block = ["### Sintesis Lintas Bab", section_synthesis, ""] if section_synthesis else []
        return DocumentBuilder.reader_facing_text("\n".join([
            "### Kesimpulan Utama",
            conclusion_sentence,
            "",
            "### Keputusan yang Diminta", *decision_lines, "",
            "### Alasan Utama", *reason_lines, "",
            *dashboard_visual,
            "### Bukti Pendukung", *evidence_lines, "",
            "### Interpretasi Manajemen", interpretation_table, "",
            "### Dasbor Keputusan", dashboard_table, "",
            "### Matriks Keputusan", action_table, "",
            "### Catatan Keyakinan dan Batasan", *confidence_lines, "",
            "### Agenda Follow-up Meeting", *meeting_agenda, "",
            *synthesis_block,
            "### Konteks Pendukung", context_table,
        ]))

import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import pandas as pd
from docx import Document
from docx.shared import Twips

from config import CX_SENTIMENT_STRUCTURE
from document_builder import DocumentBuilder
from report_evidence import ReportEvidenceBuilder
from report_analytics import FeedbackAnalyticsEngine
from report_agents import FeedbackProposalTeam
from report_quality import ReportQualityValidator


class ReportAnalyticsContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dataframe = pd.read_csv(PROJECT_DIR / "data" / "db.csv")
        cls.engine = FeedbackAnalyticsEngine(cls.dataframe)
        cls.timeframe = "1 Bulan Terakhir (Monthly)"
        cls.notes = "Periksa risiko jadwal dan tindak lanjut layanan."
        cls.macro_trends = (
            "**Insight Mendalam (via example.com):** Benchmark pelatihan IT menekankan follow-up pasca kelas.\n\n"
            "1. Tren pelatihan korporat Indonesia | Kebutuhan evaluasi dampak meningkat | sumber=example.com | tanggal=2026"
        )

    def test_report_sections_match_configured_structure(self):
        sections = self.engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
        )

        self.assertEqual([item["id"] for item in sections], [item["id"] for item in CX_SENTIMENT_STRUCTURE])
        self.assertEqual([item["title"] for item in sections], [item["title"] for item in CX_SENTIMENT_STRUCTURE])
        self.assertTrue(all(section["content"].strip() for section in sections))
        self.assertTrue(all(section["content"].lstrip().startswith("### Bukti yang Dipakai") for section in sections))

        combined = "\n".join(section["content"] for section in sections)
        required_markers = [
            "## 1.1 Ringkasan Cakupan Umpan Balik dan Tata Kelola",
            "## 2.1 Akar Masalah Utama dan Titik Keluhan Dominan",
            "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut",
            "## 4.1 Intervensi Prioritas 30 Hari",
            "## 5.1 Prioritas Sasaran Bisnis",
            "Indeks Pengalaman",
        ]
        for marker in required_markers:
            self.assertIn(marker, combined)

    def test_report_evidence_cards_are_source_safe(self):
        sections = [
            {
                "id": "cx_chap_1",
                "title": "Analisis Deskriptif",
                "content": "APIDog source=/api/Resource/dataset menunjukkan 120 respons dan risiko onboarding.",
            }
        ]

        enriched = ReportEvidenceBuilder.attach_to_sections(sections)
        content = enriched[0]["content"]

        self.assertIn("### Bukti yang Dipakai", content)
        self.assertIn("120 respons", content)
        for forbidden in ["APIDog", "/api/Resource/dataset", "source=", "Evidence Ledger"]:
            self.assertNotIn(forbidden, content)

    def test_narrative_preflight_rejects_empty_and_raw_source_content(self):
        result = ReportQualityValidator.evaluate_narrative(
            "## Ringkasan Eksekutif\nRingkas.",
            [{"id": "cx_chap_1", "title": "Bab 1", "content": "APIDog endpoint."}],
        )

        self.assertFalse(result["passes"])
        self.assertIn("raw_source_label", result["categories"])

    def test_executive_snapshot_keeps_decision_ready_contract(self):
        snapshot = self.engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
        )

        required_markers = [
            "## Ringkasan Eksekutif",
            "### Inti Keputusan",
            "### Temuan Utama",
            "### Rekomendasi",
            "### Dasbor Keputusan",
            "| Pertanyaan Eksekutif | Jawaban Singkat |",
            "### Keputusan yang Perlu Diambil",
            "### Agenda Diskusi",
        ]
        for marker in required_markers:
            self.assertIn(marker, snapshot)
        self.assertNotIn("Formula Experience Index", snapshot)
        self.assertLess(snapshot.index("### Inti Keputusan"), snapshot.index("### Temuan Utama"))
        self.assertLess(snapshot.index("### Temuan Utama"), snapshot.index("### Rekomendasi"))
        self.assertLess(snapshot.index("### Rekomendasi"), snapshot.index("### Dasbor Keputusan"))
        self.assertLess(snapshot.index("### Dasbor Keputusan"), snapshot.index("### Keputusan yang Perlu Diambil"))

        sections = self.engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
        )
        combined_sections = "\n".join(section["content"] for section in sections)
        self.assertIn("Penjelasan Perhitungan Indeks Pengalaman", combined_sections)

    def test_executive_snapshot_reports_raw_response_volume_for_aggregated_class_report(self):
        dataframe = pd.DataFrame(
            [
                {
                    "Record ID": "class_report-00001",
                    "Sumber Feedback": "class_report",
                    "Kanal Feedback": "Evaluasi Kelas Internal",
                    "Tanggal Feedback": "Tanggal tidak tersedia",
                    "Tipe Stakeholder": "Peserta Kelas",
                    "Layanan": "Materi dan kurikulum",
                    "Lokasi": "",
                    "Tipe Instruktur": "",
                    "Rentang Waktu": "Semua Data APIDog (tanggal tidak tersedia)",
                    "Rating": "3",
                    "Komentar": "Rata-rata rating Kesesuaian materi bahan ajar: 3.0 dari 5. Mengapa: Materi perlu contoh praktik tambahan.",
                    "Customer Journey Hint": "Pelaksanaan Layanan",
                    "Raw Response Count": "3",
                    "Rating Response Count": "2",
                    "Text Response Count": "1",
                    "Rating Distribution": "2: 1; 4: 1",
                    "Representative Why": "Materi perlu contoh praktik tambahan.",
                }
            ]
        )
        engine = FeedbackAnalyticsEngine(dataframe)

        snapshot = engine.build_executive_snapshot(
            "Semua Data APIDog (tanggal tidak tersedia)",
            score_engine="experience_index",
        )

        self.assertIn("3 respons", snapshot)
        self.assertIn("1 dimensi", snapshot)
        self.assertIn("2 penilaian", snapshot)
        self.assertIn("1 komentar teks", snapshot)

    def test_docx_quality_accepts_generated_contract(self):
        snapshot = self.engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
        )
        sections = self.engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
        )

        document = Document()
        DocumentBuilder.process_content(document, snapshot)
        for section in sections:
            document.add_heading(section["title"], level=1)
            DocumentBuilder.process_content(document, section["content"])

        quality = ReportQualityValidator.evaluate(document, snapshot, sections, "Experience Index")
        self.assertTrue(quality["verified_complete"], quality)
        self.assertEqual(quality["passed_checks"], quality["total_checks"])

    def test_docx_ordered_lists_restart_and_use_hanging_indents(self):
        document = Document()
        DocumentBuilder.process_content(
            document,
            "1. Tahap pertama\n2. Tahap kedua\n\nParagraf pemisah.\n\n1. Siklus baru\n2. Nomor kedua",
        )

        numbered_paragraphs = [paragraph for paragraph in document.paragraphs if paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None]
        self.assertEqual(len(numbered_paragraphs), 4)

        first_list_num_id = numbered_paragraphs[0]._p.pPr.numPr.numId.val
        second_list_num_id = numbered_paragraphs[2]._p.pPr.numPr.numId.val
        self.assertNotEqual(first_list_num_id, second_list_num_id)

        for paragraph in numbered_paragraphs:
            paragraph_format = paragraph.paragraph_format
            self.assertEqual(paragraph_format.left_indent, Twips(720))
            self.assertEqual(paragraph_format.first_line_indent, Twips(-360))

    def test_predictive_section_links_osint_to_company_current_and_future_state(self):
        sections = self.engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
        )
        predictive = next(section["content"] for section in sections if section["id"] == "cx_chap_3")

        self.assertIn("## 3.4 Keterkaitan Faktor Eksternal dengan Kondisi Perusahaan", predictive)
        self.assertIn("kondisi perusahaan saat ini", predictive)
        self.assertIn("kondisi perusahaan ke depan", predictive)

    def test_specialist_agents_generate_bounded_briefings_from_internal_datasets(self):
        briefing = FeedbackProposalTeam().run(
            self.engine,
            self.dataframe,
            self.timeframe,
            self.macro_trends,
        )

        self.assertEqual(
            [item["role"] for item in briefing["specialists"]],
            [
                "Data Steward",
                "Rating Analyst",
                "Voice-of-Customer Analyst",
                "External Context Analyst",
                "Action Planner",
            ],
        )
        self.assertTrue(all(item["finding"].strip() for item in briefing["specialists"]))
        self.assertIn("rating", briefing["sources_used"])
        self.assertIn("komentar", briefing["sources_used"])
        self.assertIn("osint", briefing["sources_used"])
        self.assertLessEqual(len(briefing["manager_summary"]), 420)

    def test_specialist_briefing_exposes_evidence_confidence_and_qa_guardrails(self):
        briefing = FeedbackProposalTeam().run(
            self.engine,
            self.dataframe,
            self.timeframe,
            self.macro_trends,
        )

        self.assertIn("evidence_ledger", briefing)
        self.assertIn("qa_review", briefing)
        self.assertTrue(briefing["evidence_ledger"])
        self.assertIn(briefing["confidence"], {"Tinggi", "Sedang", "Rendah"})
        self.assertTrue(all(item["evidence_type"] for item in briefing["evidence_ledger"]))
        self.assertTrue(all(item["source"] for item in briefing["evidence_ledger"]))
        self.assertTrue(all(item["confidence"] in {"Tinggi", "Sedang", "Rendah"} for item in briefing["specialists"]))
        self.assertTrue(all("temuan evaluasi" in item for item in briefing["qa_review"]))

    def test_specialist_briefing_adds_audit_trend_contradiction_and_prediction_controls(self):
        briefing = FeedbackProposalTeam().run(
            self.engine,
            self.dataframe,
            self.timeframe,
            self.macro_trends,
            score_engine="experience_index",
        )

        self.assertIn("audit_trail", briefing)
        self.assertIn("contradiction_review", briefing)
        self.assertIn("trend_review", briefing)
        self.assertIn("prediction_review", briefing)
        self.assertEqual(briefing["audit_trail"]["timeframe"], self.timeframe)
        self.assertEqual(briefing["audit_trail"]["score_engine"], "experience_index")
        self.assertGreater(briefing["audit_trail"]["raw_response_count"], 0)
        self.assertIn("rating_text_alignment", briefing["contradiction_review"])
        self.assertIn(briefing["contradiction_review"]["severity"], {"Rendah", "Sedang", "Tinggi"})
        self.assertIn("comparison_period", briefing["trend_review"])
        self.assertIn("rating_delta", briefing["trend_review"])
        self.assertFalse(briefing["prediction_review"]["statistical_forecast"])
        self.assertIn("early warning", briefing["prediction_review"]["method"].lower())

    def test_executive_summary_hides_specialist_workflow_labels(self):
        snapshot = self.engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
        )

        self.assertIn("### Inti Keputusan", snapshot)
        self.assertIn("### Temuan Utama", snapshot)
        self.assertIn("### Rekomendasi", snapshot)
        for forbidden in [
            "Review Tim Analis Internal",
            "Data Steward",
            "Rating Analyst",
            "Voice-of-Customer Analyst",
            "Confidence Desk",
            "Evidence Ledger",
            "QA Guardrail",
            "Report Audit Trail",
            "Contradiction Check",
            "Historical Trend Desk",
            "Prediction Boundary",
        ]:
            self.assertNotIn(forbidden, snapshot)


if __name__ == "__main__":
    unittest.main()

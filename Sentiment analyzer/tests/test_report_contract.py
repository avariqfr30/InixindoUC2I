import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import pandas as pd
from docx import Document

from config import CX_SENTIMENT_STRUCTURE
from document_builder import DocumentBuilder
from report_analytics import FeedbackAnalyticsEngine
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

        combined = "\n".join(section["content"] for section in sections)
        required_markers = [
            "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola",
            "## 2.1 Akar Masalah Utama dan Pain Point Dominan",
            "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut",
            "## 4.1 Intervensi Prioritas 30 Hari",
            "## 5.1 Prioritas Sasaran Bisnis",
            "Experience Index",
        ]
        for marker in required_markers:
            self.assertIn(marker, combined)

    def test_executive_snapshot_keeps_decision_ready_contract(self):
        snapshot = self.engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
        )

        required_markers = [
            "## Ringkasan Eksekutif",
            "| Indikator Kunci | Nilai |",
            "Formula Experience Index",
            "### Agenda Diskusi Prioritas",
            "### Poin Utama untuk Pembacaan Cepat",
        ]
        for marker in required_markers:
            self.assertIn(marker, snapshot)

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


if __name__ == "__main__":
    unittest.main()

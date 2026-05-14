import sys
import unittest
from pathlib import Path

from docx import Document

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from document_builder import DocumentBuilder
from report_analytics import FeedbackAnalyticsEngine


class ReaderFacingDocumentContractTests(unittest.TestCase):
    def test_reader_facing_text_suppresses_internal_source_disclosures(self):
        raw = (
            "Semua Data APIDog (tanggal tidak tersedia) memakai Internal API / feedback cache. "
            "Jejak sumber yang dipakai: internal_api, komentar, osint, rating. "
            "internal facts remain separated from OSINT benchmark context."
        )

        clean = DocumentBuilder.reader_facing_text(raw)

        forbidden = ["APIDog", "Internal API", "internal_api", "internal facts", "feedback cache"]
        for token in forbidden:
            self.assertNotIn(token, clean)
        self.assertIn("Seluruh Periode Evaluasi", clean)
        self.assertIn("Jejak bukti yang dipakai", clean)

    def test_generated_cover_has_static_toc_and_neutral_period_label(self):
        document = Document()

        DocumentBuilder.create_cover(document, "Semua Data APIDog (tanggal tidak tersedia)")
        document.add_heading("EXECUTIVE SUMMARY", level=1)
        document.add_heading("BAB I - Detail", level=1)

        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("DAFTAR ISI", text)
        self.assertIn("EXECUTIVE SUMMARY", text)
        self.assertIn("BAB I", text)
        self.assertIn("Seluruh Periode Evaluasi", text)
        self.assertNotIn("APIDog", text)
        self.assertNotIn("Perbarui field", text)

    def test_executive_summary_is_business_first_and_source_neutral(self):
        data = [
            {
                "Rentang Waktu": "Semua Data APIDog (tanggal tidak tersedia)",
                "Rating Numeric": 4,
                "Sentiment Label": "positive",
                "Layanan": "Kinerja instruktur",
                "Tipe Stakeholder": "Peserta Kelas",
                "Sumber Feedback": "Internal API",
                "Kanal Feedback": "Evaluasi Kelas Internal",
                "Komentar": "Instruktur jelas tetapi waktu praktik masih kurang.",
                "Customer Journey Stage": "Pelaksanaan Layanan",
            }
        ]
        engine = FeedbackAnalyticsEngine.from_records(data)

        summary = engine.build_executive_snapshot("Semua Data APIDog (tanggal tidak tersedia)")

        self.assertIn("## Executive Summary", summary)
        self.assertIn("### What leadership needs to know", summary)
        self.assertIn("### Decisions to make", summary)
        for token in ["APIDog", "Internal API", "internal_api", "internal facts", "Dataset Spesialis"]:
            self.assertNotIn(token, summary)


if __name__ == "__main__":
    unittest.main()

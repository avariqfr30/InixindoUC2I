import unittest

from feedback_deliberation import FeedbackDeliberationBuilder
from report_planning import FeedbackSectionPlanner
from report_quality import ReportQualityValidator


class FeedbackDeliberationTests(unittest.TestCase):
    def setUp(self):
        FeedbackDeliberationBuilder.clear_cache()
        FeedbackSectionPlanner.clear_cache()
        self.sections = [
            {"id": "cx_chap_1", "title": "Analisis Deskriptif", "content": "Pola rating menunjukkan pengalaman yang perlu dibaca lebih lanjut."},
            {"id": "cx_chap_2", "title": "Analisis Diagnostik", "content": "Komentar peserta mengarah pada hambatan fasilitas."},
            {"id": "cx_chap_4", "title": "Rekomendasi Preskriptif", "content": "Tindakan memiliki penanggung jawab dan target layanan."},
        ]
        self.context = {
            "timeframe_label": "3 Bulan Terakhir",
            "sentiment": "all",
            "segment": "all",
            "row_count": 939,
            "text_response_count": 259,
            "external_context_ready": False,
            "insight_cards": [
                {
                    "observation": "939 respons evaluasi terolah pada cakupan aktif.",
                    "implication": "Temuan perlu diperlakukan sebagai pola pengalaman.",
                    "confidence": "high",
                },
                {
                    "observation": "259 respons memiliki komentar teks.",
                    "implication": "Komentar membantu menjelaskan pola tanpa membuktikan sebab tunggal.",
                    "confidence": "medium",
                },
            ],
        }

    def test_builds_feedback_research_chapter_and_editorial_contracts(self):
        builder = FeedbackDeliberationBuilder()
        contract = builder.build(self.sections, self.context, data_version="feedback-v1")

        self.assertEqual(
            {
                "cache_key", "data_version", "evidence_dossier", "research_plan",
                "document_thesis", "chapter_contracts", "claim_ledger",
                "data_gap_register", "editorial_contract", "appendix_manifest",
            },
            set(contract),
        )
        self.assertEqual("cx_chap_1", contract["chapter_contracts"][1]["depends_on"])
        self.assertTrue(any("segmen" in item["question"].lower() for item in contract["research_plan"]["questions"]))
        self.assertTrue(any("eksternal" in item["gap"].lower() for item in contract["data_gap_register"]))
        self.assertIn("meaning_lock", contract["editorial_contract"])
        self.assertEqual(contract, builder.build(self.sections, self.context, data_version="feedback-v1"))
        self.assertEqual(1, builder.cache_stats()["hits"])

    def test_builds_methodology_measurement_and_gap_appendices(self):
        builder = FeedbackDeliberationBuilder()
        contract = builder.build(self.sections, self.context, data_version="feedback-v1")
        appendix = builder.build_appendix_markdown(contract)

        self.assertIn("# Lampiran Metodologi, Pengukuran, dan Kesenjangan Data", appendix)
        self.assertIn("## A. Cakupan dan Metodologi", appendix)
        self.assertIn("## B. Matriks Temuan dan Pengukuran", appendix)
        self.assertIn("## C. Kesenjangan Data", appendix)
        self.assertNotIn("ClassReport", appendix)
        self.assertNotIn("SECTION_PLAN_JSON", appendix)

    def test_planner_and_quality_gate_require_the_same_contract(self):
        builder = FeedbackDeliberationBuilder()
        contract = builder.build(self.sections, self.context, data_version="feedback-v1")
        plan = FeedbackSectionPlanner().build_plan(
            [section["title"] for section in self.sections],
            {**self.context, "document_contract": contract},
        )
        appendix = builder.build_appendix_markdown(contract)
        accepted = ReportQualityValidator.evaluate_narrative(
            "Ringkasan eksekutif yang cukup substantif untuk menjelaskan pola pengalaman, bukti, keputusan, dan batas keyakinan manajemen secara jelas.",
            self.sections,
            deliberation_contract=contract,
            appendix_content=appendix,
        )
        rejected = ReportQualityValidator.evaluate_narrative(
            "Ringkasan eksekutif yang cukup substantif untuk menjelaskan pola pengalaman, bukti, keputusan, dan batas keyakinan manajemen secara jelas.",
            self.sections,
            deliberation_contract=contract,
            appendix_content="",
        )

        self.assertEqual(contract["document_thesis"], plan["document_thesis"])
        self.assertIn("missing_tiered_appendix", rejected["categories"])
        self.assertNotIn("missing_tiered_appendix", accepted["categories"])


if __name__ == "__main__":
    unittest.main()

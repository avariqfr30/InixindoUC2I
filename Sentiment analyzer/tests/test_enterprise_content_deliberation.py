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
            "trust_packet": {
                "version": "feedback-fact-registry-v1",
                "snapshot_fingerprint": "sha256:0123456789abcdef",
                "confidence_basis": {
                    "level": "Tinggi",
                    "response_count": 939,
                    "rating_coverage_pct": 98.0,
                    "comment_coverage_pct": 27.6,
                    "field_completeness_pct": 96.5,
                    "limitations": ["Pembanding eksternal belum cukup kuat."],
                },
                "contradiction_review": {
                    "severity": "Rendah",
                    "rating_text_alignment": "rating dan komentar relatif sejalan",
                },
                "theme_evidence_ids": {"material": "F-TEMA-MATERI"},
                "facts": [],
            },
        }

    def test_builds_feedback_research_chapter_and_editorial_contracts(self):
        builder = FeedbackDeliberationBuilder()
        contract = builder.build(self.sections, self.context, data_version="feedback-v1")

        self.assertEqual(
            {
                "cache_key", "data_version", "evidence_dossier", "research_plan",
                "document_thesis", "chapter_contracts", "claim_ledger",
                "data_gap_register", "editorial_contract", "appendix_manifest", "trust_packet",
            },
            set(contract),
        )
        self.assertEqual("cx_chap_1", contract["chapter_contracts"][1]["depends_on"])
        self.assertTrue(any("segmen" in item["question"].lower() for item in contract["research_plan"]["questions"]))
        self.assertTrue(any("eksternal" in item["gap"].lower() for item in contract["data_gap_register"]))
        self.assertIn("meaning_lock", contract["editorial_contract"])
        self.assertEqual("sha256:0123456789abcdef", contract["trust_packet"]["snapshot_fingerprint"])
        self.assertEqual("Rendah", contract["trust_packet"]["contradiction_review"]["severity"])
        self.assertEqual(contract, builder.build(self.sections, self.context, data_version="feedback-v1"))
        self.assertEqual(1, builder.cache_stats()["hits"])

    def test_embeds_distilled_uc2_exemplar_profile_without_source_identity_leakage(self):
        builder = FeedbackDeliberationBuilder()
        contract = builder.build(self.sections, self.context, data_version="feedback-v1")
        profile = contract["editorial_contract"]["exemplar_profile"]

        self.assertEqual("uc2-feedback-exemplar-profile-v1", profile["version"])
        self.assertEqual("use_existing_report_structure_only", profile["hardcoded_structure_policy"])
        self.assertIn("theme_bank", profile["analysis_moves"])
        self.assertIn("sentiment_volume_movement", profile["analysis_moves"])
        self.assertIn("tercatat", " ".join(profile["indonesian_language_rules"]).lower())
        self.assertTrue(any("bukan bukti" in rule.lower() for rule in profile["factual_boundaries"]))

        serialized = repr(profile).lower()
        for leaked_name in ["redquadrant", "wordnerds", "bury", "exeter", "ecc", "sk-februari", "c.1.1"]:
            self.assertNotIn(leaked_name, serialized)

    def test_builds_methodology_measurement_and_gap_appendices(self):
        builder = FeedbackDeliberationBuilder()
        contract = builder.build(self.sections, self.context, data_version="feedback-v1")
        appendix = builder.build_appendix_markdown(contract)

        self.assertIn("# Lampiran Metodologi, Pengukuran, dan Kesenjangan Data", appendix)
        self.assertIn("## A. Cakupan dan Metodologi", appendix)
        self.assertIn("## B. Matriks Temuan dan Pengukuran", appendix)
        self.assertIn("## C. Kesenjangan Data", appendix)
        self.assertIn("sha256:0123456789abcdef", appendix)
        self.assertIn("Basis keyakinan", appendix)
        self.assertIn("98.0%", appendix)
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

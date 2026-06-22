import unittest
import json
from pathlib import Path

import pandas as pd


class FeedbackEditorialQualityTests(unittest.TestCase):
    def test_golden_fixture_contains_human_and_deterministic_closure_gates(self):
        fixture = json.loads((Path(__file__).parent / "fixtures" / "golden_feedback_quality.json").read_text())

        self.assertTrue(all(score == 4 for score in fixture["expected"]["human_rubric_minimums"].values()))
        self.assertEqual(fixture["expected"]["max_repeated_nontrivial_cell_count"], 2)
        self.assertLessEqual(fixture["expected"]["max_table_cell_words"], 18)

    def test_policy_excludes_irrelevant_datasets_and_centres_participant_experience(self):
        from editorial_intelligence import EXCLUDED_DATASETS, feedback_voice_rules

        self.assertEqual(EXCLUDED_DATASETS, {"FinanceInvoice", "ProjectStandards"})
        rules = " ".join(feedback_voice_rules()).lower()
        self.assertIn("suara peserta", rules)
        self.assertIn("tindak lanjut", rules)

    def test_repeated_table_detail_is_omitted_without_another_generic_phrase(self):
        from editorial_intelligence import compact_feedback_table_rows

        rows = [["Jadwal", "Perlu mitigasi agar pengalaman tidak melemah"] for _ in range(3)]
        compacted = compact_feedback_table_rows(rows)

        self.assertEqual(compacted[2][1], "")
        joined = " ".join(cell for row in compacted for cell in row).lower()
        self.assertNotIn("sinyal serupa", joined)

    def test_issue_story_action_keeps_fact_interpretation_and_action_distinct(self):
        from editorial_intelligence import build_issue_story_action

        result = build_issue_story_action(
            signal="12 peserta menyebut perubahan jadwal mendadak.",
            meaning="Peserta kesulitan menyesuaikan agenda kerja.",
            action="Konfirmasi jadwal minimal tiga hari sebelum kelas.",
        )

        self.assertEqual(result["participant_signal"], "12 peserta menyebut perubahan jadwal mendadak.")
        self.assertEqual(result["service_meaning"], "Peserta kesulitan menyesuaikan agenda kerja.")
        self.assertEqual(result["next_action"], "Konfirmasi jadwal minimal tiga hari sebelum kelas.")

    def test_style_assessment_flags_metric_led_repetition(self):
        from editorial_intelligence import assess_feedback_style

        result = assess_feedback_style(
            "Experience Index menguat pada layanan A.\n\n"
            "Experience Index menguat pada layanan B.\n\n"
            "Experience Index menguat pada layanan C."
        )

        self.assertFalse(result["passed"])
        self.assertIn("metric_led_repetition", result["findings"])

    def test_protected_editor_uses_feedback_specific_style_findings(self):
        from writing_quality import ProtectedIndonesianEditor

        issues = ProtectedIndonesianEditor.local_template_issues(
            "Experience Index menguat pada layanan A.\n\n"
            "Experience Index menguat pada layanan B.\n\n"
            "Experience Index menguat pada layanan C."
        )

        self.assertIn("metric_led_repetition", issues)

    def test_protected_editor_rejects_polish_that_changes_numeric_facts(self):
        from writing_quality import ProtectedIndonesianEditor

        class FakeModelClient:
            def chat(self, **kwargs):
                return {"message": {"content": "Rating tercatat 4,8/5 dari 20 respons."}}

        calls = {"count": 0}

        def quality_check(_text, _protected):
            calls["count"] += 1
            if calls["count"] == 1:
                return {"issues": ["long_sentence"], "protected_missing": []}
            return {"issues": [], "protected_missing": []}

        editor = ProtectedIndonesianEditor(
            model_client=FakeModelClient(),
            quality_fn=quality_check,
        )
        original = "Rating tercatat 4,5/5 dari 20 respons."

        self.assertEqual(original, editor.polish(original))

    def test_full_cached_range_and_strict_iso_dates(self):
        from timeframe_filters import FULL_CACHED_TIMEFRAME, filter_by_timeframe, parse_custom_timeframe

        dataframe = pd.DataFrame(
            [
                {"Record ID": "A", "Tanggal Feedback": "2026-06-12"},
                {"Record ID": "B", "Tanggal Feedback": "2023-10-31T09:15:00"},
                {"Record ID": "C", "Tanggal Feedback": "31/10/2023"},
            ]
        )

        scoped = filter_by_timeframe(dataframe, FULL_CACHED_TIMEFRAME)

        self.assertEqual(set(scoped["Record ID"]), {"A", "B"})
        self.assertIsNone(parse_custom_timeframe("custom_range:01/06/2026..12/06/2026"))

    def test_osint_cards_reject_unrelated_market_news(self):
        from osint_research import Researcher

        cards = Researcher._extract_brief_cards(
            "1. Harga batu bara naik tajam pada perdagangan Asia. (Sumber: example.com, 2026)\n"
            "2. Peserta pelatihan mengharapkan tindak lanjut dan materi yang relevan. (Sumber: training.example, 2026)",
            context={"segment": "Peserta", "score_engine": "Experience Index"},
        )

        self.assertEqual(len(cards), 1)
        self.assertIn("Peserta pelatihan", cards[0]["claim"])
        self.assertGreaterEqual(cards[0]["relevance_score"], 1)


if __name__ == "__main__":
    unittest.main()

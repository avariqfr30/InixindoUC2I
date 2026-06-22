import sys
import unittest
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from report_factuality import NarrativeFactValidator, ReportFactRegistry


class ReportFactRegistryTests(unittest.TestCase):
    def setUp(self):
        self.dataframe = pd.DataFrame(
            [
                {
                    "Record ID": "A-001",
                    "Layanan": "Materi",
                    "Tipe Stakeholder": "Peserta",
                    "Rating Numeric": 4.0,
                    "Komentar": "Materi perlu contoh praktik tambahan.",
                    "Sentiment Label": "mixed",
                },
                {
                    "Record ID": "A-002",
                    "Layanan": "Instruktur",
                    "Tipe Stakeholder": "Peserta",
                    "Rating Numeric": 5.0,
                    "Komentar": "Instruktur jelas.",
                    "Sentiment Label": "positive",
                },
            ]
        )
        self.analysis_context = {
            "score_profile": {"label": "Experience Index"},
            "score_metrics": {
                "current_score": 82.5,
                "projected_score": 81.2,
                "delta": -1.3,
                "direction": "turun",
                "theme_rows": [
                    {
                        "theme_id": "material",
                        "label": "Materi dan relevansi konten",
                        "total_hits": 1,
                        "negative_hits": 1,
                        "positive_hits": 0,
                        "priority_score": 51.0,
                    }
                ],
            },
            "dominant_journey": {"stage_label": "Pelaksanaan Layanan"},
        }
        self.governance = {
            "total_rows": 2,
            "dimension_count": 2,
            "rating_response_count": 2,
            "text_response_count": 2,
            "completeness_pct": 100.0,
            "source_count": 1,
            "channel_count": 1,
        }
        self.contradiction = {
            "rating_text_alignment": "rating dan komentar relatif sejalan",
            "severity": "Rendah",
            "average_rating": 4.5,
            "negative_rating_share": 0.0,
            "negative_text_hits": 1,
            "positive_text_hits": 1,
        }
        self.scope = {
            "timeframe": "2026-06",
            "timeframe_label": "Juni 2026",
            "sentiment": "all",
            "segment": "all",
            "score_engine": "experience_index",
            "external_context_ready": False,
        }

    def _build(self, dataframe=None):
        return ReportFactRegistry.build(
            dataframe if dataframe is not None else self.dataframe,
            self.analysis_context,
            self.governance,
            self.contradiction,
            self.scope,
        )

    def test_builds_stable_filter_scoped_fact_registry(self):
        first = self._build()
        second = self._build(self.dataframe.iloc[::-1].reset_index(drop=True))

        self.assertEqual(first, second)
        self.assertEqual("feedback-fact-registry-v1", first["version"])
        self.assertRegex(first["snapshot_fingerprint"], r"^sha256:[0-9a-f]{16}$")
        facts = {item["fact_id"]: item for item in first["facts"]}
        self.assertEqual(2, facts["F-RESPONSES"]["value"])
        self.assertEqual(2, facts["F-RATINGS"]["value"])
        self.assertEqual(82.5, facts["F-SCORE-CURRENT"]["value"])
        self.assertEqual("F-TEMA-MATERI", first["theme_evidence_ids"]["material"])
        self.assertEqual("Rendah", first["contradiction_review"]["severity"])
        self.assertEqual(100.0, first["confidence_basis"]["rating_coverage_pct"])
        self.assertIn("Perbandingan antarsegmen belum diuji.", first["confidence_basis"]["limitations"])

    def test_fingerprint_changes_without_exposing_record_content(self):
        first = self._build()
        changed = self.dataframe.copy()
        changed.loc[0, "Komentar"] = "Materi sudah sangat relevan."
        second = self._build(changed)

        self.assertNotEqual(first["snapshot_fingerprint"], second["snapshot_fingerprint"])
        serialized = repr(first)
        self.assertNotIn("A-001", serialized)
        self.assertNotIn("Materi perlu contoh praktik tambahan", serialized)

    def test_numeric_validator_rejects_added_or_changed_numbers(self):
        original = "Rating 4,5/5 dari 20 respons; target 30 hari."

        self.assertTrue(
            NarrativeFactValidator.preserves_numeric_facts(
                original,
                "Dari 20 respons, rating tercatat 4,5/5 dengan target 30 hari.",
            )
        )
        self.assertFalse(
            NarrativeFactValidator.preserves_numeric_facts(
                original,
                "Dari 20 respons, rating tercatat 4,8/5 dengan target 30 hari.",
            )
        )
        self.assertFalse(
            NarrativeFactValidator.preserves_numeric_facts(
                original,
                "Dari 20 respons, rating tercatat 4,5/5 dengan target 30 hari dan keyakinan 99%.",
            )
        )


if __name__ == "__main__":
    unittest.main()

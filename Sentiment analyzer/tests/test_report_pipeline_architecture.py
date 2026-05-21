import inspect
import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


class ReportPipelineArchitectureTests(unittest.TestCase):
    def test_pipeline_exposes_cohesive_report_generation_stages(self):
        import report_pipeline

        expected_stage_names = [
            "ReportRequestContext",
            "ReportResearchStage",
            "ReportAnalysisStage",
            "ReportNarrativeStage",
            "ReportPreflightQualityStage",
            "DocumentRenderStage",
            "ReportQualityStage",
            "ReportPipeline",
        ]
        for name in expected_stage_names:
            self.assertTrue(hasattr(report_pipeline, name), name)

    def test_report_generator_delegates_orchestration_to_pipeline(self):
        from report_engine import ReportGenerator

        source = inspect.getsource(ReportGenerator.run)
        self.assertIn("self.pipeline.run", source)
        self.assertNotIn("Document()", source)
        self.assertNotIn("FeedbackAnalyticsEngine(", source)
        self.assertNotIn("ReportQualityValidator.evaluate", source)
        self.assertLessEqual(len(source.splitlines()), 20)

    def test_report_generator_run_preserves_api_while_allowing_pipeline_injection(self):
        from report_engine import ReportGenerator

        class FakePipeline:
            def __init__(self):
                self.calls = []

            def run(self, **kwargs):
                self.calls.append(kwargs)
                return "document", "filename", {"verified_complete": True}

        kb = object()
        pipeline = FakePipeline()
        generator = ReportGenerator(kb, pipeline=pipeline)

        result = generator.run(
            "1 Bulan Terakhir (Monthly)",
            notes="Catatan eksekutif",
            sentiment="positive",
            segment="VIP",
            score_engine="experience_index",
        )

        self.assertEqual(result, ("document", "filename", {"verified_complete": True}))
        self.assertEqual(
            pipeline.calls,
            [
                {
                    "timeframe": "1 Bulan Terakhir (Monthly)",
                    "notes": "Catatan eksekutif",
                    "sentiment": "positive",
                    "segment": "VIP",
                    "score_engine": "experience_index",
                }
            ],
        )

    def test_narrative_stage_synthesizes_executive_snapshot_after_sections(self):
        from report_pipeline import ReportNarrativeStage, ReportRequestContext

        class FakeAnalytics:
            def __init__(self):
                self.sections_built = False

            def build_report_sections(self, *args, **kwargs):
                self.sections_built = True
                return [
                    {
                        "title": "Analisis Pengalaman Pelanggan",
                        "content": "Keluhan onboarding terkonsentrasi pada proses aktivasi awal.",
                    }
                ]

            def build_executive_snapshot(self, *args, **kwargs):
                self.assert_sections_ready(kwargs.get("report_sections"))
                return "## Ringkasan Eksekutif\nKeluhan onboarding terkonsentrasi pada proses aktivasi awal."

            def assert_sections_ready(self, report_sections):
                assert self.sections_built
                assert report_sections[0]["content"].startswith("### Bukti yang Dipakai")
                assert "Keluhan onboarding" in report_sections[0]["content"]

        snapshot, sections = ReportNarrativeStage().run(
            FakeAnalytics(),
            ReportRequestContext("Seluruh Periode"),
            "Tren eksternal ringkas.",
        )

        self.assertIn("Keluhan onboarding", snapshot)
        self.assertEqual(sections[0]["title"], "Analisis Pengalaman Pelanggan")

    def test_preflight_quality_runs_before_document_rendering(self):
        from report_pipeline import ReportPipeline

        class FakeResearch:
            def run(self, context):
                return "Tren eksternal."

        class FakeAnalysis:
            def run(self, dataframe):
                return object()

        class FakeNarrative:
            def run(self, analytics, context, macro_trends):
                return "## Ringkasan Eksekutif\nRingkas.", [{"id": "cx_chap_1", "title": "Bab", "content": ""}]

        class FakeDocument:
            def run(self, *args, **kwargs):
                raise AssertionError("document render should not run when preflight fails")

        pipeline = ReportPipeline(
            kb_instance=type("KB", (), {"df": object()})(),
            research_stage=FakeResearch(),
            analysis_stage=FakeAnalysis(),
            narrative_stage=FakeNarrative(),
            document_stage=FakeDocument(),
        )

        with self.assertRaises(ValueError):
            pipeline.run("Seluruh Periode")


if __name__ == "__main__":
    unittest.main()

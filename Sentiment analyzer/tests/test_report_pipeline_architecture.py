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


if __name__ == "__main__":
    unittest.main()

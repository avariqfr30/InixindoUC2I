import concurrent.futures
import inspect
import sys
import threading
import time
import unittest
from pathlib import Path
from unittest import mock

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


class ReportPipelineArchitectureTests(unittest.TestCase):
    @staticmethod
    def _build_success_pipeline(research_stage, analysis_stage, **overrides):
        from report_pipeline import ReportPipeline

        class FakeNarrative:
            def run(self, analytics, context, macro_trends):
                return (
                    "## Ringkasan Eksekutif\nRingkas.",
                    [{"id": "cx_chap_1", "title": "Bab", "content": "Isi."}],
                    "Rencana penulisan.",
                )

        class FakeWriting:
            def run(self, executive_snapshot, report_sections, planning_block):
                return executive_snapshot, report_sections

        class FakePreflight:
            def run(self, executive_snapshot, report_sections):
                return {"passes": True}

        class FakeDocument:
            def run(self, context, executive_snapshot, report_sections):
                return "document"

        class FakeQuality:
            def run(self, *args, **kwargs):
                return {"verified_complete": True}

        stages = {
            "narrative_stage": FakeNarrative(),
            "writing_stage": FakeWriting(),
            "preflight_stage": FakePreflight(),
            "document_stage": FakeDocument(),
            "quality_stage": FakeQuality(),
        }
        stages.update(overrides)
        return ReportPipeline(
            kb_instance=type("KB", (), {"df": object()})(),
            research_stage=research_stage,
            analysis_stage=analysis_stage,
            **stages,
        )

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

    def test_writing_stage_globally_bounds_parallel_polish_and_preserves_item_order(self):
        from report_pipeline import ReportWritingQualityStage

        class TrackingEditor:
            def __init__(self):
                self.active = 0
                self.max_active = 0
                self.overlap_seen = threading.Event()
                self.calls = []
                self.lock = threading.Lock()

            def polish(self, text, guidance=""):
                with self.lock:
                    self.active += 1
                    self.max_active = max(self.max_active, self.active)
                    self.calls.append((text, guidance))
                    if self.active >= 2:
                        self.overlap_seen.set()
                time.sleep(0.03)
                with self.lock:
                    self.active -= 1
                return f"polished:{text}:{guidance}"

        editor = TrackingEditor()
        stages = [ReportWritingQualityStage(editor), ReportWritingQualityStage(editor)]
        sections = [
            {
                "id": "first",
                "title": "First",
                "content": "content-first",
                "_writing_plan": "plan-first",
                "_document_contract": {"cache_key": "kept"},
            },
            {
                "id": "second",
                "title": "Second",
                "content": "content-second",
                "_writing_plan": "plan-second",
                "other": "kept",
            },
        ]

        with mock.patch(
            "report_pipeline.repair_feedback_document_spine",
            side_effect=lambda snapshot, report_sections: (snapshot, report_sections),
        ):
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                results = list(
                    executor.map(
                        lambda stage: stage.run("snapshot", sections, "plan-default"),
                        stages,
                    )
                )

        self.assertTrue(editor.overlap_seen.is_set())
        self.assertEqual(editor.max_active, 2)
        for snapshot, polished_sections in results:
            self.assertEqual(snapshot, "polished:snapshot:plan-default")
            self.assertEqual([item["id"] for item in polished_sections], ["first", "second"])
            self.assertEqual(
                [item["content"] for item in polished_sections],
                [
                    "polished:content-first:plan-first",
                    "polished:content-second:plan-second",
                ],
            )
            self.assertNotIn("_writing_plan", polished_sections[0])
            self.assertEqual(polished_sections[0]["_document_contract"], {"cache_key": "kept"})
            self.assertEqual(polished_sections[1]["other"], "kept")

    def test_writing_stage_reraises_unexpected_polish_failure(self):
        from report_pipeline import ReportWritingQualityStage

        class FailingEditor:
            def polish(self, text, guidance=""):
                if text == "broken":
                    raise RuntimeError("unexpected polish failure")
                return text

        with self.assertRaisesRegex(RuntimeError, "unexpected polish failure"):
            ReportWritingQualityStage(FailingEditor()).run(
                "snapshot",
                [{"title": "Broken", "content": "broken"}],
            )

    def test_writing_stage_detects_failure_and_cancels_queued_polish(self):
        from report_pipeline import ReportWritingQualityStage

        class PassthroughEditor:
            def polish(self, text, guidance=""):
                return text

        class ControlledExecutor:
            def __init__(self):
                self.futures = []

            def submit(self, operation, *args, **kwargs):
                future = concurrent.futures.Future()
                self.futures.append(future)
                if len(self.futures) == 2:
                    future.set_exception(RuntimeError("unexpected polish failure"))
                return future

        controlled_executor = ControlledExecutor()
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        try:
            with mock.patch(
                "report_pipeline._REPORT_WRITING_EXECUTOR",
                controlled_executor,
            ):
                stage_future = executor.submit(
                    ReportWritingQualityStage(PassthroughEditor()).run,
                    "snapshot",
                    [
                        {"title": "Broken", "content": "broken"},
                        {"title": "Queued 1", "content": "queued-1"},
                        {"title": "Queued 2", "content": "queued-2"},
                    ],
                )
                with self.assertRaisesRegex(RuntimeError, "unexpected polish failure"):
                    stage_future.result(timeout=0.3)
        finally:
            executor.shutdown(wait=True)

        self.assertEqual(len(controlled_executor.futures), 4)
        self.assertTrue(controlled_executor.futures[0].cancelled())
        self.assertFalse(controlled_executor.futures[1].cancelled())
        self.assertTrue(controlled_executor.futures[2].cancelled())
        self.assertTrue(controlled_executor.futures[3].cancelled())

    def test_writing_stage_preserves_protected_editor_model_failure_fallback(self):
        from report_pipeline import ReportWritingQualityStage
        from writing_quality import ProtectedIndonesianEditor

        class FailingClient:
            def chat(self, **kwargs):
                raise RuntimeError("model unavailable")

        source = "Nilai layanan 87% perlu diperbaiki karena kalimat ini sengaja cukup panjang untuk memicu penyuntingan konservatif."
        editor = ProtectedIndonesianEditor(
            model_client=FailingClient(),
            quality_fn=lambda text, protected: {
                "issues": ["long_sentence"],
                "protected_missing": [],
            },
        )

        with mock.patch(
            "report_pipeline.repair_feedback_document_spine",
            side_effect=lambda snapshot, report_sections: (snapshot, report_sections),
        ):
            snapshot, sections = ReportWritingQualityStage(editor).run(
                source,
                [{"title": "Section", "content": source}],
            )

        self.assertEqual(snapshot, source)
        self.assertEqual(sections[0]["content"], source)

    def test_writing_stage_submits_snapshot_and_sections_independently(self):
        from report_pipeline import ReportWritingQualityStage

        class OverlapEditor:
            def __init__(self):
                self.active = 0
                self.overlap_seen = threading.Event()
                self.lock = threading.Lock()

            def polish(self, text, guidance=""):
                with self.lock:
                    self.active += 1
                    if self.active >= 2:
                        self.overlap_seen.set()
                self.overlap_seen.wait(timeout=0.2)
                with self.lock:
                    self.active -= 1
                return text

        editor = OverlapEditor()
        with mock.patch(
            "report_pipeline.repair_feedback_document_spine",
            side_effect=lambda snapshot, report_sections: (snapshot, report_sections),
        ):
            ReportWritingQualityStage(editor).run(
                "snapshot",
                [{"title": "Section", "content": "content"}],
            )

        self.assertTrue(editor.overlap_seen.is_set())

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
                assert not report_sections[0]["content"].startswith("### Bukti yang Dipakai")
                assert "Keluhan onboarding" in report_sections[0]["content"]

        snapshot, sections, planning_block = ReportNarrativeStage().run(
            FakeAnalytics(),
            ReportRequestContext("Seluruh Periode"),
            "Tren eksternal ringkas.",
        )

        self.assertIn("Keluhan onboarding", snapshot)
        self.assertEqual(sections[0]["title"], "Analisis Pengalaman Pelanggan")
        self.assertIn("SECTION_PLANNER", planning_block)

    def test_narrative_stage_passes_synthesized_context_to_sections_and_snapshot(self):
        from report_pipeline import ReportNarrativeStage, ReportRequestContext

        class FakeAnalytics:
            full_df = None

            def __init__(self):
                self.section_context = None
                self.snapshot_context = None

            def build_report_sections(self, *args, **kwargs):
                self.section_context = kwargs.get("section_context")
                return [{"title": "Analisis", "content": "Konteks layanan perlu diprioritaskan."}]

            def build_executive_snapshot(self, *args, **kwargs):
                self.snapshot_context = kwargs.get("section_context")
                return "## Ringkasan Eksekutif\nKonteks layanan perlu diprioritaskan."

        analytics = FakeAnalytics()
        ReportNarrativeStage().run(
            analytics,
            ReportRequestContext(
                "Seluruh Periode",
                notes="APIDog source=/api/Resource/dataset meminta fokus Problem, Opportunity, Directive pada onboarding.",
            ),
            "Tren eksternal.",
        )

        self.assertIsNotNone(analytics.section_context)
        self.assertEqual(analytics.section_context, analytics.snapshot_context)
        self.assertNotIn("APIDog", analytics.section_context["focus_note"])
        self.assertIn("onboarding", analytics.section_context["focus_note"].lower())

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

    def test_research_and_local_analysis_overlap_without_changing_result_shape(self):
        analysis_started = threading.Event()

        class FakeResearch:
            def run(self, context):
                if not analysis_started.wait(timeout=1):
                    raise AssertionError("analysis did not overlap research")
                return "Tren eksternal."

        class FakeAnalysis:
            def run(self, dataframe):
                analysis_started.set()
                return "analytics"

        pipeline = self._build_success_pipeline(FakeResearch(), FakeAnalysis())

        result = pipeline.run("Seluruh Periode")

        self.assertEqual(
            result,
            (
                "document",
                "Inixindo_Feedback_Intelligence_Report_Experience_Index_Seluruh_Periode",
                {"verified_complete": True, "preflight": {"passes": True}},
            ),
        )

    def test_pipeline_stage_injection_remains_compatible_for_successful_run(self):
        calls = []

        class FakeResearch:
            def run(self, context):
                calls.append(("research", context.timeframe))
                return "Tren eksternal."

        class FakeAnalysis:
            def run(self, dataframe):
                calls.append(("analysis", dataframe))
                return "analytics"

        dataframe = object()
        pipeline = self._build_success_pipeline(FakeResearch(), FakeAnalysis())
        pipeline.kb.df = dataframe

        document, filename, quality = pipeline.run("Seluruh Periode")

        self.assertEqual(document, "document")
        self.assertTrue(filename.startswith("Inixindo_Feedback_Intelligence_Report_"))
        self.assertTrue(quality["verified_complete"])
        self.assertCountEqual(calls, [("research", "Seluruh Periode"), ("analysis", dataframe)])

    def test_stage_timings_are_logged_once_and_never_added_to_quality(self):
        class FakeResearch:
            def run(self, context):
                return "Tren eksternal."

        class FakeAnalysis:
            def run(self, dataframe):
                return "analytics"

        pipeline = self._build_success_pipeline(FakeResearch(), FakeAnalysis())

        with self.assertLogs("report_pipeline", level="INFO") as captured:
            _, _, quality = pipeline.run("Seluruh Periode")

        timing_logs = [message for message in captured.output if "stage timings" in message]
        self.assertEqual(len(timing_logs), 1)
        for stage_name in ("research", "analysis", "narrative", "writing", "preflight", "render", "quality"):
            self.assertRegex(timing_logs[0], rf"\b{stage_name}=\d+\.\d{{6}}\b")
        self.assertEqual(quality, {"verified_complete": True, "preflight": {"passes": True}})
        self.assertFalse(any("tim" in key.lower() for key in quality))

    def test_default_research_stage_uses_pipeline_executor_without_nested_submission(self):
        class RecordingExecutor:
            def __init__(self, delegate):
                self.delegate = delegate
                self.operations = []

            def submit(self, operation, *args):
                self.operations.append(operation)
                return self.delegate.submit(operation, *args)

        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as delegate:
            executor = RecordingExecutor(delegate)
            pipeline = self._build_success_pipeline(
                None,
                type("Analysis", (), {"run": lambda self, dataframe: "analytics"})(),
                orchestration_executor=executor,
            )
            stage = pipeline.research_stage
            with mock.patch("report_pipeline.Researcher.get_macro_trends", return_value="Tren eksternal."):
                pipeline.run("Seluruh Periode")

        self.assertIsNone(stage.executor)
        self.assertEqual(len(executor.operations), 1)
        self.assertEqual(executor.operations[0].__name__, "_lookup")

    def test_research_resolve_preserves_timeout_and_fallback(self):
        from report_pipeline import ReportResearchStage

        class FailingFuture:
            timeout = None

            def result(self, timeout=None):
                self.timeout = timeout
                raise concurrent.futures.TimeoutError("research timed out")

        future = FailingFuture()
        with self.assertLogs("report_pipeline", level="ERROR") as captured:
            result = ReportResearchStage().resolve(future)

        self.assertEqual(future.timeout, 45)
        self.assertEqual(result, "Tidak ada tren eksternal yang berhasil dimuat.")
        self.assertTrue(any("OSINT macro trend lookup failed" in message for message in captured.output))

    def test_analysis_failure_cancels_pending_research(self):
        class PendingFuture:
            def __init__(self):
                self.cancel_called = False
                self.callback = None

            def add_done_callback(self, callback):
                self.callback = callback

            def cancel(self):
                self.cancel_called = True
                if self.callback:
                    self.callback(self)
                return True

        class PendingExecutor:
            def __init__(self, future):
                self.future = future

            def submit(self, operation, *args):
                return self.future

        class FailingAnalysis:
            def run(self, dataframe):
                raise ValueError("analysis failed")

        future = PendingFuture()
        pipeline = self._build_success_pipeline(
            None,
            FailingAnalysis(),
            orchestration_executor=PendingExecutor(future),
        )

        with self.assertRaisesRegex(ValueError, "analysis failed"):
            pipeline.run("Seluruh Periode")

        self.assertTrue(future.cancel_called)


if __name__ == "__main__":
    unittest.main()

import re
import sys
import threading
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from report_agents import FeedbackProposalTeam
from report_analytics import FeedbackAnalyticsEngine
from report_pipeline import (
    ReportNarrativeStage,
    ReportPipeline,
    ReportQualityStage,
    ReportRequestContext,
)


class PreparedReportAnalysisTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dataframe = pd.read_csv(PROJECT_DIR / "data" / "db.csv")
        cls.timeframe = "1 Bulan Terakhir (Monthly)"
        cls.notes = "Periksa risiko jadwal dan tindak lanjut layanan."
        cls.macro_trends = "Tidak ada tren eksternal yang berhasil dimuat."

    def test_preparation_is_reused_by_sections_executive_and_final_quality(self):
        class CountingEngine(FeedbackAnalyticsEngine):
            def __init__(self, dataframe):
                self.analysis_context_builds = 0
                super().__init__(dataframe)

            def _build_analysis_context(self, *args, **kwargs):
                self.analysis_context_builds += 1
                return super()._build_analysis_context(*args, **kwargs)

        engine = CountingEngine(self.dataframe)
        context = ReportRequestContext(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
        )

        with mock.patch.object(
            FeedbackProposalTeam,
            "_contradiction_review",
            wraps=FeedbackProposalTeam._contradiction_review,
        ) as contradiction_review:
            prepared = engine.prepare_report_analysis(
                context.timeframe,
                sentiment=context.sentiment,
                segment=context.segment,
                score_engine=context.score_engine,
            )
            snapshot, sections, _planning = ReportNarrativeStage().run(
                engine,
                context,
                self.macro_trends,
            )
            with mock.patch(
                "report_pipeline.ReportQualityValidator.evaluate",
                return_value={"verified_complete": True},
            ):
                quality = ReportQualityStage().run(
                    object(),
                    snapshot,
                    sections,
                    engine,
                    self.dataframe,
                    context,
                    self.macro_trends,
                    prepared_analysis=prepared,
                )

        self.assertTrue(snapshot)
        self.assertTrue(sections)
        self.assertEqual(engine.analysis_context_builds, 1)
        self.assertEqual(contradiction_review.call_count, 1)
        self.assertIs(
            engine.prepare_report_analysis(
                context.timeframe,
                sentiment=context.sentiment,
                segment=context.segment,
                score_engine=context.score_engine,
            ),
            prepared,
        )
        self.assertIs(
            engine.get_prepared_report_analysis(
                context.timeframe,
                sentiment=context.sentiment,
                segment=context.segment,
                score_engine=context.score_engine,
            ),
            prepared,
        )
        self.assertEqual(quality["contradiction_review"], prepared.contradiction_review)

    def test_prepared_helpers_compute_once_for_the_scoped_dataframe(self):
        engine = FeedbackAnalyticsEngine(self.dataframe)

        with (
            mock.patch.object(
                engine,
                "_compute_governance_summary",
                wraps=engine._compute_governance_summary,
            ) as governance_compute,
            mock.patch.object(
                engine,
                "_compute_theme_hits",
                wraps=engine._compute_theme_hits,
            ) as theme_compute,
            mock.patch.object(
                engine,
                "_compute_group_risk",
                wraps=engine._compute_group_risk,
            ) as group_risk_compute,
        ):
            prepared = engine.prepare_report_analysis(
                self.timeframe,
                score_engine="experience_index",
            )
            snapshot, sections, _planning = ReportNarrativeStage().run(
                engine,
                ReportRequestContext(
                    self.timeframe,
                    self.notes,
                    score_engine="experience_index",
                ),
                self.macro_trends,
            )
            FeedbackProposalTeam().run(
                engine,
                self.dataframe,
                self.timeframe,
                macro_trends=self.macro_trends,
                score_engine="experience_index",
                prepared_analysis=prepared,
            )

        scoped = prepared.scoped_dataframe
        self.assertTrue(snapshot)
        self.assertTrue(sections)
        self.assertEqual(
            sum(call.args[0] is scoped for call in governance_compute.call_args_list),
            1,
        )
        self.assertEqual(
            sum(call.args[0] is scoped for call in theme_compute.call_args_list),
            1,
        )
        grouped_columns = [
            call.args[1]
            for call in group_risk_compute.call_args_list
            if call.args[0] is scoped
        ]
        self.assertEqual(grouped_columns.count("Layanan"), 1)
        self.assertEqual(grouped_columns.count("Tipe Stakeholder"), 1)
        self.assertEqual(grouped_columns.count("Lokasi"), 1)
        self.assertEqual(grouped_columns.count("Tipe Instruktur"), 1)

    def test_prepared_packet_from_another_engine_is_rejected(self):
        first_engine = FeedbackAnalyticsEngine(self.dataframe)
        second_engine = FeedbackAnalyticsEngine(self.dataframe)
        foreign_packet = first_engine.prepare_report_analysis(
            self.timeframe,
            score_engine="experience_index",
        )
        local_packet = second_engine.prepare_report_analysis(
            self.timeframe,
            score_engine="experience_index",
        )

        self.assertIsNone(
            second_engine.resolve_prepared_report_analysis(
                foreign_packet,
                self.timeframe,
                score_engine="experience_index",
            )
        )
        self.assertIs(
            second_engine.resolve_prepared_report_analysis(
                local_packet,
                self.timeframe,
                score_engine="experience_index",
            ),
            local_packet,
        )

    def test_default_pipeline_prepares_analysis_while_research_is_running(self):
        preparation_started = threading.Event()

        class WaitingResearch:
            def run(self, context):
                if not preparation_started.wait(timeout=2):
                    raise AssertionError("prepared analysis did not overlap research")
                return "Tidak ada tren eksternal yang berhasil dimuat."

        class FakeNarrative:
            def run(self, analytics, context, macro_trends):
                self.prepared = analytics.get_prepared_report_analysis(
                    context.timeframe,
                    sentiment=context.sentiment,
                    segment=context.segment,
                    score_engine=context.score_engine,
                )
                return "Ringkasan", [{"title": "Bab", "content": "Isi"}], "Rencana"

        class PassthroughWriting:
            def run(self, snapshot, sections, planning):
                return snapshot, sections

        class PassingPreflight:
            def run(self, snapshot, sections):
                return {"passes": True}

        class FakeDocument:
            def run(self, *args):
                return "document"

        class FakeQuality:
            def run(self, *args):
                return {"verified_complete": True}

        original_prepare = FeedbackAnalyticsEngine.prepare_report_analysis

        def signal_prepare(engine, *args, **kwargs):
            preparation_started.set()
            return original_prepare(engine, *args, **kwargs)

        narrative = FakeNarrative()
        pipeline = ReportPipeline(
            kb_instance=type("KB", (), {"df": self.dataframe})(),
            research_stage=WaitingResearch(),
            narrative_stage=narrative,
            writing_stage=PassthroughWriting(),
            preflight_stage=PassingPreflight(),
            document_stage=FakeDocument(),
            quality_stage=FakeQuality(),
        )

        with mock.patch.object(
            FeedbackAnalyticsEngine,
            "prepare_report_analysis",
            new=signal_prepare,
        ):
            pipeline.run(self.timeframe)

        self.assertIsNotNone(narrative.prepared)

    def test_direct_unprepared_builders_remain_compatible(self):
        engine = FeedbackAnalyticsEngine(self.dataframe)

        sections = engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
        )
        snapshot = engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
            macro_trends=self.macro_trends,
            report_sections=sections,
        )

        self.assertTrue(snapshot.strip())
        self.assertTrue(all(section["content"].strip() for section in sections))
        self.assertIsNone(
            engine.get_prepared_report_analysis(
                self.timeframe,
                score_engine="experience_index",
            )
        )

    def test_legacy_builder_does_not_implicitly_consume_prepared_packet(self):
        engine = FeedbackAnalyticsEngine(self.dataframe)
        engine.prepare_report_analysis(
            self.timeframe,
            score_engine="experience_index",
        )

        with mock.patch.object(
            engine,
            "_build_analysis_context",
            wraps=engine._build_analysis_context,
        ) as context_builder:
            engine.build_report_sections(
                self.timeframe,
                self.notes,
                self.macro_trends,
                score_engine="experience_index",
            )

        self.assertEqual(context_builder.call_count, 1)

    def test_prepared_and_unprepared_outputs_are_equivalent(self):
        prepared_engine = FeedbackAnalyticsEngine(self.dataframe)
        prepared = prepared_engine.prepare_report_analysis(
            self.timeframe,
            score_engine="experience_index",
        )
        prepared_sections = prepared_engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
            prepared_analysis=prepared,
        )
        prepared_snapshot = prepared_engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
            macro_trends=self.macro_trends,
            report_sections=prepared_sections,
            prepared_analysis=prepared,
        )
        prepared_briefing = FeedbackProposalTeam().run(
            prepared_engine,
            self.dataframe,
            self.timeframe,
            macro_trends=self.macro_trends,
            score_engine="experience_index",
            prepared_analysis=prepared,
        )

        unprepared_engine = FeedbackAnalyticsEngine(self.dataframe)
        unprepared_sections = unprepared_engine.build_report_sections(
            self.timeframe,
            self.notes,
            self.macro_trends,
            score_engine="experience_index",
        )
        unprepared_snapshot = unprepared_engine.build_executive_snapshot(
            self.timeframe,
            self.notes,
            score_engine="experience_index",
            macro_trends=self.macro_trends,
            report_sections=unprepared_sections,
        )
        unprepared_briefing = FeedbackProposalTeam().run(
            unprepared_engine,
            self.dataframe,
            self.timeframe,
            macro_trends=self.macro_trends,
            score_engine="experience_index",
        )

        timestamp_pattern = r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?\+00:00"
        self.assertEqual(prepared_sections, unprepared_sections)
        self.assertEqual(
            re.sub(timestamp_pattern, "<generated-at>", prepared_snapshot),
            re.sub(timestamp_pattern, "<generated-at>", unprepared_snapshot),
        )
        prepared_briefing["audit_trail"].pop("generated_at_utc", None)
        unprepared_briefing["audit_trail"].pop("generated_at_utc", None)
        self.assertEqual(prepared_briefing, unprepared_briefing)


if __name__ == "__main__":
    unittest.main()

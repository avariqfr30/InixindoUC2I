import concurrent.futures
import logging
from dataclasses import dataclass

from docx import Document

from config import DEFAULT_COLOR, DEFAULT_SCORE_ENGINE, SCORE_ENGINE_PROFILES
from document_builder import DocumentBuilder
from osint_research import Researcher
from report_agents import FeedbackProposalTeam
from report_analytics import FeedbackAnalyticsEngine
from report_evidence import ContextIntelligenceDesk
from report_quality import ReportQualityValidator
from timeframe_filters import readable_timeframe_label

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ReportRequestContext:
    timeframe: str
    notes: str = ""
    sentiment: str = "all"
    segment: str = "all"
    score_engine: str = DEFAULT_SCORE_ENGINE

    @property
    def score_profile(self):
        return SCORE_ENGINE_PROFILES.get(
            self.score_engine,
            SCORE_ENGINE_PROFILES[DEFAULT_SCORE_ENGINE],
        )

    @property
    def timeframe_label(self):
        return readable_timeframe_label(self.timeframe)


class ReportResearchStage:
    def __init__(self, executor=None):
        self.executor = executor or concurrent.futures.ThreadPoolExecutor(max_workers=4)

    def run(self, context):
        future = self.executor.submit(
            Researcher.get_macro_trends,
            context.timeframe_label,
            context.notes,
            context.score_profile["label"],
        )
        try:
            return future.result(timeout=45)
        except Exception:
            logger.exception("OSINT macro trend lookup failed during report generation.")
            return "Tidak ada tren eksternal yang berhasil dimuat."


class ReportAnalysisStage:
    def run(self, dataframe):
        return FeedbackAnalyticsEngine(dataframe)


class ReportNarrativeStage:
    def run(self, analytics, context, macro_trends):
        section_context = ContextIntelligenceDesk.build(
            dataframe=getattr(analytics, "full_df", None),
            notes=context.notes,
            timeframe=context.timeframe,
            sentiment=context.sentiment,
            segment=context.segment,
            score_engine=context.score_engine,
            macro_trends=macro_trends,
        )
        report_sections = analytics.build_report_sections(
            context.timeframe,
            section_context["focus_note"],
            macro_trends,
            sentiment=context.sentiment,
            segment=context.segment,
            score_engine=context.score_engine,
            section_context=section_context,
        )
        executive_snapshot = analytics.build_executive_snapshot(
            context.timeframe,
            section_context["focus_note"],
            sentiment=context.sentiment,
            segment=context.segment,
            score_engine=context.score_engine,
            macro_trends=macro_trends,
            report_sections=report_sections,
            section_context=section_context,
        )
        return executive_snapshot, report_sections


class ReportPreflightQualityStage:
    def run(self, executive_snapshot, report_sections):
        result = ReportQualityValidator.evaluate_narrative(executive_snapshot, report_sections)
        if not result["passes"]:
            raise ValueError("Report narrative preflight failed: " + "; ".join(result["findings"]))
        return result


class DocumentRenderStage:
    def run(self, context, executive_snapshot, report_sections):
        document = Document()
        DocumentBuilder.create_cover(document, context.timeframe_label, DEFAULT_COLOR)
        document.add_heading("Ringkasan Eksekutif", level=1)
        DocumentBuilder.process_content(document, executive_snapshot, DEFAULT_COLOR)
        document.add_page_break()

        for index, section in enumerate(report_sections):
            document.add_heading(section["title"], level=1)
            DocumentBuilder.process_content(document, section["content"], DEFAULT_COLOR)
            if index < len(report_sections) - 1:
                document.add_page_break()
        return document


DocumentAssemblyStage = DocumentRenderStage


class ReportQualityStage:
    def run(
        self,
        document,
        executive_snapshot,
        report_sections,
        analytics,
        dataframe,
        context,
        macro_trends,
    ):
        quality = ReportQualityValidator.evaluate(
            document,
            executive_snapshot,
            report_sections,
            context.score_profile["label"],
        )
        briefing = FeedbackProposalTeam().run(
            analytics,
            dataframe,
            context.timeframe,
            macro_trends=macro_trends,
            sentiment=context.sentiment,
            segment=context.segment,
            score_engine=context.score_engine,
        )
        quality["audit_trail"] = briefing.get("audit_trail", {})
        quality["confidence"] = briefing.get("confidence")
        quality["contradiction_review"] = briefing.get("contradiction_review", {})
        quality["trend_review"] = briefing.get("trend_review", {})
        quality["prediction_review"] = briefing.get("prediction_review", {})
        quality["agent_desk"] = briefing.get("agent_desk", {})
        if not quality["verified_complete"]:
            logger.warning(
                "Generated report is below completeness target: %s",
                quality["summary"],
            )
        return quality


class ReportPipeline:
    def __init__(
        self,
        kb_instance,
        research_stage=None,
        analysis_stage=None,
        narrative_stage=None,
        document_stage=None,
        quality_stage=None,
        preflight_stage=None,
    ):
        self.kb = kb_instance
        self.research_stage = research_stage or ReportResearchStage()
        self.analysis_stage = analysis_stage or ReportAnalysisStage()
        self.narrative_stage = narrative_stage or ReportNarrativeStage()
        self.document_stage = document_stage or DocumentRenderStage()
        self.preflight_stage = preflight_stage or ReportPreflightQualityStage()
        self.quality_stage = quality_stage or ReportQualityStage()

    def run(
        self,
        timeframe,
        notes="",
        sentiment="all",
        segment="all",
        score_engine=DEFAULT_SCORE_ENGINE,
    ):
        context = ReportRequestContext(timeframe, notes, sentiment, segment, score_engine)
        macro_trends = self.research_stage.run(context)
        analytics = self.analysis_stage.run(self.kb.df)
        executive_snapshot, report_sections = self.narrative_stage.run(
            analytics,
            context,
            macro_trends,
        )
        preflight_quality = self.preflight_stage.run(executive_snapshot, report_sections)
        document = self.document_stage.run(context, executive_snapshot, report_sections)
        quality = self.quality_stage.run(
            document,
            executive_snapshot,
            report_sections,
            analytics,
            self.kb.df,
            context,
            macro_trends,
        )
        quality["preflight"] = preflight_quality
        filename = (
            f"Inixindo_Feedback_Intelligence_Report_{context.score_profile['label']}_{context.timeframe_label}"
        ).replace(" ", "_")
        return document, filename, quality

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
from report_planning import FeedbackSectionPlanner
from report_quality import ReportQualityValidator
from editorial_intelligence import repair_feedback_document_spine
from feedback_deliberation import FeedbackDeliberationBuilder
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
    def __init__(self, planner=None):
        self.planner = planner or FeedbackSectionPlanner()

    def run(self, analytics, context, macro_trends):
        osint_dossier = Researcher.build_osint_dossier(
            macro_trends,
            {
                "timeframe_label": context.timeframe_label,
                "timeframe": context.timeframe,
                "sentiment": context.sentiment,
                "segment": context.segment,
                "score_engine": context.score_engine,
                "score_engine_label": context.score_profile["label"],
            },
        )
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
        deliberation_builder = FeedbackDeliberationBuilder()
        document_contract = deliberation_builder.build(
            report_sections or [],
            {
                **section_context,
                "timeframe_label": context.timeframe_label,
                "timeframe": context.timeframe,
                "sentiment": context.sentiment,
                "segment": context.segment,
            },
            data_version=str(section_context.get("data_version") or ""),
        )
        planning_block = self.planner.build_prompt_block(
            sections=["Ringkasan Eksekutif", *[section.get("title", "") for section in report_sections or []]],
            context={
                "timeframe_label": context.timeframe_label,
                "timeframe": context.timeframe,
                "sentiment": context.sentiment,
                "segment": context.segment,
                "osint_dossier": osint_dossier,
                "row_count": section_context.get("row_count"),
                "text_response_count": section_context.get("text_response_count"),
                "external_context_ready": section_context.get("external_context_ready"),
                "insight_cards": section_context.get("insight_cards"),
                "document_contract": document_contract,
            },
        )
        for section in report_sections or []:
            section["_writing_plan"] = "\n".join(
                [planning_block, deliberation_builder.for_section(document_contract, section.get("id") or "")]
            )
            section["_document_contract"] = document_contract
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
        return executive_snapshot, report_sections, planning_block


class ReportWritingQualityStage:
    def __init__(self, editor=None):
        if editor is None:
            from writing_quality import ProtectedIndonesianEditor
            editor = ProtectedIndonesianEditor()
        self.editor = editor

    def run(self, executive_snapshot, report_sections, planning_block=""):
        polished_snapshot = self.editor.polish(executive_snapshot, guidance=planning_block)
        polished_sections = []
        for section in report_sections or []:
            item = dict(section)
            item["content"] = self.editor.polish(
                item.get("content", ""),
                guidance=item.get("_writing_plan") or planning_block,
            )
            item.pop("_writing_plan", None)
            polished_sections.append(item)
        polished_snapshot, polished_sections = repair_feedback_document_spine(polished_snapshot, polished_sections)
        return polished_snapshot, polished_sections


class ReportPreflightQualityStage:
    def run(self, executive_snapshot, report_sections):
        contract = next(
            (section.get("_document_contract") for section in report_sections or [] if section.get("_document_contract")),
            None,
        )
        appendix = FeedbackDeliberationBuilder.build_appendix_markdown(contract) if contract else ""
        result = ReportQualityValidator.evaluate_narrative(
            executive_snapshot,
            report_sections,
            deliberation_contract=contract,
            appendix_content=appendix,
        )
        if not result["passes"]:
            raise ValueError("Report narrative preflight failed: " + "; ".join(result["findings"]))
        return result


class DocumentRenderStage:
    def run(self, context, executive_snapshot, report_sections):
        document = Document()
        DocumentBuilder.create_cover(document, context.timeframe_label, DEFAULT_COLOR)
        document.add_heading("Ringkasan Eksekutif", level=1)
        DocumentBuilder.process_content(document, executive_snapshot, DEFAULT_COLOR)

        for section in report_sections:
            document.add_page_break()
            document.add_heading(section["title"], level=1)
            DocumentBuilder.process_content(document, section["content"], DEFAULT_COLOR)
        contract = next(
            (section.get("_document_contract") for section in report_sections or [] if section.get("_document_contract")),
            None,
        )
        if contract:
            document.add_page_break()
            appendix = FeedbackDeliberationBuilder.build_appendix_markdown(contract)
            DocumentBuilder.process_content(document, appendix, DEFAULT_COLOR)
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
        contract = next(
            (section.get("_document_contract") for section in report_sections or [] if section.get("_document_contract")),
            {},
        )
        quality["document_deliberation"] = {
            "cache_key": contract.get("cache_key"),
            "accepted_claim_count": len(contract.get("claim_ledger") or []),
            "data_gap_count": len(contract.get("data_gap_register") or []),
            "appendix_sections": list((contract.get("appendix_manifest") or {}).keys()),
        }
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
        writing_stage=None,
    ):
        self.kb = kb_instance
        self.research_stage = research_stage or ReportResearchStage()
        self.analysis_stage = analysis_stage or ReportAnalysisStage()
        self.narrative_stage = narrative_stage or ReportNarrativeStage()
        self.document_stage = document_stage or DocumentRenderStage()
        self.preflight_stage = preflight_stage or ReportPreflightQualityStage()
        self.writing_stage = writing_stage or ReportWritingQualityStage()
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
        executive_snapshot, report_sections, planning_block = self.narrative_stage.run(
            analytics,
            context,
            macro_trends,
        )
        executive_snapshot, report_sections = self.writing_stage.run(executive_snapshot, report_sections, planning_block)
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

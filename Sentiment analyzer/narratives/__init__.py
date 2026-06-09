from config import CX_SENTIMENT_STRUCTURE, DEFAULT_SCORE_ENGINE
from document_builder import DocumentBuilder

from .base import BaseNarrativeMixin
from .descriptive import DescriptiveNarrativeMixin
from .diagnostic import DiagnosticNarrativeMixin
from .predictive import PredictiveNarrativeMixin
from .prescriptive import PrescriptiveNarrativeMixin
from .implementation import ImplementationNarrativeMixin
from .executive import ExecutiveNarrativeMixin


class ReportNarrativeBuilderMixin(
    DescriptiveNarrativeMixin,
    DiagnosticNarrativeMixin,
    PredictiveNarrativeMixin,
    PrescriptiveNarrativeMixin,
    ImplementationNarrativeMixin,
    ExecutiveNarrativeMixin,
):
    """Markdown report narrative builders for FeedbackAnalyticsEngine.

    The mixin deliberately depends on analytics/context helper methods supplied by
    FeedbackAnalyticsEngine; it owns wording, tables, and report section assembly.
    """

    def build_report_sections(self, timeframe, notes, macro_trends, sentiment="all", segment="all", score_engine=DEFAULT_SCORE_ENGINE, section_context=None):
        timeframe_df = self._filter_view(timeframe, sentiment=sentiment, segment=segment)
        context = self._build_analysis_context(timeframe_df, timeframe, sentiment, segment, score_engine)
        section_map = {
            "cx_chap_1": self._descriptive_markdown(timeframe_df, timeframe, notes, context, section_context=section_context),
            "cx_chap_2": self._diagnostic_markdown(timeframe_df, context),
            "cx_chap_3": self._predictive_markdown(timeframe_df, macro_trends, context, section_context=section_context),
            "cx_chap_4": self._prescriptive_markdown(timeframe_df, context),
            "cx_chap_5": self._implementation_readiness_markdown(timeframe_df, timeframe, notes, macro_trends, context, section_context=section_context),
        }
        return [
            {
                "id": chapter["id"],
                "title": DocumentBuilder.reader_facing_text(chapter["title"]),
                "content": DocumentBuilder.reader_facing_text(section_map.get(chapter["id"], "")),
            }
            for chapter in CX_SENTIMENT_STRUCTURE
        ]

"""Compatibility facade for the feedback intelligence domain modules.

The implementation is split by responsibility:
- data_pipeline: internal data loading, normalization, and knowledge-base cache
- osint_research: external benchmark/OSINT collection
- report_analytics: scoring, risk analysis, and report context
- report_narratives: markdown report section assembly
- document_builder: DOCX styling, charts, and rendering
- report_engine: final report orchestration

Existing app and scripts can continue importing from core while newer code can import
from the focused modules directly.
"""

from data_pipeline import (
    CANONICAL_INTERNAL_COLUMNS,
    COLUMN_ALIASES,
    DATE_COLUMN_ALIASES,
    DemoCsvProvider,
    InternalApiProvider,
    InternalDataProvider,
    KnowledgeBase,
)
from document_builder import ChartEngine, DocumentBuilder, StyleEngine, append_field
from osint_research import InsightSchema, Researcher, osint_cache
from report_analytics import FeedbackAnalyticsEngine
from report_engine import ReportGenerator
from report_quality import ReportQualityValidator

__all__ = (
    "CANONICAL_INTERNAL_COLUMNS",
    "COLUMN_ALIASES",
    "DATE_COLUMN_ALIASES",
    "DemoCsvProvider",
    "InternalApiProvider",
    "InternalDataProvider",
    "KnowledgeBase",
    "InsightSchema",
    "Researcher",
    "osint_cache",
    "ChartEngine",
    "DocumentBuilder",
    "FeedbackAnalyticsEngine",
    "ReportGenerator",
    "ReportQualityValidator",
    "StyleEngine",
    "append_field",
)

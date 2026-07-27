import logging

from config import DEFAULT_SCORE_ENGINE
from report_pipeline import ReportPipeline

logger = logging.getLogger(__name__)


class ReportGenerator:
    def __init__(self, kb_instance, pipeline=None):
        self.kb = kb_instance
        self.pipeline = pipeline or ReportPipeline(kb_instance)

    def run(self, timeframe, notes="", sentiment="all", segment="all", score_engine=DEFAULT_SCORE_ENGINE, improvement_guidance=""):
        logger.info(
            "Starting feedback intelligence report generation for timeframe=%s, sentiment=%s, segment=%s, score_engine=%s",
            timeframe,
            sentiment,
            segment,
            score_engine,
        )
        # ReportPipeline owns the stage orchestration; this facade preserves the public API.
        pipeline_kwargs = dict(
            timeframe=timeframe,
            notes=notes,
            sentiment=sentiment,
            segment=segment,
            score_engine=score_engine,
        )
        if improvement_guidance:
            pipeline_kwargs["improvement_guidance"] = improvement_guidance
        return self.pipeline.run(**pipeline_kwargs)

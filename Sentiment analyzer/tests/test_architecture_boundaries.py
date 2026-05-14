import importlib
import sys
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


class ArchitectureBoundaryTests(unittest.TestCase):
    def test_data_contract_and_class_report_adapter_are_separate_modules(self):
        data_contract = importlib.import_module("data_contract")
        class_report_adapter = importlib.import_module("class_report_adapter")

        self.assertTrue(hasattr(data_contract, "CANONICAL_INTERNAL_COLUMNS"))
        self.assertTrue(hasattr(data_contract, "COLUMN_ALIASES"))
        self.assertTrue(hasattr(class_report_adapter, "ClassReportAdapter"))

        data_pipeline_source = (PROJECT_DIR / "data_pipeline.py").read_text()
        self.assertNotIn("CLASS_REPORT_LABEL_OVERRIDES =", data_pipeline_source)
        self.assertNotIn("CLASS_REPORT_JOURNEY_RULES =", data_pipeline_source)

    def test_knowledge_base_is_not_defined_inside_data_pipeline(self):
        knowledge_base = importlib.import_module("knowledge_base")
        data_pipeline_source = (PROJECT_DIR / "data_pipeline.py").read_text()

        self.assertTrue(hasattr(knowledge_base, "KnowledgeBase"))
        self.assertNotIn("class KnowledgeBase", data_pipeline_source)

    def test_report_trust_rendering_is_not_embedded_in_narrative_mixin(self):
        report_trust_sections = importlib.import_module("report_trust_sections")
        narrative_source = (PROJECT_DIR / "report_narratives.py").read_text()

        self.assertTrue(hasattr(report_trust_sections, "build_specialist_review_markdown"))
        self.assertNotIn("Report Audit Trail", narrative_source)
        self.assertNotIn("Contradiction Check", narrative_source)


if __name__ == "__main__":
    unittest.main()

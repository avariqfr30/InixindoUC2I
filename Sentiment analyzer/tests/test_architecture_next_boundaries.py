import importlib
import inspect
import sys
from pathlib import Path
import unittest

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))


class NextArchitectureBoundaryTests(unittest.TestCase):
    def test_report_request_centralizes_payload_validation(self):
        services = importlib.import_module("app_services")
        self.assertTrue(hasattr(services, "ReportRequest"))
        request_obj = services.ReportRequest.from_mapping({"timeframe": " Q1 ", "notes": " note "})
        self.assertEqual(request_obj.timeframe, "Q1")
        self.assertEqual(request_obj.notes, "note")
        self.assertEqual(request_obj.to_job_payload()["sentiment"], "all")
        with self.assertRaises(ValueError):
            services.ReportRequest.from_mapping({"timeframe": ""}).validate()

    def test_job_presenter_returns_user_ready_status_contract(self):
        presenter = importlib.import_module("job_presenter")
        job = {
            "job_id": "abc",
            "status": "completed",
            "queue_position": 0,
            "total_elapsed_seconds": 12.34,
            "quality": {"verified_complete": True},
        }
        presented = presenter.present_job(job, status_url="/jobs/abc", download_url="/download/abc")
        self.assertEqual(presented["stage_label"], "Laporan siap diunduh")
        self.assertTrue(presented["can_download"])
        self.assertFalse(presented["retryable"])
        self.assertEqual(presented["status_url"], "/jobs/abc")
        self.assertEqual(presented["download_url"], "/download/abc")

    def test_internal_api_service_owns_settings_and_refresh_boundaries(self):
        service_module = importlib.import_module("internal_api_service")
        self.assertTrue(hasattr(service_module, "InternalApiService"))
        self.assertTrue(hasattr(service_module.InternalApiService, "get_state"))
        self.assertTrue(hasattr(service_module.InternalApiService, "refresh_dataset"))
        self.assertTrue(hasattr(service_module.InternalApiService, "save_and_refresh"))
        self.assertTrue(hasattr(service_module.InternalApiService, "ensure_mutation_allowed"))
        app_source = Path(APP_DIR / "app.py").read_text(encoding="utf-8")
        self.assertNotIn("build_connector_payload", app_source)
        self.assertNotIn("write_connector_payload", app_source)

    def test_report_engine_delegates_pipeline_stages(self):
        pipeline = importlib.import_module("report_pipeline")
        for name in [
            "ReportResearchStage",
            "ReportAnalysisStage",
            "DocumentAssemblyStage",
            "ReportQualityStage",
            "ReportPipeline",
        ]:
            self.assertTrue(hasattr(pipeline, name), name)
        source = inspect.getsource(importlib.import_module("report_engine").ReportGenerator.run)
        self.assertIn("ReportPipeline", source)
        self.assertNotIn("DocumentBuilder.process_content", source)
        self.assertNotIn("Researcher.get_macro_trends", source)

    def test_application_factory_exists_without_removing_legacy_wsgi_app(self):
        app_module = importlib.import_module("app")
        self.assertTrue(hasattr(app_module, "create_app"))
        self.assertIs(app_module.create_app(), app_module.app)


if __name__ == "__main__":
    unittest.main()

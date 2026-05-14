import sys
import unittest
from pathlib import Path
from unittest import mock

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from internal_api_service import InternalApiService


class InternalApiServiceTests(unittest.TestCase):
    def _service(self, connector_state=None, queued=0, running=0, refresh_success=True):
        fake_kb = mock.Mock()
        fake_kb.provider = mock.Mock(source_name="company_api")
        fake_kb.df = [{"id": "A-1"}, {"id": "A-2"}]
        fake_kb.refresh_data.return_value = refresh_success

        fake_jobs = mock.Mock()
        fake_jobs.stats.return_value = {"jobs": {"queued": queued, "running": running}}

        state = connector_state or {
            "connector_exists": True,
            "enabled": True,
            "connector_path": "/tmp/internal_api_config.json",
            "resources": [],
            "recommended_endpoints": [],
        }

        return InternalApiService(
            knowledge_base=fake_kb,
            job_manager=fake_jobs,
            connector_settings_state=lambda: dict(state),
            build_connector_payload=lambda data, existing_payload=None: {"enabled": data.get("enabled", True)},
            load_connector_payload=lambda: {"enabled": True},
            write_connector_payload=mock.Mock(),
        ), fake_kb, fake_jobs

    def test_state_reports_active_api_connection_and_record_count(self):
        service, _, _ = self._service()

        state = service.get_state()

        self.assertEqual(state["project_data_source"], "api")
        self.assertEqual(state["active_runtime_source"], "company_api")
        self.assertTrue(state["api_connection_active"])
        self.assertTrue(state["can_refresh_dataset"])
        self.assertEqual(state["active_record_count"], 2)
        self.assertEqual(state["connection_label"], "Aktif memakai Internal API/APIDog")

    def test_can_mutate_blocks_when_report_jobs_are_active(self):
        service, _, _ = self._service(queued=1)

        allowed, response, status_code = service.can_mutate(
            "Pengaturan data internal belum bisa diubah karena masih ada laporan yang sedang diproses."
        )

        self.assertFalse(allowed)
        self.assertEqual(status_code, 409)
        self.assertEqual(response["status"], "busy")
        self.assertIn("laporan", response["error"])

    def test_refresh_dataset_requires_enabled_connector_before_refreshing(self):
        service, fake_kb, _ = self._service(connector_state={"connector_exists": False, "enabled": True})

        response, status_code = service.refresh_dataset()

        self.assertEqual(status_code, 400)
        self.assertEqual(response["status"], "not_configured")
        fake_kb.activate_internal_api_provider.assert_not_called()
        fake_kb.refresh_data.assert_not_called()

    def test_refresh_dataset_activates_provider_and_refreshes_data(self):
        service, fake_kb, _ = self._service(refresh_success=True)

        response, status_code = service.refresh_dataset()

        self.assertEqual(status_code, 200)
        self.assertEqual(response["status"], "refreshed")
        self.assertEqual(response["refresh_status"], "success")
        self.assertTrue(response["api_connection_active"])
        fake_kb.activate_internal_api_provider.assert_called_once()
        fake_kb.refresh_data.assert_called_once()

    def test_save_and_refresh_writes_payload_then_refreshes_knowledge_base(self):
        service, fake_kb, _ = self._service(refresh_success=False)

        response, status_code = service.save_and_refresh({"enabled": True})

        self.assertEqual(status_code, 200)
        self.assertEqual(response["status"], "saved")
        self.assertEqual(response["refresh_status"], "degraded")
        service.write_connector_payload.assert_called_once_with({"enabled": True})
        fake_kb.activate_internal_api_provider.assert_called_once()
        fake_kb.refresh_data.assert_called_once()


if __name__ == "__main__":
    unittest.main()

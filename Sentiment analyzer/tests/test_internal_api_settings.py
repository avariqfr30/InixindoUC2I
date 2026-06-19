import sys
import unittest
from unittest import mock
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from data_pipeline import InternalApiProvider
from class_report_adapter import ClassReportAdapter
from internal_api import EndpointSpec, InternalApiClient
from internal_api_settings import build_connector_payload, connector_settings_state, write_connector_payload
from internal_connector import InternalConnectorSpec


class FakeInternalApiClient:
    auth_mode = "api_key"
    auth_prefix = "Bearer"

    def interpret_payload(self, endpoint_spec):
        records = {
            "feedback_alpha": [
                {
                    "id": "A-1",
                    "stakeholder_type": "BUMN / Corporate",
                    "service_name": "Pelatihan Cloud",
                    "feedback_date": "2026-04-20",
                    "rating": 4,
                    "comment": "Materi relevan dan instruktur responsif.",
                }
            ],
            "feedback_beta": [
                {
                    "id": "B-1",
                    "stakeholder_type": "Instansi Pemerintah",
                    "service_name": "Audit SPBE",
                    "feedback_date": "2026-04-21",
                    "rating": 2,
                    "comment": "Follow up pasca sesi perlu lebih jelas.",
                }
            ],
        }[endpoint_spec.name]
        return {"records": records, "record_count": len(records)}


class FakeClassReportApiClient:
    auth_mode = "api_key"
    auth_prefix = "Bearer"

    def __init__(self):
        self.seen_specs = []

    def interpret_payload(self, endpoint_spec):
        self.seen_specs.append(endpoint_spec)
        records = {
            "class_report": [
                {
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "5",
                }
            ],
            "reference_class_report": [
                {
                    "class_start_date": "2026-04-01",
                    "class_end_date": "2026-04-03",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "",
                }
            ],
        }[endpoint_spec.name]
        return {"records": records, "record_count": len(records)}


class InternalApiSettingsTests(unittest.TestCase):
    def _connector_payload(self):
        return build_connector_payload(
            {
                "enabled": True,
                "context_enhancer": "Gunakan istilah internal secara konsisten.",
                "endpoints": [
                    {
                        "endpoint_name": "feedback_alpha",
                        "url": "https://internal.example/api/alpha",
                        "field_map": {
                            "id": "Record ID",
                            "stakeholder_type": "Tipe Stakeholder",
                            "service_name": "Layanan",
                            "feedback_date": "Tanggal Feedback",
                            "rating": "Rating",
                            "comment": "Komentar",
                        },
                    },
                    {
                        "endpoint_name": "feedback_beta",
                        "url": "https://internal.example/api/beta",
                        "field_map": {
                            "id": "Record ID",
                            "stakeholder_type": "Tipe Stakeholder",
                            "service_name": "Layanan",
                            "feedback_date": "Tanggal Feedback",
                            "rating": "Rating",
                            "comment": "Komentar",
                        },
                    },
                ],
            }
        )

    def test_build_connector_payload_keeps_multiple_endpoints(self):
        payload = self._connector_payload()
        connector = InternalConnectorSpec.from_mapping(payload)

        self.assertEqual(len(connector.endpoints), 2)
        self.assertEqual(len(connector.active_endpoints()), 2)
        self.assertEqual(connector.context_enhancer, "Gunakan istilah internal secara konsisten.")
        self.assertEqual(connector.endpoints[0].to_endpoint_spec().name, "feedback_alpha")

    def test_build_connector_payload_accepts_simple_apidog_setup(self):
        payload = build_connector_payload(
            {
                "url": "https://internal.example/api/Resource/dataset",
                "auth_mode": "bearer_env",
                "body_mode": "form",
            }
        )
        connector = InternalConnectorSpec.from_mapping(payload)

        self.assertEqual(payload["auth_mode"], "bearer_env")
        self.assertEqual(len(connector.endpoints), 2)
        self.assertEqual(
            [endpoint.request_data["dataset"] for endpoint in connector.endpoints],
            ["ClassReport", "ReferenceClassReport"],
        )
        self.assertEqual(connector.endpoints[0].endpoint_name, "class_report")
        self.assertEqual(connector.endpoints[0].method, "POST")
        self.assertEqual(connector.endpoints[0].body_mode, "form")
        self.assertEqual(connector.endpoints[0].record_path, "data.dataset_result")
        self.assertIn("class_name", connector.endpoints[0].field_map)
        self.assertIn("Komentar", connector.endpoints[0].field_map.values())

    def test_connector_settings_state_describes_recommended_endpoint_contracts(self):
        state = connector_settings_state(path="/tmp/non-existent-internal-api-config.json")
        recommended = state["recommended_endpoints"]

        self.assertEqual(
            [(item.get("name"), item.get("dataset")) for item in recommended],
            [
                ("class_report", "ClassReport"),
                ("reference_class_report", "ReferenceClassReport"),
            ],
        )
        self.assertNotIn("FeedbackDataset", str(recommended))
        for endpoint in recommended:
            self.assertIn("purpose", endpoint)
            self.assertIn("required_fields", endpoint)
            self.assertTrue(endpoint["required_fields"])
            self.assertEqual(endpoint["record_path"], "data.dataset_result")

    def test_connector_settings_state_reports_mapping_without_secrets(self):
        import tempfile

        payload = build_connector_payload(
            {
                "url": "https://internal.example/api/Resource/dataset",
                "auth_mode": "bearer_env",
                "body_mode": "form",
            }
        )
        payload["endpoints"][0]["headers"] = {"Authorization": "Bearer secret"}

        with tempfile.TemporaryDirectory() as temp_dir:
            path = str(Path(temp_dir) / "internal_api_config.json")
            write_connector_payload(payload, path=path)
            state = connector_settings_state(path=path)

        self.assertTrue(state["connector_exists"])
        self.assertEqual(state["project_data_source"], "api")
        self.assertEqual(state["connector_path"], path)
        self.assertEqual(state["auth_mode"], "bearer_env")
        self.assertEqual(state["resources"][0]["status"], "ok")
        self.assertNotIn("secret", str(state))

    def test_settings_template_exposes_refresh_without_requiring_resave(self):
        template = (PROJECT_DIR / "templates" / "index.html").read_text(encoding="utf-8")

        self.assertIn("Muat Ulang Data Sekarang", template)
        self.assertIn("btn-refresh-internal-api", template)
        self.assertIn("/api/internal-api/refresh", template)
        self.assertIn("setInternalApiConnectionState", template)

    def test_internal_api_refresh_endpoint_forces_provider_reload(self):
        import app as app_module

        fake_kb = mock.Mock()
        fake_kb.provider = mock.Mock(source_name="company_api")
        fake_kb.df = [{"id": "A-1"}]
        fake_kb.refresh_data.return_value = True
        fake_job_manager = mock.Mock()
        fake_job_manager.stats.return_value = {"jobs": {"queued": 0, "running": 0}}
        connector_state = {
            "connector_exists": True,
            "enabled": True,
            "connector_path": "/tmp/internal_api_config.json",
            "endpoints": [{"url": "https://internal.example/api/Resource/dataset"}],
            "resources": [],
            "recommended_endpoints": [],
        }

        with mock.patch.object(app_module, "current_user", return_value="tester"), \
            mock.patch.object(app_module, "kb", fake_kb), \
            mock.patch.object(app_module, "job_manager", fake_job_manager), \
            mock.patch.object(app_module, "connector_settings_state", return_value=connector_state):
            client = app_module.app.test_client()
            response = client.post("/api/internal-api/refresh")

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertEqual(payload["status"], "refreshed")
        self.assertTrue(payload["api_connection_active"])
        fake_kb.activate_internal_api_provider.assert_called_once()
        fake_kb.refresh_data.assert_called_once()

    def test_internal_api_provider_aggregates_multiple_connector_endpoints(self):
        provider = InternalApiProvider.__new__(InternalApiProvider)
        provider.client = FakeInternalApiClient()
        provider.connector = InternalConnectorSpec.from_mapping(self._connector_payload())

        dataframe = provider._load_via_connector()

        self.assertEqual(len(dataframe), 2)
        self.assertEqual(set(dataframe["Record ID"].tolist()), {"A-1", "B-1"})
        self.assertEqual(set(dataframe["Sumber Feedback"].tolist()), {"feedback_alpha", "feedback_beta"})
        self.assertTrue((dataframe["Komentar"].astype(str).str.len() > 0).all())

    def test_internal_api_provider_normalizes_apidog_class_report_shape(self):
        raw_df = __import__("pandas").DataFrame(
            [
                {
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "4",
                },
                {
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "2",
                },
                {
                    "response_id": "9",
                    "response_parent_id": "1",
                    "response_name": "Komentar material",
                    "response_type": "text",
                    "response_answer": "Materi perlu contoh praktik tambahan.",
                },
            ]
        )

        dataframe = ClassReportAdapter.normalize(raw_df, "class_report")

        self.assertEqual(len(dataframe), 1)
        self.assertEqual(dataframe.loc[0, "Rating"], "3")
        self.assertEqual(dataframe.loc[0, "Raw Response Count"], "3")
        self.assertEqual(dataframe.loc[0, "Rating Response Count"], "2")
        self.assertEqual(dataframe.loc[0, "Text Response Count"], "1")
        self.assertEqual(dataframe.loc[0, "Rating Distribution"], "2: 1; 4: 1")
        self.assertIn("Materi perlu contoh praktik tambahan", dataframe.loc[0, "Representative Why"])
        self.assertEqual(dataframe.loc[0, "Tipe Stakeholder"], "Peserta Kelas")
        self.assertEqual(dataframe.loc[0, "Layanan"], "Materi dan kurikulum")
        self.assertEqual(dataframe.loc[0, "Kanal Feedback"], "Evaluasi Kelas Internal")
        self.assertIn("Rata-rata rating Kesesuaian materi bahan ajar: 3.0 dari 5", dataframe.loc[0, "Komentar"])
        self.assertIn("Mengapa: Materi perlu contoh praktik tambahan", dataframe.loc[0, "Komentar"])
        self.assertNotIn("KESESUAIAN MATERIAL", dataframe.loc[0, "Komentar"])
        self.assertEqual(dataframe.loc[0, "Customer Journey Hint"], "Pelaksanaan Layanan")
        self.assertEqual(dataframe.loc[0, "Rentang Waktu"], "Semua Data APIDog (tanggal tidak tersedia)")
        self.assertTrue(dataframe["Tanggal Feedback"].astype(str).str.len().gt(0).all())

    def test_class_report_adapter_uses_reference_parent_when_class_child_parent_is_missing(self):
        pandas = __import__("pandas")
        class_report_df = pandas.DataFrame(
            [
                {
                    "response_id": "1",
                    "response_parent_id": None,
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "4",
                },
                {
                    "response_id": "9",
                    "response_parent_id": None,
                    "response_name": "Komentar material",
                    "response_type": "text",
                    "response_answer": "Materi perlu contoh praktik tambahan.",
                },
            ]
        )
        reference_df = pandas.DataFrame(
            [
                {
                    "class_start_date": "2026-04-01",
                    "class_end_date": "2026-04-03",
                    "response_id": "1",
                    "response_parent_id": None,
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                },
                {
                    "class_start_date": "2026-04-01",
                    "class_end_date": "2026-04-03",
                    "response_id": "9",
                    "response_parent_id": "1",
                    "response_name": "Komentar material",
                    "response_type": "text",
                },
            ]
        )

        dataframe = ClassReportAdapter.normalize(
            class_report_df,
            "class_report",
            reference_lookup=ClassReportAdapter.question_lookup(reference_df),
        )

        self.assertEqual(len(dataframe), 1)
        self.assertEqual(dataframe.loc[0, "Raw Response Count"], "2")
        self.assertEqual(dataframe.loc[0, "Rating Response Count"], "1")
        self.assertEqual(dataframe.loc[0, "Text Response Count"], "1")
        self.assertIn("Materi perlu contoh praktik tambahan", dataframe.loc[0, "Representative Why"])
        self.assertIn("Rata-rata rating Kesesuaian materi bahan ajar", dataframe.loc[0, "Komentar"])
        self.assertNotIn("Komentar material:", dataframe.loc[0, "Komentar"])
        self.assertEqual(dataframe.loc[0, "Layanan"], "Materi dan kurikulum")

    def test_internal_api_provider_uses_reference_class_report_as_dictionary_not_feedback_rows(self):
        provider = InternalApiProvider.__new__(InternalApiProvider)
        provider.client = FakeClassReportApiClient()
        provider.connector = InternalConnectorSpec.from_mapping(
            build_connector_payload(
                {
                    "url": "https://internal.example/api/Resource/dataset",
                    "auth_mode": "none",
                    "body_mode": "form",
                    "request_data": '{"dataset_cache": "enabled"}',
                }
            )
        )

        dataframe = provider._load_via_connector()

        self.assertEqual(len(dataframe), 1)
        self.assertEqual(dataframe.loc[0, "Raw Response Count"], "1")
        self.assertEqual(dataframe.loc[0, "Rating Response Count"], "1")
        self.assertEqual(dataframe.loc[0, "Layanan"], "Materi dan kurikulum")
        self.assertEqual(dataframe.loc[0, "Tanggal Feedback"], "2026-04-03")
        self.assertEqual(dataframe.loc[0, "Rentang Waktu"], "2026-04-01 sampai 2026-04-03")
        self.assertIn("Rata-rata rating Kesesuaian materi bahan ajar: 5.0 dari 5", dataframe.loc[0, "Komentar"])
        self.assertNotIn("reference_class_report", set(dataframe["Sumber Feedback"]))

        reference_specs = [spec for spec in provider.client.seen_specs if spec.name == "reference_class_report"]
        self.assertEqual(reference_specs[0].query_params["dataset_cache"], "enabled")

    def test_reference_class_report_detection_allows_response_answer_column_when_endpoint_is_reference(self):
        reference_df = __import__("pandas").DataFrame(
            [
                {
                    "class_start_date": "2024-07-01",
                    "class_end_date": "2024-07-02",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "",
                }
            ]
        )

        self.assertTrue(
            ClassReportAdapter.looks_like_reference_class_report(
                reference_df,
                endpoint_name="reference_class_report",
                dataset_code="ReferenceClassReport",
            )
        )

        lookup = ClassReportAdapter.question_lookup(reference_df)
        self.assertEqual(lookup["1"]["class_start_dates"], ["2024-07-01"])
        self.assertEqual(lookup["1"]["class_end_dates"], ["2024-07-02"])

    def test_reference_class_report_lookup_retains_parent_child_relationships(self):
        reference_df = __import__("pandas").DataFrame(
            [
                {
                    "class_start_date": "2024-07-01",
                    "class_end_date": "2024-07-02",
                    "response_id": "1",
                    "response_parent_id": None,
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                },
                {
                    "class_start_date": "2024-07-01",
                    "class_end_date": "2024-07-02",
                    "response_id": "9",
                    "response_parent_id": "1",
                    "response_name": "Komentar material",
                    "response_type": "text",
                },
            ]
        )

        lookup = ClassReportAdapter.question_lookup(reference_df)

        self.assertEqual(lookup["1"]["parent_id"], "")
        self.assertEqual(lookup["9"]["parent_id"], "1")
        self.assertEqual(lookup["1"]["label"], "Kesesuaian materi bahan ajar")
        self.assertEqual(lookup["9"]["label"], "Komentar material")
        self.assertEqual(lookup["9"]["class_start_dates"], ["2024-07-01"])
        self.assertEqual(lookup["9"]["class_end_dates"], ["2024-07-02"])

    def test_reference_class_report_with_actual_answers_is_processed_as_date_bearing_feedback(self):
        reference_df = __import__("pandas").DataFrame(
            [
                {
                    "class_start_date": "2024-07-01",
                    "class_end_date": "2024-07-02",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "5",
                }
            ]
        )

        self.assertFalse(
            ClassReportAdapter.looks_like_reference_class_report(
                reference_df,
                endpoint_name="reference_class_report",
                dataset_code="ReferenceClassReport",
            )
        )
        self.assertTrue(ClassReportAdapter.looks_like_class_report(reference_df))

    def test_class_report_adapter_preserves_reference_class_dates(self):
        class_report_df = __import__("pandas").DataFrame(
            [
                {
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "4",
                }
            ]
        )
        reference_df = __import__("pandas").DataFrame(
            [
                {
                    "class_start_date": "2026-04-01",
                    "class_end_date": "2026-04-03",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                }
            ]
        )

        dataframe = ClassReportAdapter.normalize(
            class_report_df,
            "class_report",
            reference_lookup=ClassReportAdapter.question_lookup(reference_df),
        )

        self.assertEqual(dataframe.loc[0, "Tanggal Feedback"], "2026-04-03")
        self.assertEqual(dataframe.loc[0, "Rentang Waktu"], "2026-04-01 sampai 2026-04-03")

    def test_class_report_adapter_preserves_dates_from_class_report_rows(self):
        raw_df = __import__("pandas").DataFrame(
            [
                {
                    "class_start_date": "2026-05-10 09:00:00",
                    "class_end_date": "2026-05-12 16:30:00",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "5",
                },
                {
                    "class_start_date": "2026-05-10 09:00:00",
                    "class_end_date": "2026-05-12 16:30:00",
                    "response_id": "8",
                    "response_parent_id": "1",
                    "response_name": "Komentar material",
                    "response_type": "text",
                    "response_answer": "Materi sangat relevan.",
                },
            ]
        )

        dataframe = ClassReportAdapter.normalize(raw_df, "class_report")

        self.assertEqual(dataframe.loc[0, "Tanggal Feedback"], "2026-05-12")
        self.assertEqual(dataframe.loc[0, "Rentang Waktu"], "2026-05-10 sampai 2026-05-12")
        self.assertNotIn("tanggal tidak tersedia", dataframe.loc[0, "Rentang Waktu"].lower())

    def test_class_report_adapter_groups_repeated_question_by_class_date_window(self):
        raw_df = __import__("pandas").DataFrame(
            [
                {
                    "class_start_date": "2024-07-01",
                    "class_end_date": "2024-07-02",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "4",
                },
                {
                    "class_start_date": "2025-01-06",
                    "class_end_date": "2025-01-07",
                    "response_id": "1",
                    "response_parent_id": "",
                    "response_name": "KESESUAIAN MATERIAL BAHAN AJAR",
                    "response_type": "rating_5",
                    "response_answer": "2",
                },
                {
                    "class_start_date": "2025-01-06",
                    "class_end_date": "2025-01-07",
                    "response_id": "9",
                    "response_parent_id": "1",
                    "response_name": "Komentar material",
                    "response_type": "text",
                    "response_answer": "Materi perlu contoh praktik tambahan.",
                },
            ]
        )

        dataframe = ClassReportAdapter.normalize(raw_df, "class_report")

        self.assertEqual(len(dataframe), 2)
        self.assertEqual(set(dataframe["Tanggal Feedback"]), {"2024-07-02", "2025-01-07"})
        latest = dataframe[dataframe["Tanggal Feedback"] == "2025-01-07"].iloc[0]
        self.assertEqual(latest["Rating"], "2")
        self.assertIn("Materi perlu contoh praktik tambahan", latest["Representative Why"])

    def test_connector_auth_mode_updates_runtime_client(self):
        provider = InternalApiProvider.__new__(InternalApiProvider)
        provider.client = FakeInternalApiClient()
        provider.connector = InternalConnectorSpec.from_mapping(
            {
                **self._connector_payload(),
                "auth_mode": "none",
            }
        )

        provider._load_via_connector()

        self.assertEqual(provider.client.auth_mode, "none")

    def test_internal_api_client_none_auth_does_not_send_environment_key(self):
        client = InternalApiClient(
            api_key="secret-token",
            auth_mode="none",
            default_headers={},
        )
        headers = client._build_headers(EndpointSpec(name="feedback", path="/feedback"))

        self.assertNotIn("Authorization", headers)
        self.assertNotIn("X-API-Key", headers)

    def test_internal_api_client_supports_apidog_multipart_body(self):
        client = InternalApiClient(auth_mode="none", default_headers={})
        endpoint = EndpointSpec(
            name="feedback",
            path="https://internal.example/api/Resource/dataset",
            method="POST",
            body_mode="multipart",
            query_params={"dataset": "ClassReport", "dataset_cache": "enabled"},
        )
        response = mock.Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {"success": True, "data": {"dataset_result": []}}

        with mock.patch("internal_api.requests.request", return_value=response) as request_mock:
            client.request_endpoint(endpoint)

        _, kwargs = request_mock.call_args
        self.assertEqual(kwargs["files"]["dataset"], (None, "ClassReport"))
        self.assertEqual(kwargs["files"]["dataset_cache"], (None, "enabled"))
        self.assertNotIn("data", kwargs)
        self.assertNotIn("json", kwargs)

    def test_build_connector_payload_rejects_non_http_endpoint(self):
        with self.assertRaisesRegex(ValueError, "HTTP/HTTPS"):
            build_connector_payload({"endpoints": [{"url": "not-a-url"}]})

    def test_blank_header_edit_preserves_existing_endpoint_headers(self):
        existing_payload = self._connector_payload()
        existing_payload["endpoints"][0]["headers"] = {"Authorization": "Bearer secret"}

        updated_payload = build_connector_payload(
            {
                "endpoints": [
                    {
                        "endpoint_name": "feedback_alpha",
                        "url": "https://internal.example/api/alpha-updated",
                        "headers": "",
                    }
                ]
            },
            existing_payload=existing_payload,
        )

        self.assertEqual(
            updated_payload["endpoints"][0]["headers"],
            {"Authorization": "Bearer secret"},
        )


if __name__ == "__main__":
    unittest.main()

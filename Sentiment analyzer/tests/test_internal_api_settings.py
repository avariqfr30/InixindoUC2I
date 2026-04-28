import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from data_pipeline import InternalApiProvider
from internal_api_settings import build_connector_payload
from internal_connector import InternalConnectorSpec


class FakeInternalApiClient:
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

    def test_internal_api_provider_aggregates_multiple_connector_endpoints(self):
        provider = InternalApiProvider.__new__(InternalApiProvider)
        provider.client = FakeInternalApiClient()
        provider.connector = InternalConnectorSpec.from_mapping(self._connector_payload())

        dataframe = provider._load_via_connector()

        self.assertEqual(len(dataframe), 2)
        self.assertEqual(set(dataframe["Record ID"].tolist()), {"A-1", "B-1"})
        self.assertEqual(set(dataframe["Sumber Feedback"].tolist()), {"feedback_alpha", "feedback_beta"})
        self.assertTrue((dataframe["Komentar"].astype(str).str.len() > 0).all())

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

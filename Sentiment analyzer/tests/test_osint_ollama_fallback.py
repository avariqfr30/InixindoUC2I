import sys
import unittest
import warnings
from pathlib import Path
from unittest import mock

warnings.filterwarnings("ignore", message="urllib3 v2 only supports OpenSSL.*")

import requests

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import osint_research
from osint_research import Researcher


class FakeResponse:
    def __init__(self, payload, status_error=None):
        self.payload = payload
        self.status_error = status_error

    def raise_for_status(self):
        if self.status_error:
            raise self.status_error

    def json(self):
        return self.payload


class OllamaWebSearchFallbackTests(unittest.TestCase):
    def test_missing_serper_key_uses_ollama_web_search_results(self):
        ollama_payload = {
            "results": [
                {
                    "title": "Indonesia IT training benchmark rises",
                    "url": "https://example.org/it-training-benchmark",
                    "content": "Customer expectations for IT training in Indonesia are shifting toward measurable outcomes.",
                }
            ]
        }

        def fake_post(url, **kwargs):
            self.assertEqual(url, "https://ollama.example/api/web_search")
            self.assertEqual(kwargs["headers"]["Authorization"], "Bearer test-ollama-key")
            self.assertEqual(kwargs["json"]["query"], "benchmark pelatihan it indonesia")
            return FakeResponse(ollama_payload)

        with (
            mock.patch.object(osint_research, "SERPER_API_KEY", "YOUR_SERPER_API_KEY"),
            mock.patch.object(osint_research, "OLLAMA_API_KEY", "test-ollama-key", create=True),
            mock.patch.object(osint_research, "OLLAMA_WEB_SEARCH_URL", "https://ollama.example/api/web_search", create=True),
            mock.patch.object(osint_research.requests, "post", side_effect=fake_post),
        ):
            results = Researcher._run_query_batch(["benchmark pelatihan IT Indonesia"], max_signals=3)

        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["title"], "Indonesia IT training benchmark rises")
        self.assertEqual(results[0]["snippet"], ollama_payload["results"][0]["content"])
        self.assertEqual(results[0]["url"], "https://example.org/it-training-benchmark")
        self.assertEqual(results[0]["source_type"], "organic")

    def test_serper_failure_falls_back_to_ollama_web_search(self):
        ollama_payload = {
            "results": [
                {
                    "title": "Consulting customer experience benchmark",
                    "url": "https://trusted.org/customer-experience-benchmark",
                    "content": "Consulting buyers expect tighter follow-up loops and clearer service recovery.",
                }
            ]
        }
        calls = []

        def fake_post(url, **kwargs):
            calls.append(url)
            if url == Researcher.SERPER_ENDPOINT:
                return FakeResponse({}, status_error=requests.HTTPError("serper unavailable"))
            self.assertEqual(url, "https://ollama.example/api/web_search")
            return FakeResponse(ollama_payload)

        with (
            mock.patch.object(osint_research, "SERPER_API_KEY", "real-serper-key"),
            mock.patch.object(osint_research, "OLLAMA_API_KEY", "test-ollama-key", create=True),
            mock.patch.object(osint_research, "OLLAMA_WEB_SEARCH_URL", "https://ollama.example/api/web_search", create=True),
            mock.patch.object(osint_research.requests, "post", side_effect=fake_post),
            mock.patch.object(osint_research.logger, "warning"),
        ):
            results = Researcher._run_query_batch(["customer experience consulting benchmark"], max_signals=3)

        self.assertEqual(calls, [Researcher.SERPER_ENDPOINT, "https://ollama.example/api/web_search"])
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["title"], "Consulting customer experience benchmark")
        self.assertEqual(results[0]["snippet"], ollama_payload["results"][0]["content"])
        self.assertEqual(results[0]["url"], "https://trusted.org/customer-experience-benchmark")


if __name__ == "__main__":
    unittest.main()

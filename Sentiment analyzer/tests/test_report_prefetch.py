import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


class ReportPrefetchTests(unittest.TestCase):
    def test_warm_report_context_uses_existing_osint_cache_path(self):
        import app as app_module

        class RecordingExecutor:
            def __init__(self):
                self.calls = []

            def submit(self, *args):
                self.calls.append(args)
                return object()

        executor = RecordingExecutor()
        original_executor = app_module.prefetch_executor
        try:
            app_module.prefetch_executor = executor
            result = app_module.warm_report_context(
                {
                    "timeframe": "1 Bulan Terakhir (Monthly)",
                    "notes": "Periksa tindak lanjut layanan.",
                    "sentiment": "negative",
                    "segment": "Peserta Kelas",
                    "score_engine": "experience_index",
                }
            )
        finally:
            app_module.prefetch_executor = original_executor

        self.assertEqual(result["status"], "warming")
        self.assertEqual(result["timeframe"], "1 Bulan Terakhir (Monthly)")
        self.assertEqual(result["score_engine"], "experience_index")
        self.assertEqual(len(executor.calls), 1)
        submitted = executor.calls[0]
        self.assertEqual(submitted[0].__func__, app_module.Researcher.get_macro_trends.__func__)
        self.assertEqual(submitted[1:], ("1 Bulan Terakhir (Monthly)", "Periksa tindak lanjut layanan.", "Experience Index"))

    def test_warm_report_context_requires_timeframe(self):
        import app as app_module

        with self.assertRaises(ValueError):
            app_module.warm_report_context({"notes": "Tanpa periode."})


if __name__ == "__main__":
    unittest.main()

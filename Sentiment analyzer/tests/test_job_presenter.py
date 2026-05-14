import sys
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


class JobPresenterTests(unittest.TestCase):
    def test_queued_job_includes_user_ready_queue_state(self):
        from job_presenter import present_job

        presented = present_job(
            {
                "job_id": "job-1",
                "status": "queued",
                "created_at": "2026-05-14T03:00:00.000Z",
                "queue_position": 3,
            },
            stats={"jobs": {"queued": 3, "running": 1}},
            now="2026-05-14T03:00:42.000Z",
        )

        self.assertEqual(presented["stage_label"], "Menunggu giliran")
        self.assertEqual(presented["stage_detail"], "Posisi antrian 3. 1 laporan sedang diproses.")
        self.assertFalse(presented["can_download"])
        self.assertFalse(presented["retryable"])
        self.assertEqual(presented["queue_position"], 3)
        self.assertEqual(presented["elapsed_seconds"], 42)

    def test_running_job_reports_processing_stage_and_elapsed_time(self):
        from job_presenter import present_job

        presented = present_job(
            {
                "job_id": "job-2",
                "status": "running",
                "created_at": "2026-05-14T03:00:00.000Z",
                "started_at": "2026-05-14T03:00:10.000Z",
                "queue_wait_seconds": 10.2,
                "queue_position": 0,
            },
            now="2026-05-14T03:01:10.000Z",
        )

        self.assertEqual(presented["stage_label"], "Sedang menyusun laporan")
        self.assertEqual(presented["stage_detail"], "Diproses selama 60 detik setelah menunggu 10 detik.")
        self.assertFalse(presented["can_download"])
        self.assertFalse(presented["retryable"])
        self.assertEqual(presented["queue_position"], 0)
        self.assertEqual(presented["elapsed_seconds"], 70)

    def test_completed_job_is_downloadable_and_uses_recorded_duration(self):
        from job_presenter import present_job

        presented = present_job(
            {
                "job_id": "job-3",
                "status": "completed",
                "created_at": "2026-05-14T03:00:00.000Z",
                "completed_at": "2026-05-14T03:02:00.000Z",
                "total_elapsed_seconds": 120.4,
                "filename": "report.docx",
                "queue_position": 0,
            },
            now="2026-05-14T03:05:00.000Z",
        )

        self.assertEqual(presented["stage_label"], "Laporan siap diunduh")
        self.assertEqual(presented["stage_detail"], "report.docx selesai dalam 120 detik.")
        self.assertTrue(presented["can_download"])
        self.assertFalse(presented["retryable"])
        self.assertEqual(presented["elapsed_seconds"], 120)

    def test_failed_job_is_retryable_with_safe_error_detail(self):
        from job_presenter import present_job

        presented = present_job(
            {
                "job_id": "job-4",
                "status": "failed",
                "created_at": "2026-05-14T03:00:00.000Z",
                "completed_at": "2026-05-14T03:00:30.000Z",
                "total_elapsed_seconds": 30,
                "error": "Traceback: internal stack with /secret/path",
                "queue_position": 0,
            },
            now="2026-05-14T03:01:00.000Z",
        )

        self.assertEqual(presented["stage_label"], "Laporan gagal dibuat")
        self.assertEqual(presented["stage_detail"], "Silakan coba lagi. Detail teknis tersimpan di log server.")
        self.assertFalse(presented["can_download"])
        self.assertTrue(presented["retryable"])
        self.assertEqual(presented["elapsed_seconds"], 30)


if __name__ == "__main__":
    unittest.main()

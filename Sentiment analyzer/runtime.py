import concurrent.futures
import json
import logging
import os
import threading
import time
import uuid
from datetime import datetime, timedelta

from config import (
    JOB_STATE_PATH,
    REPORT_ARTIFACT_DIR,
    REPORT_JOB_RETENTION_SECONDS,
    REPORT_JOB_WORKERS,
    REPORT_MAX_PENDING_JOBS,
)

logger = logging.getLogger(__name__)


class QueueCapacityError(RuntimeError):
    pass


class ReportJobManager:
    ACTIVE_STATUSES = {"queued", "running"}

    def __init__(
        self,
        generator,
        artifact_dir=REPORT_ARTIFACT_DIR,
        state_path=JOB_STATE_PATH,
        max_workers=REPORT_JOB_WORKERS,
        max_pending_jobs=REPORT_MAX_PENDING_JOBS,
        retention_seconds=REPORT_JOB_RETENTION_SECONDS,
    ):
        self.generator = generator
        self.artifact_dir = artifact_dir
        self.state_path = state_path
        self.max_workers = max_workers
        self.max_pending_jobs = max_pending_jobs
        self.retention_seconds = retention_seconds
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
        self.lock = threading.Lock()
        self.jobs = {}
        self.futures = {}

        os.makedirs(self.artifact_dir, exist_ok=True)
        self._load_state()
        self._cleanup_expired()

    def _load_state(self):
        if not os.path.exists(self.state_path):
            self.jobs = {}
            return

        try:
            with open(self.state_path, "r", encoding="utf-8") as handle:
                raw_state = json.load(handle)
            if isinstance(raw_state, dict):
                self.jobs = raw_state
            else:
                self.jobs = {}
        except Exception as exc:
            logger.warning("Failed to load report job state: %s", exc)
            self.jobs = {}

    def _persist_state(self):
        os.makedirs(os.path.dirname(self.state_path), exist_ok=True)
        temp_path = f"{self.state_path}.tmp"
        with open(temp_path, "w", encoding="utf-8") as handle:
            json.dump(self.jobs, handle, ensure_ascii=True, indent=2, sort_keys=True)
        os.replace(temp_path, self.state_path)

    @staticmethod
    def _utc_now():
        return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

    def _expired_job_ids(self):
        cutoff = datetime.utcnow() - timedelta(seconds=self.retention_seconds)
        expired_ids = []
        for job_id, job in self.jobs.items():
            completed_at = job.get("completed_at") or job.get("updated_at") or job.get("created_at")
            if not completed_at:
                continue
            try:
                completed_dt = datetime.fromisoformat(completed_at.replace("Z", "+00:00")).replace(tzinfo=None)
            except ValueError:
                continue
            if completed_dt < cutoff and job.get("status") not in self.ACTIVE_STATUSES:
                expired_ids.append(job_id)
        return expired_ids

    def _cleanup_expired(self):
        with self.lock:
            expired_ids = self._expired_job_ids()
            if not expired_ids:
                return

            for job_id in expired_ids:
                artifact_path = self.jobs.get(job_id, {}).get("artifact_path")
                if artifact_path and os.path.exists(artifact_path):
                    try:
                        os.remove(artifact_path)
                    except OSError:
                        logger.warning("Failed to remove expired artifact for job %s.", job_id)
                self.jobs.pop(job_id, None)
                self.futures.pop(job_id, None)

            self._persist_state()

    def _active_job_count(self):
        return sum(1 for job in self.jobs.values() if job.get("status") in self.ACTIVE_STATUSES)

    def _queue_position(self, job_id):
        queued_job_ids = [
            current_id
            for current_id, job in self.jobs.items()
            if job.get("status") == "queued"
        ]
        if job_id not in queued_job_ids:
            return 0
        return queued_job_ids.index(job_id) + 1

    def _build_job_record(self, payload):
        job_id = uuid.uuid4().hex
        created_at = self._utc_now()
        return {
            "job_id": job_id,
            "status": "queued",
            "created_at": created_at,
            "updated_at": created_at,
            "started_at": None,
            "completed_at": None,
            "duration_seconds": None,
            "error": None,
            "filename": None,
            "artifact_path": None,
            "artifact_size_bytes": None,
            "input": {
                "timeframe": payload["timeframe"],
                "sentiment": payload.get("sentiment", "all"),
                "segment": payload.get("segment", "all"),
                "score_engine": payload.get("score_engine"),
            },
        }

    def submit(self, payload):
        self._cleanup_expired()

        with self.lock:
            if self._active_job_count() >= self.max_pending_jobs:
                raise QueueCapacityError(
                    "Antrian laporan sedang penuh. Silakan coba lagi beberapa saat lagi."
                )

            job = self._build_job_record(payload)
            self.jobs[job["job_id"]] = job
            self._persist_state()
            self.futures[job["job_id"]] = self.executor.submit(
                self._run_job,
                job["job_id"],
                payload,
            )
            return self._public_job(job["job_id"])

    def _run_job(self, job_id, payload):
        start_time = time.perf_counter()
        with self.lock:
            job = self.jobs[job_id]
            job["status"] = "running"
            job["started_at"] = self._utc_now()
            job["updated_at"] = job["started_at"]
            self._persist_state()

        try:
            document, filename = self.generator.run(
                payload["timeframe"],
                payload.get("notes", ""),
                sentiment=payload.get("sentiment", "all"),
                segment=payload.get("segment", "all"),
                score_engine=payload.get("score_engine"),
            )
            artifact_path = os.path.join(self.artifact_dir, f"{job_id}.docx")
            document.save(artifact_path)

            with self.lock:
                job = self.jobs[job_id]
                job["status"] = "completed"
                job["completed_at"] = self._utc_now()
                job["updated_at"] = job["completed_at"]
                job["duration_seconds"] = round(time.perf_counter() - start_time, 2)
                job["filename"] = f"{filename}.docx"
                job["artifact_path"] = artifact_path
                job["artifact_size_bytes"] = os.path.getsize(artifact_path)
                job["error"] = None
                self._persist_state()
        except Exception as exc:
            logger.exception("Report job %s failed.", job_id)
            with self.lock:
                job = self.jobs[job_id]
                job["status"] = "failed"
                job["completed_at"] = self._utc_now()
                job["updated_at"] = job["completed_at"]
                job["duration_seconds"] = round(time.perf_counter() - start_time, 2)
                job["error"] = str(exc)
                self._persist_state()

    def _public_job(self, job_id):
        job = self.jobs.get(job_id)
        if not job:
            return None

        job_copy = dict(job)
        artifact_path = job_copy.pop("artifact_path", None)
        if artifact_path and not os.path.exists(artifact_path):
            job_copy["status"] = "failed"
            job_copy["error"] = "Berkas laporan tidak lagi tersedia."

        job_copy["queue_position"] = self._queue_position(job_id)
        return job_copy

    def get(self, job_id):
        self._cleanup_expired()
        with self.lock:
            return self._public_job(job_id)

    def artifact_for(self, job_id):
        with self.lock:
            job = self.jobs.get(job_id)
            if not job or job.get("status") != "completed":
                return None
            artifact_path = job.get("artifact_path")
            if not artifact_path or not os.path.exists(artifact_path):
                return None
            return {
                "path": artifact_path,
                "filename": job.get("filename") or f"{job_id}.docx",
            }

    def stats(self):
        self._cleanup_expired()
        with self.lock:
            status_counts = {
                "queued": 0,
                "running": 0,
                "completed": 0,
                "failed": 0,
            }
            for job in self.jobs.values():
                status = job.get("status")
                if status in status_counts:
                    status_counts[status] += 1

            writable = os.access(self.artifact_dir, os.W_OK)
            return {
                "max_workers": self.max_workers,
                "max_pending_jobs": self.max_pending_jobs,
                "artifact_dir": self.artifact_dir,
                "artifact_dir_writable": writable,
                "jobs": status_counts,
            }

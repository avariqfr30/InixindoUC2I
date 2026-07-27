from datetime import datetime, timezone


def _parse_timestamp(value):
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def _now(value=None):
    parsed = _parse_timestamp(value)
    return parsed or datetime.now(timezone.utc)


def _duration_seconds(value):
    if value is None:
        return None
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def _elapsed_from(job, now=None):
    explicit = _duration_seconds(job.get("total_elapsed_seconds") or job.get("duration_seconds"))
    if explicit is not None:
        return explicit
    created_at = _parse_timestamp(job.get("created_at"))
    if not created_at:
        return None
    return max(0, int(round((_now(now) - created_at).total_seconds())))



def _quality_summary(value):
    if not isinstance(value, dict):
        return None
    preflight = value.get("preflight") if isinstance(value.get("preflight"), dict) else {}
    summary = {
        "verification_status": value.get("verification_status"),
        "verified_complete": bool(value.get("verified_complete")),
        "completeness_score": value.get("completeness_score"),
        "passed_checks": value.get("passed_checks"),
        "total_checks": value.get("total_checks"),
        "missing_checks": value.get("missing_checks") or [],
        "preflight_passes": bool(preflight.get("passes")),
    }
    return {key: item for key, item in summary.items() if item is not None}


def _public_error(status):
    if str(status or "").lower() != "failed":
        return None
    return "Laporan belum berhasil dibuat. Silakan coba lagi atau hubungi admin jika berulang."


def _running_detail(job, now=None):
    started_at = _parse_timestamp(job.get("started_at"))
    processed_seconds = None
    if started_at:
        processed_seconds = max(0, int(round((_now(now) - started_at).total_seconds())))
    wait_seconds = _duration_seconds(job.get("queue_wait_seconds"))
    if processed_seconds is not None and wait_seconds is not None:
        return f"Diproses selama {processed_seconds} detik setelah menunggu {wait_seconds} detik."
    if processed_seconds is not None:
        return f"Diproses selama {processed_seconds} detik."
    return "Sistem sedang menganalisis data, menarik konteks, dan menyusun dokumen."


def present_job(job, stats=None, now=None, status_url=None, download_url=None):
    response = dict(job or {})
    internal_quality = response.pop("quality", None)
    status = str(response.get("status") or "unknown").lower()
    if status == "failed":
        response["error"] = _public_error(status)
    elif "error" in response and not response.get("error"):
        response.pop("error", None)
    queue_position = int(response.get("queue_position") or 0)
    elapsed_seconds = _elapsed_from(response, now=now)

    if status == "queued":
        running_count = int(((stats or {}).get("jobs") or {}).get("running") or 0)
        stage_label = "Menunggu giliran"
        stage_detail = f"Posisi antrian {queue_position}. {running_count} laporan sedang diproses."
    elif status == "running":
        stage_label = "Sedang menyusun laporan"
        stage_detail = _running_detail(response, now=now)
    elif status == "completed":
        stage_label = "Laporan siap diunduh"
        filename = response.get("filename") or "Dokumen laporan"
        if elapsed_seconds is not None:
            stage_detail = f"{filename} selesai dalam {elapsed_seconds} detik."
        else:
            stage_detail = f"{filename} selesai dibuat."
    elif status == "failed":
        stage_label = "Laporan gagal dibuat"
        stage_detail = "Silakan coba lagi. Detail teknis tersimpan di log server."
    else:
        stage_label = "Status laporan tidak dikenal"
        stage_detail = "Sistem menerima status pekerjaan yang belum dikenali."

    response.update(
        {
            "stage_label": stage_label,
            "stage_detail": stage_detail,
            "can_download": status == "completed",
            "retryable": status == "failed",
            "queue_position": queue_position,
            "elapsed_seconds": elapsed_seconds,
        }
    )
    public_quality = _quality_summary(internal_quality)
    if public_quality and status == "completed":
        response["quality_summary"] = public_quality
    if status_url:
        response["status_url"] = status_url
    if download_url and response["can_download"]:
        response["download_url"] = download_url
    return response

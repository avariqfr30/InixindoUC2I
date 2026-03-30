import io
import logging
import os
from flask import Flask, jsonify, render_template, request, send_file
from flask_cors import CORS

from config import (
    APP_MODE,
    DB_URI,
    DEFAULT_SCORE_ENGINE,
    SCORE_ENGINE_OPTIONS,
    SENTIMENT_OPTIONS,
    SMART_SUGGESTIONS,
)
from core import KnowledgeBase, ReportGenerator
from runtime import QueueCapacityError, ReportJobManager

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

kb = KnowledgeBase(DB_URI)
generator = ReportGenerator(kb)
job_manager = ReportJobManager(generator)
logger.info("Application started in %s mode.", APP_MODE)


def _request_payload(data):
    payload = {
        "timeframe": data.get("timeframe"),
        "notes": data.get("notes", ""),
        "sentiment": data.get("sentiment", "all"),
        "segment": data.get("segment", "all"),
        "score_engine": data.get("score_engine", DEFAULT_SCORE_ENGINE),
    }
    return payload

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/health")
def health():
    data_ready = kb.df is not None and not kb.df.empty
    status_code = 200 if data_ready else 503
    return (
        jsonify(
            {
                "status": "ok" if data_ready else "degraded",
                "data_ready": data_ready,
            }
        ),
        status_code,
    )


@app.route("/ready")
def ready():
    data_ready = kb.df is not None and not kb.df.empty
    job_stats = job_manager.stats()
    ready_state = data_ready and job_stats["artifact_dir_writable"]
    status_code = 200 if ready_state else 503
    return (
        jsonify(
            {
                "status": "ready" if ready_state else "not_ready",
                "data_ready": data_ready,
                "artifact_dir_writable": job_stats["artifact_dir_writable"],
                "job_capacity": {
                    "max_workers": job_stats["max_workers"],
                    "max_pending_jobs": job_stats["max_pending_jobs"],
                    "queued": job_stats["jobs"]["queued"],
                    "running": job_stats["jobs"]["running"],
                },
            }
        ),
        status_code,
    )


@app.route("/get-config")
def get_config():
    if kb.df is None or kb.df.empty:
        return jsonify(
            {
                "error": (
                    "Internal data is not available. "
                    "Check the server-side data source configuration."
                )
            }
        )

    timeframes = sorted(kb.df["Rentang Waktu"].dropna().unique().tolist())
    segments = sorted(
        value
        for value in kb.df["Tipe Stakeholder"].fillna("").astype(str).str.strip().unique().tolist()
        if value
    )

    return jsonify(
        {
            "timeframes": timeframes,
            "sentiments": SENTIMENT_OPTIONS,
            "segments": [{"id": "all", "label": "Semua Segmen"}]
            + [{"id": segment, "label": segment} for segment in segments],
            "score_engines": SCORE_ENGINE_OPTIONS,
            "default_score_engine": DEFAULT_SCORE_ENGINE,
            "suggestions": SMART_SUGGESTIONS,
        }
    )


@app.route("/generate", methods=["POST"])
def generate_report():
    data = request.get_json(silent=True) or {}
    payload = _request_payload(data)
    timeframe = payload["timeframe"]

    if not timeframe:
        return jsonify({"error": "Parameter 'timeframe' wajib diisi."}), 400

    logger.info(
        "Generating report for timeframe='%s', sentiment='%s', segment='%s', score_engine='%s'.",
        timeframe,
        payload["sentiment"],
        payload["segment"],
        payload["score_engine"],
    )
    doc, filename, quality = generator.run(
        timeframe,
        payload["notes"],
        sentiment=payload["sentiment"],
        segment=payload["segment"],
        score_engine=payload["score_engine"],
    )

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    response = send_file(
        out,
        as_attachment=True,
        download_name=f"{filename}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response.headers["X-Report-Completeness"] = str(quality["completeness_score"])
    response.headers["X-Report-Verified"] = str(quality["verified_complete"]).lower()
    return response


@app.route("/generate-job", methods=["POST"])
def generate_report_job():
    data = request.get_json(silent=True) or {}
    payload = _request_payload(data)

    if not payload["timeframe"]:
        return jsonify({"error": "Parameter 'timeframe' wajib diisi."}), 400

    try:
        job = job_manager.submit(payload)
    except QueueCapacityError as exc:
        return jsonify({"error": str(exc)}), 429

    logger.info(
        "Queued report job %s for timeframe='%s', sentiment='%s', segment='%s', score_engine='%s'.",
        job["job_id"],
        payload["timeframe"],
        payload["sentiment"],
        payload["segment"],
        payload["score_engine"],
    )
    return (
        jsonify(
            {
                **job,
                "status_url": f"/jobs/{job['job_id']}",
                "download_url": f"/download/{job['job_id']}",
            }
        ),
        202,
    )


@app.route("/jobs/<job_id>")
def get_report_job(job_id):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"error": "Job tidak ditemukan."}), 404

    response = dict(job)
    response["status_url"] = f"/jobs/{job_id}"
    if job["status"] == "completed":
        response["download_url"] = f"/download/{job_id}"
    return jsonify(response)


@app.route("/download/<job_id>")
def download_report(job_id):
    artifact = job_manager.artifact_for(job_id)
    if not artifact:
        return jsonify({"error": "Berkas laporan belum tersedia."}), 404

    return send_file(
        artifact["path"],
        as_attachment=True,
        download_name=artifact["filename"],
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/refresh-knowledge", methods=["POST"])
def refresh_knowledge():
    job_stats = job_manager.stats()
    if job_stats["jobs"]["queued"] or job_stats["jobs"]["running"]:
        return (
            jsonify(
                {
                    "status": "busy",
                    "error": "Sinkronisasi data sementara dikunci karena masih ada laporan yang sedang diproses.",
                }
            ),
            409,
        )

    success = kb.refresh_data()
    return jsonify({"status": "success" if success else "error"})


if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    debug = os.getenv("FLASK_DEBUG", "0").strip().lower() in {"1", "true", "yes"}
    app.run(host=host, port=port, debug=debug, threaded=True)

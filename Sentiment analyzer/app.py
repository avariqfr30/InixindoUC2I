import io
import logging
import os
from datetime import timedelta
from functools import wraps

from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from flask_cors import CORS

from auth_service import (
    ActiveSessionCapacityError,
    create_user,
    current_user,
    get_user_by_username,
    init_auth_db,
    is_user_approved,
    logout_current_session,
    session_capacity_snapshot,
    start_authenticated_session,
    user_count,
    verify_password,
)
from config import (
    ALLOW_SIGNUP,
    APP_MODE,
    APP_SECRET_KEY,
    DB_URI,
    DEFAULT_SCORE_ENGINE,
    SCORE_ENGINE_OPTIONS,
    SENTIMENT_OPTIONS,
    SESSION_COOKIE_SECURE,
    SESSION_IDLE_TIMEOUT_SECONDS,
    SIGNUP_ALLOWED_EMAIL_DOMAIN,
    SIGNUP_REQUIRES_APPROVAL,
    SMART_SUGGESTIONS,
)
from data_pipeline import KnowledgeBase
from internal_api_settings import (
    build_connector_payload,
    connector_settings_state,
    load_connector_payload,
    write_connector_payload,
)
from report_engine import ReportGenerator
from runtime import QueueCapacityError, ReportJobManager

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)
app.secret_key = APP_SECRET_KEY
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = SESSION_COOKIE_SECURE
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(seconds=SESSION_IDLE_TIMEOUT_SECONDS)

kb = KnowledgeBase(DB_URI)
generator = ReportGenerator(kb)
job_manager = ReportJobManager(generator)
logger.info("Application started in %s mode.", APP_MODE)


def _internal_api_settings_state():
    state = connector_settings_state()
    state["project_data_source"] = (
        "api" if getattr(kb.provider, "source_name", "") == "company_api" else "local"
    )
    return state


def _request_payload(data):
    return {
        "timeframe": data.get("timeframe"),
        "notes": data.get("notes", ""),
        "sentiment": data.get("sentiment", "all"),
        "segment": data.get("segment", "all"),
        "score_engine": data.get("score_engine", DEFAULT_SCORE_ENGINE),
    }


def _allowed_signup_domain():
    domain = str(SIGNUP_ALLOWED_EMAIL_DOMAIN or "").strip().lower()
    if domain and not domain.startswith("@"):
        domain = f"@{domain}"
    return domain


def _is_allowed_signup_email(value):
    email = str(value or "").strip().lower()
    allowed_domain = _allowed_signup_domain()
    return bool(email and allowed_domain and email.endswith(allowed_domain))


def login_required(view_func):
    @wraps(view_func)
    def wrapped_view(*args, **kwargs):
        if current_user():
            return view_func(*args, **kwargs)

        wants_json = request.path.startswith("/jobs/") or request.path in {
            "/api/internal-api/settings",
            "/get-config",
            "/generate",
            "/generate-job",
            "/refresh-knowledge",
        } or request.is_json
        if wants_json:
            return jsonify({"error": "Silakan login terlebih dahulu."}), 401
        return redirect(url_for("login"))

    return wrapped_view


init_auth_db()

@app.after_request
def apply_security_headers(response):
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("Referrer-Policy", "same-origin")

    if request.endpoint not in {"health", "ready", "static"}:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"

    return response


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("home"))

    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        user = get_user_by_username(username)

        if not user:
            error = "Username atau password tidak valid."
        elif not verify_password(user["password_hash"], password):
            if str(user["password_hash"]).startswith("scrypt:"):
                error = (
                    "Akun ini memakai format kata sandi lama yang tidak didukung di server ini. "
                    "Silakan buat ulang akun atau reset akun lama."
                )
            else:
                error = "Username atau password tidak valid."
        elif SIGNUP_REQUIRES_APPROVAL and not is_user_approved(user):
            error = "Akun ini masih menunggu konfirmasi admin sebelum bisa masuk."
        else:
            try:
                revoked_count = start_authenticated_session(user)
                if revoked_count:
                    logger.info(
                        "User '%s' login replaced %s previous active session(s).",
                        user["username"],
                        revoked_count,
                    )
                return redirect(url_for("home"))
            except ActiveSessionCapacityError as exc:
                error = str(exc)

    return render_template(
        "auth.html",
        mode="login",
        error=error,
        allow_signup=ALLOW_SIGNUP,
        signup_requires_approval=SIGNUP_REQUIRES_APPROVAL,
        user_count=user_count(),
    )


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if not ALLOW_SIGNUP:
        return redirect(url_for("login"))

    if current_user():
        return redirect(url_for("home"))

    error = None
    warning = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        confirm_password = request.form.get("confirm_password") or ""

        if not _is_allowed_signup_email(username):
            warning = f"Gunakan email internal dengan domain {_allowed_signup_domain()} untuk mendaftar."
        elif len(username) < 4:
            error = "Email minimal 4 karakter."
        elif len(password) < 8:
            error = "Kata sandi minimal 8 karakter."
        elif password != confirm_password:
            error = "Konfirmasi password tidak cocok."
        elif get_user_by_username(username):
            error = "Username sudah dipakai."
        else:
            created_user = create_user(
                username,
                password,
                approved=not SIGNUP_REQUIRES_APPROVAL,
                approved_by="signup_without_approval",
            )
            if SIGNUP_REQUIRES_APPROVAL:
                return render_template(
                    "auth.html",
                    mode="login",
                    error=None,
                    notice="Pendaftaran diterima. Akun harus dikonfirmasi admin sebelum bisa masuk.",
                    allow_signup=ALLOW_SIGNUP,
                    signup_requires_approval=SIGNUP_REQUIRES_APPROVAL,
                    user_count=user_count(),
                )
            try:
                start_authenticated_session(created_user)
                return redirect(url_for("home"))
            except ActiveSessionCapacityError as exc:
                error = str(exc)

    return render_template(
        "auth.html",
        mode="signup",
        error=error,
        warning=warning,
        allow_signup=ALLOW_SIGNUP,
        signup_requires_approval=SIGNUP_REQUIRES_APPROVAL,
        user_count=user_count(),
    )


@app.route("/logout", methods=["GET", "POST"])
@login_required
def logout():
    logout_current_session()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def home():
    return render_template("index.html", current_user=current_user())


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
    session_stats = session_capacity_snapshot()
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
                "auth_capacity": {
                    "active_sessions": session_stats["total_active"],
                    "max_active_total": session_stats["max_total"],
                    "max_active_per_user": session_stats["max_per_user"],
                    "session_idle_timeout_seconds": session_stats["idle_timeout_seconds"],
                },
            }
        ),
        status_code,
    )


@app.route("/get-config")
@login_required
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
    active_user = current_user()

    return jsonify(
        {
            "timeframes": timeframes,
            "sentiments": SENTIMENT_OPTIONS,
            "segments": [{"id": "all", "label": "Semua Segmen"}]
            + [{"id": segment, "label": segment} for segment in segments],
            "score_engines": SCORE_ENGINE_OPTIONS,
            "default_score_engine": DEFAULT_SCORE_ENGINE,
            "suggestions": SMART_SUGGESTIONS,
            "current_user": active_user,
        }
    )


@app.route("/generate", methods=["POST"])
@login_required
def generate_report():
    data = request.get_json(silent=True) or {}
    payload = _request_payload(data)
    timeframe = payload["timeframe"]

    if not timeframe:
        return jsonify({"error": "Parameter 'timeframe' wajib diisi."}), 400

    logger.info(
        "Generating report for timeframe='%s', sentiment='%s', segment='%s', score_engine='%s' by user '%s'.",
        timeframe,
        payload["sentiment"],
        payload["segment"],
        payload["score_engine"],
        current_user(),
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
@login_required
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
        "Queued report job %s for timeframe='%s', sentiment='%s', segment='%s', score_engine='%s' by user '%s'.",
        job["job_id"],
        payload["timeframe"],
        payload["sentiment"],
        payload["segment"],
        payload["score_engine"],
        current_user(),
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
@login_required
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
@login_required
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
@login_required
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


@app.route("/api/internal-api/settings", methods=["GET", "POST"])
@login_required
def internal_api_settings():
    if request.method == "GET":
        return jsonify(_internal_api_settings_state())

    job_stats = job_manager.stats()
    if job_stats["jobs"]["queued"] or job_stats["jobs"]["running"]:
        return (
            jsonify(
                {
                    "error": "Pengaturan data internal belum bisa diubah karena masih ada laporan yang sedang diproses.",
                }
            ),
            409,
        )

    data = request.get_json(silent=True) or {}
    try:
        payload = build_connector_payload(data, existing_payload=load_connector_payload())
        write_connector_payload(payload)
        if payload.get("enabled", True):
            kb.activate_internal_api_provider()
        refresh_success = kb.refresh_data()
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        logger.exception("Failed to update Internal API settings.")
        return jsonify({"error": str(exc)}), 500

    return jsonify(
        {
            "status": "saved",
            "refresh_status": "success" if refresh_success else "degraded",
            **_internal_api_settings_state(),
        }
    )


if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    debug = os.getenv("FLASK_DEBUG", "0").strip().lower() in {"1", "true", "yes"}
    app.run(host=host, port=port, debug=debug, threaded=True)

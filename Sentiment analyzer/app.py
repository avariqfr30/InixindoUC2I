import io
import logging
import os
import sqlite3
from functools import wraps

from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_cors import CORS
from werkzeug.security import check_password_hash, generate_password_hash

from config import (
    ALLOW_SIGNUP,
    APP_MODE,
    APP_SECRET_KEY,
    AUTH_DB_PATH,
    DB_URI,
    DEFAULT_SCORE_ENGINE,
    SCORE_ENGINE_OPTIONS,
    SENTIMENT_OPTIONS,
    SESSION_COOKIE_SECURE,
    SMART_SUGGESTIONS,
)
from core import KnowledgeBase, ReportGenerator
from runtime import QueueCapacityError, ReportJobManager

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
PASSWORD_HASH_METHOD = "pbkdf2:sha256"

app = Flask(__name__)
CORS(app)
app.secret_key = APP_SECRET_KEY
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = SESSION_COOKIE_SECURE

kb = KnowledgeBase(DB_URI)
generator = ReportGenerator(kb)
job_manager = ReportJobManager(generator)
logger.info("Application started in %s mode.", APP_MODE)


def _auth_connection():
    connection = sqlite3.connect(AUTH_DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def _init_auth_db():
    os.makedirs(os.path.dirname(AUTH_DB_PATH), exist_ok=True)
    with _auth_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.commit()


def _request_payload(data):
    return {
        "timeframe": data.get("timeframe"),
        "notes": data.get("notes", ""),
        "sentiment": data.get("sentiment", "all"),
        "segment": data.get("segment", "all"),
        "score_engine": data.get("score_engine", DEFAULT_SCORE_ENGINE),
    }


def _get_user_by_username(username):
    clean_username = str(username or "").strip()
    if not clean_username:
        return None

    with _auth_connection() as connection:
        return connection.execute(
            "SELECT id, username, password_hash FROM users WHERE username = ?",
            (clean_username,),
        ).fetchone()


def _create_user(username, password):
    with _auth_connection() as connection:
        connection.execute(
            "INSERT INTO users (username, password_hash) VALUES (?, ?)",
            (
                username.strip(),
                generate_password_hash(password, method=PASSWORD_HASH_METHOD),
            ),
        )
        connection.commit()


def _user_count():
    with _auth_connection() as connection:
        row = connection.execute("SELECT COUNT(*) AS total FROM users").fetchone()
        return int(row["total"]) if row else 0


def _verify_password(password_hash, password):
    try:
        return check_password_hash(password_hash, password)
    except AttributeError:
        if str(password_hash or "").startswith("scrypt:"):
            logger.warning(
                "Stored password hash uses scrypt, but this Python build does not support hashlib.scrypt."
            )
            return False
        raise


def _current_user():
    username = session.get("username")
    if isinstance(username, str) and username.strip():
        user = _get_user_by_username(username)
        if user:
            return user["username"]
        session.clear()
    return None


def login_required(view_func):
    @wraps(view_func)
    def wrapped_view(*args, **kwargs):
        if _current_user():
            return view_func(*args, **kwargs)

        wants_json = request.path.startswith("/jobs/") or request.path in {
            "/get-config",
            "/generate",
            "/generate-job",
            "/refresh-knowledge",
        } or request.is_json
        if wants_json:
            return jsonify({"error": "Silakan login terlebih dahulu."}), 401
        return redirect(url_for("login"))

    return wrapped_view


_init_auth_db()


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
    if _current_user():
        return redirect(url_for("home"))

    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        user = _get_user_by_username(username)

        if not user:
            error = "Username atau password tidak valid."
        elif not _verify_password(user["password_hash"], password):
            if str(user["password_hash"]).startswith("scrypt:"):
                error = (
                    "Akun ini memakai format kata sandi lama yang tidak didukung di server ini. "
                    "Silakan buat ulang akun atau reset akun lama."
                )
            else:
                error = "Username atau password tidak valid."
        else:
            session.clear()
            session["username"] = user["username"]
            return redirect(url_for("home"))

    return render_template(
        "auth.html",
        mode="login",
        error=error,
        allow_signup=ALLOW_SIGNUP,
        user_count=_user_count(),
    )


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if not ALLOW_SIGNUP:
        return redirect(url_for("login"))

    if _current_user():
        return redirect(url_for("home"))

    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        confirm_password = request.form.get("confirm_password") or ""

        if len(username) < 4:
            error = "Username minimal 4 karakter."
        elif len(password) < 8:
            error = "Kata sandi minimal 8 karakter."
        elif password != confirm_password:
            error = "Konfirmasi password tidak cocok."
        elif _get_user_by_username(username):
            error = "Username sudah dipakai."
        else:
            _create_user(username, password)
            session.clear()
            session["username"] = username
            return redirect(url_for("home"))

    return render_template(
        "auth.html",
        mode="signup",
        error=error,
        allow_signup=ALLOW_SIGNUP,
        user_count=_user_count(),
    )


@app.route("/logout", methods=["GET", "POST"])
@login_required
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def home():
    return render_template("index.html", current_user=_current_user())


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

    return jsonify(
        {
            "timeframes": timeframes,
            "sentiments": SENTIMENT_OPTIONS,
            "segments": [{"id": "all", "label": "Semua Segmen"}]
            + [{"id": segment, "label": segment} for segment in segments],
            "score_engines": SCORE_ENGINE_OPTIONS,
            "default_score_engine": DEFAULT_SCORE_ENGINE,
            "suggestions": SMART_SUGGESTIONS,
            "current_user": _current_user(),
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
        _current_user(),
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
        _current_user(),
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


if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    debug = os.getenv("FLASK_DEBUG", "0").strip().lower() in {"1", "true", "yes"}
    app.run(host=host, port=port, debug=debug, threaded=True)

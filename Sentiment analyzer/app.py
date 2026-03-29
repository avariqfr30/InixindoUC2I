import io
import logging
import os
from flask import Flask, jsonify, render_template, request, send_file
from flask_cors import CORS

from config import APP_MODE, DB_URI, SMART_SUGGESTIONS
from core import KnowledgeBase, ReportGenerator

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

kb = KnowledgeBase(DB_URI)
generator = ReportGenerator(kb)
logger.info("Application started in %s mode.", APP_MODE)

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

    return jsonify(
        {
            "timeframes": timeframes,
            "suggestions": SMART_SUGGESTIONS,
        }
    )


@app.route("/generate", methods=["POST"])
def generate_report():
    data = request.get_json(silent=True) or {}
    timeframe = data.get("timeframe")
    notes = data.get("notes", "")

    if not timeframe:
        return jsonify({"error": "Parameter 'timeframe' wajib diisi."}), 400

    logger.info("Generating report for timeframe '%s'.", timeframe)
    doc, filename = generator.run(timeframe, notes)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    return send_file(
        out,
        as_attachment=True,
        download_name=f"{filename}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/refresh-knowledge", methods=["POST"])
def refresh_knowledge():
    success = kb.refresh_data()
    return jsonify({"status": "success" if success else "error"})


if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    debug = os.getenv("FLASK_DEBUG", "0").strip().lower() in {"1", "true", "yes"}
    app.run(host=host, port=port, debug=debug, threaded=True)

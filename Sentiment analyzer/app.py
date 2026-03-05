"""
app.py
------
Flask server entry point. Connects the web interface to the reporting pipeline.
"""

import io
import logging
from flask import Flask, send_file, request, jsonify, render_template
from flask_cors import CORS

from core import SurveyData, HRReportBuilder

# Configure basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

server = Flask(__name__)
CORS(server)

# Boot up the core systems
survey_db = SurveyData()
report_builder = HRReportBuilder(survey_db)

@server.route('/')
def render_ui(): 
    """Serves the main HTML dashboard."""
    return render_template('index.html')

@server.route('/api/build-report', methods=['POST'])
def handle_report_request():
    """Receives the requested timeframe and returns the generated Word doc."""
    payload = request.json
    duration_code = payload.get('duration')
    duration_label = payload.get('duration_label')
    
    if not duration_code:
        return jsonify({"error": "Rentang waktu analisis tidak valid."}), 400
        
    doc_object, file_name = report_builder.compile_report(duration_code, duration_label)
    
    memory_file = io.BytesIO()
    doc_object.save(memory_file)
    memory_file.seek(0)
    
    return send_file(
        memory_file, 
        as_attachment=True, 
        download_name=f"{file_name}.docx", 
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == '__main__':
    logger.info("Initiating HR Sentiment Analyzer backend...")
    server.run(port=5000, debug=True, threaded=True)
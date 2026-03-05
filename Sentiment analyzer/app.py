"""
app.py
------
Main Application Entry Point.
Initializes Flask, binds routes to services, and runs the server.
"""

import io
import logging
from flask import Flask, send_file, request, jsonify, render_template
from flask_cors import CORS

# Import the core classes directly from core.py
from core import FeedbackDatabase, ReportGenerator

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(module)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Initialize core services
db = FeedbackDatabase()
generator = ReportGenerator(db)

@app.route('/')
def home(): 
    """Serves the main web interface from templates/index.html."""
    return render_template('index.html')

@app.route('/get-clients')
def get_clients():
    """API endpoint: Returns a list of clients available in the database."""
    if db.df is None: 
        return jsonify({"error": "Database belum dimuat"}), 500
    
    clients = ["ALL_OVERALL"] + sorted(db.df['Client/Partner'].unique().tolist())
    return jsonify({"clients": clients})

@app.route('/generate-report', methods=['POST'])
def generate_doc():
    """API endpoint: Accepts a client name and returns the generated Word document."""
    data = request.json
    client_name = data.get('client')
    
    if not client_name:
        return jsonify({"error": "Nama klien tidak diberikan"}), 400
        
    doc, file_name = generator.generate(client_name)
    
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    
    return send_file(
        out, 
        as_attachment=True, 
        download_name=f"{file_name}.docx", 
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == '__main__':
    logger.info("Starting AI Sentiment Analyzer Server...")
    app.run(port=5000, debug=True, threaded=True)
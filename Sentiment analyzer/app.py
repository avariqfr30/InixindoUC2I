# app.py
import io
import logging
from flask import Flask, send_file, request, jsonify, render_template
from flask_cors import CORS

from core import ReportGenerator, KnowledgeBase
from config import DB_URI, DATA_MAPPING

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

kb = KnowledgeBase(DB_URI)
generator = ReportGenerator(kb)

@app.route('/')
def home(): 
    return render_template('index.html')

@app.route('/get-config')
def get_config():
    if kb.df is None: 
        return jsonify({"error": "DB Load Failed"}), 500
        
    tree = {}
    for entity in kb.df['entity'].dropna().unique():
        tree[entity] = {"label": entity, "topics": {}}
        entity_df = kb.df[kb.df['entity'] == entity]
        for topic in entity_df['topic'].dropna().unique():
            budgets = entity_df[entity_df['topic'] == topic]['budget'].dropna().unique().tolist()
            tree[entity]["topics"][topic] = {"label": topic, "timeframes": budgets}
            
    return jsonify({
        "structure": tree, 
        "labels": [DATA_MAPPING["entity"], DATA_MAPPING["topic"]]
    })

@app.route('/generate', methods=['POST'])
def generate_doc():
    data = request.json
    
    stakeholder = data.get('stakeholder')
    service = data.get('service')
    timeframe = data.get('timeframe')
    goal = data.get('goal', 'Evaluasi Layanan Umum')
    notes = data.get('notes', '')
    
    doc, filename = generator.run(stakeholder, service, timeframe, goal, notes)
    
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    
    return send_file(
        out, 
        as_attachment=True, 
        download_name=f"{filename}.docx", 
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route('/refresh-knowledge', methods=['POST'])
def refresh():
    success = kb.refresh_data()
    return jsonify({"status": "success" if success else "error"})

if __name__ == '__main__':
    app.run(port=5000, debug=True, threaded=True)
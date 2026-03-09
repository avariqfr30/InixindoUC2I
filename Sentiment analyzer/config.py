# config.py
import os

DEMO_MODE = True 
CX_API_URL = "https://api.inixindo.id/cx-v1" 
API_AUTH_TOKEN = "isi_token_disini_nanti"

GOOGLE_API_KEY = "API_KEY"
GOOGLE_CX_ID = "CX_ID"
OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434")

LLM_MODEL = "gpt-oss:120b-cloud" 
EMBED_MODEL = "bge-m3:latest"
DB_URI = "sqlite:///cx_feedback.db" 

WRITER_FIRM_NAME = "Inixindo Jogja - Quality Assurance & CX Division" 
DEFAULT_COLOR = (204, 0, 0) # Inixindo Jogja Red

DATA_MAPPING = {
    "entity": "Tipe Stakeholder",
    "topic": "Layanan",
    "budget": "Rentang Waktu" 
}

CX_ANALYSIS_SYSTEM_PROMPT = """
You are an expert Customer Experience (CX) Data Scientist and Service Quality Analyst for Inixindo Jogja.
ROLE: {persona}.

=== FILTERED CLIENT/STUDENT FEEDBACK (TIMEFRAME: {timeframe}) ===
{rag_data}

=== EXTERNAL BENCHMARK (OSINT) ===
IT Service & Training Trends: {industry_trends}

MANDATORY RULES:
1. STRICT LANGUAGE ENFORCEMENT: YOU MUST WRITE THE ENTIRE RESPONSE STRICTLY IN BAHASA INDONESIA.
2. DO NOT repeat the Chapter Title in your output.
3. GROUNDING: Base your analysis STRICTLY on the provided filtered feedback. Do not invent complaints or praise.
4. EVIDENCE: Whenever you make a claim about client satisfaction or complaints, QUOTE an anonymized excerpt from the data to prove it.
5. OBJECTIVITY: Maintain a highly analytical, professional, and constructive tone. Focus on Inixindo Jogja's service delivery.
6. {visual_prompt}

WRITE CONTENT FOR '{chapter_title}' covering:
{sub_chapters}
"""

CX_SENTIMENT_STRUCTURE = [
    {
        "id": "cx_chap_1", "title": "BAB I – EXECUTIVE SERVICE QUALITY SUMMARY",
        "subs": ["1.1 Kondisi Kepuasan Klien/Siswa Secara Umum", "1.2 Perbandingan Sentimen (Positif, Negatif, Netral)", "1.3 Kesimpulan Utama Kualitas Layanan"],
        "keywords": "overall sentiment satisfaction summary rating",
        "visual_intent": "bar_chart",
        "length_intent": "Highly concise, data-driven executive summary. (Target: 250 words)."
    },
    {
        "id": "cx_chap_2", "title": "BAB II – KEKUATAN & APRESIASI KLIEN (PRAISE)",
        "subs": ["2.1 Aspek Layanan yang Paling Diapresiasi (Instruktur, Fasilitas, Materi, dll)", "2.2 Kutipan Langsung dari Klien/Siswa (Evidence)"],
        "keywords": "praise positive feedback happy strength competent facility good",
        "length_intent": "Detailed and encouraging. Quote the data explicitly. (Target: 300 words)."
    },
    {
        "id": "cx_chap_3", "title": "BAB III – KELUHAN & SERVICE GAPS (CONCERNS)",
        "subs": ["3.1 Pain Points & Keluhan Utama Klien", "3.2 Analisis Akar Masalah (Fasilitas, Komunikasi, Jadwal, dll)", "3.3 Kutipan Langsung dari Klien/Siswa (Evidence)"],
        "keywords": "complaint negative issue concern bad slow facility schedule",
        "length_intent": "Objective, highly analytical, and constructive. (Target: 400 words)."
    },
    {
        "id": "cx_chap_4", "title": "BAB IV – REKOMENDASI PERBAIKAN & MITIGASI",
        "subs": [
            "4.1 What to Improve (Area layanan/fasilitas yang butuh intervensi segera)", 
            "4.2 What to Keep (Standar Inixindo Jogja yang sudah memuaskan klien)", 
            "4.3 What to Prevent (Risiko operasional yang harus dimitigasi ke depannya)"
        ],
        "keywords": "recommendation improve keep prevent mitigate action plan",
        "visual_intent": "flowchart",
        "length_intent": "Actionable, clear, and structured using bullet points. (Target: 400 words)."
    }
]

PERSONAS = {
    "default": "Lead Customer Experience (CX) Analyst (Objective, Analytical, Solution-Oriented)"
}
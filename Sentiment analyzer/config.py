# config.py
import os

DEMO_MODE = True 

SERPER_API_KEY = "YOUR_SERPER_API_KEY"

OLLAMA_HOST = "http://127.0.0.1:11434"

LLM_MODEL = "gpt-oss:120b-cloud" 
EMBED_MODEL = "bge-m3:latest"
DB_URI = "sqlite:///data/cx_feedback.db" 

WRITER_FIRM_NAME = "Inixindo Jogja - Quality Assurance & CX Division" 
DEFAULT_COLOR = (204, 0, 0) # Inixindo Jogja Red

# --- SMART SUGGESTIONS (Holistic / Macro Level) ---
SMART_SUGGESTIONS = [
    "Tolong bandingkan ekspektasi klien BUMN/Corporate dengan peserta Mahasiswa/Personal.",
    "Fokuskan analisis pada keluhan terkait infrastruktur fisik dan jaringan secara keseluruhan.",
    "Berikan analisis mendalam mengenai performa instruktur dan konsultan di semua layanan.",
    "Bandingkan sentimen klien Pemerintahan (Gov) terhadap layanan pelatihan vs layanan audit/konsultasi."
]

CX_ANALYSIS_SYSTEM_PROMPT = """
You are the Chief Customer Experience (CX) Officer for Inixindo Jogja.
ROLE: {persona}.

=== HOLISTIC COMPANY FEEDBACK (TIMEFRAME: {timeframe}) ===
{rag_data}

=== EXTERNAL OSINT BENCHMARKS (MACRO TRENDS) ===
Demographic & IT Education Trends in Indonesia: {industry_trends}

MANDATORY RULES:
1. STRICT LANGUAGE: Write the entire response strictly in professional Bahasa Indonesia.
2. HOLISTIC SYNTHESIS: You are analyzing the ENTIRE company's performance. You MUST compare different demographics (e.g., Gov vs Corporate vs Students) and different services (e.g., Training vs Consulting). 
3. STRICT SUB-CHAPTER ENFORCEMENT: You MUST use Markdown Headers (###) for EVERY single sub-chapter listed below. You are FORBIDDEN from leaving any sub-chapter empty. Write at least 150 words per sub-chapter.
4. EVIDENCE: Quote anonymized excerpts from the internal feedback data to prove your points.
5. NO TITLE REPETITION: Do NOT write '{chapter_title}' at the start of your response.
6. {visual_prompt}

WRITE DETAILED CONTENT FOR THE FOLLOWING SUB-CHAPTERS:
{sub_chapters}
"""

CX_SENTIMENT_STRUCTURE = [
    {
        "id": "cx_chap_1", "title": "BAB I – MACRO CX & SERVICE QUALITY SUMMARY",
        "subs": [
            "1.1 Kondisi Kepuasan Pelanggan Inixindo Jogja Secara Keseluruhan", 
            "1.2 Perbandingan Sentimen Antar Demografi (Pemerintah vs BUMN vs Mahasiswa/Personal)", 
            "1.3 Layanan dengan Performa Terbaik & Terburuk Periode Ini"
        ],
        "keywords": "overall sentiment satisfaction summary rating compare demographic service",
        "visual_intent": "bar_chart",
        "length_intent": "Highly concise, data-driven executive summary."
    },
    {
        "id": "cx_chap_2", "title": "BAB II – APRESIASI & KEKUATAN LINTAS LAYANAN",
        "subs": [
            "2.1 Puncak Apresiasi Klien (Instruktur, Materi, atau Layanan Administratif)", 
            "2.2 Testimoni Positif Berdasarkan Segmen Klien",
            "2.3 Kutipan Langsung dari Klien/Siswa (Evidence)"
        ],
        "keywords": "praise positive feedback happy strength competent facility good across services",
        "length_intent": "Detailed and encouraging. Quote the data explicitly."
    },
    {
        "id": "cx_chap_3", "title": "BAB III – ANALISIS KESENJANGAN & KELUHAN SISTEMIK",
        "subs": [
            "3.1 Pain Points Utama Lintas Layanan (Fasilitas, Komunikasi, Timeline)", 
            "3.2 Analisis Kesenjangan Ekspektasi Berdasarkan Demografi (Misal: Ekspektasi Senior vs Pemula)", 
            "3.3 Kutipan Langsung dari Klien/Siswa (Evidence)"
        ],
        "keywords": "complaint negative issue concern bad slow facility schedule expectation gap",
        "length_intent": "Objective, highly analytical, and constructive."
    },
    {
        "id": "cx_chap_4", "title": "BAB IV – REKOMENDASI STRATEGIS & MITIGASI RISIKO",
        "subs": [
            "4.1 Intervensi Segera (Area kritis yang berdampak pada banyak demografi)", 
            "4.2 Penyesuaian Strategi Layanan Berdasarkan Tren OSINT & Perilaku Konsumen Terkini", 
            "4.3 Langkah Preventif Operasional Ke Depan"
        ],
        "keywords": "recommendation improve keep prevent mitigate action plan demographic behavior trend future",
        "visual_intent": "flowchart",
        "length_intent": "Actionable, clear, and structured using bullet points."
    }
]

PERSONAS = {
    "default": "Chief CX Officer (Visionary, Analytical, Highly Objective, Strategic)"
}
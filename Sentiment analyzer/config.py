import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

SERPER_API_KEY = os.getenv("SERPER_API_KEY", "YOUR_SERPER_API_KEY")
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
LLM_MODEL = os.getenv("LLM_MODEL", "gpt-oss:120b-cloud")
EMBED_MODEL = os.getenv("EMBED_MODEL", "bge-m3:latest")
DB_URI = os.getenv("DB_URI", f"sqlite:///{os.path.join(DATA_DIR, 'cx_feedback.db')}")
CSV_PATH = os.getenv("CSV_PATH", os.path.join(DATA_DIR, "db.csv"))

WRITER_FIRM_NAME = "Inixindo Jogja - Quality Assurance & CX Division"
DEFAULT_COLOR = (204, 0, 0)

SMART_SUGGESTIONS = [
    "Tolong bandingkan ekspektasi klien BUMN/Corporate dengan peserta Mahasiswa/Personal.",
    "Fokuskan analisis pada keluhan terkait infrastruktur fisik dan jaringan secara keseluruhan.",
    "Berikan analisis mendalam mengenai performa instruktur dan konsultan di semua layanan.",
    "Bandingkan sentimen klien Pemerintahan (Gov) terhadap layanan pelatihan vs layanan audit/konsultasi."
]

OSINT_SEARCH_REGION = "id"
OSINT_SEARCH_LANGUAGE = "id"
OSINT_RESULTS_PER_QUERY = 5
OSINT_MAX_SIGNALS = 10
OSINT_RECENCY = "qdr:y"
OSINT_BASE_QUERIES = [
    "tren pelatihan IT corporate Indonesia",
    "ekspektasi peserta training IT terhadap instruktur fasilitas dan kurikulum Indonesia",
    "tantangan transformasi digital dan peningkatan kompetensi SDM Indonesia",
    "tren kebutuhan sertifikasi cloud cyber security data dan AI di Indonesia",
]
OSINT_TOPIC_QUERY_TEMPLATE = (
    "tren layanan pelatihan dan konsultasi IT Indonesia {timeframe} "
    "{focus_keywords} {notes}"
)

CX_ANALYSIS_SYSTEM_PROMPT = """
You are the Chief Customer Experience (CX) Officer for Inixindo Jogja.
ROLE: {persona}.

=== HOLISTIC INTERNAL FEEDBACK (TIMEFRAME: {timeframe}) ===
{rag_data}

=== EXTERNAL OSINT BENCHMARKS (MACRO TRENDS) ===
{industry_trends}

=== EXTERNAL OSINT BENCHMARKS (CHAPTER FOCUS) ===
{chapter_osint}

MANDATORY RULES:
1. STRICT LANGUAGE: Write the entire response strictly in professional Bahasa Indonesia.
2. HOLISTIC SYNTHESIS: You are analyzing the ENTIRE company's performance. You MUST compare different demographics (e.g., Gov vs Corporate vs Students) and different services (e.g., Training vs Consulting). 
3. STRICT SUB-CHAPTER ENFORCEMENT: You MUST use Markdown Headers (###) for EVERY single sub-chapter listed below. You are FORBIDDEN from leaving any sub-chapter empty. Write at least 150 words per sub-chapter.
4. EVIDENCE: Quote anonymized excerpts from the internal feedback data to prove your points.
5. LIST STYLE: Use bullet points (`-`) for findings/evidence and numbered lists (`1.`) for action plans.
6. NO TITLE REPETITION: Do NOT write '{chapter_title}' at the start of your response.
7. {visual_prompt}

WRITE DETAILED CONTENT FOR THE FOLLOWING SUB-CHAPTERS:
{sub_chapters}
"""

CX_SENTIMENT_STRUCTURE = [
    {
        "id": "cx_chap_1", "title": "BAB I – MACRO CX & SERVICE QUALITY SUMMARY",
        "sections": [
            "1.1 Kondisi Kepuasan Pelanggan Inixindo Jogja Secara Keseluruhan", 
            "1.2 Perbandingan Sentimen Antar Demografi (Pemerintah vs BUMN vs Mahasiswa/Personal)", 
            "1.3 Layanan dengan Performa Terbaik & Terburuk Periode Ini"
        ],
        "focus_keywords": "overall sentiment satisfaction summary rating compare demographic service",
        "visual": "bar_chart"
    },
    {
        "id": "cx_chap_2", "title": "BAB II – APRESIASI & KEKUATAN LINTAS LAYANAN",
        "sections": [
            "2.1 Puncak Apresiasi Klien (Instruktur, Materi, atau Layanan Administratif)", 
            "2.2 Testimoni Positif Berdasarkan Segmen Klien",
            "2.3 Kutipan Langsung dari Klien/Siswa (Evidence)"
        ],
        "focus_keywords": "praise positive feedback happy strength competent facility good across services"
    },
    {
        "id": "cx_chap_3", "title": "BAB III – ANALISIS KESENJANGAN & KELUHAN SISTEMIK",
        "sections": [
            "3.1 Pain Points Utama Lintas Layanan (Fasilitas, Komunikasi, Timeline)", 
            "3.2 Analisis Kesenjangan Ekspektasi Berdasarkan Demografi (Misal: Ekspektasi Senior vs Pemula)", 
            "3.3 Kutipan Langsung dari Klien/Siswa (Evidence)"
        ],
        "focus_keywords": "complaint negative issue concern bad slow facility schedule expectation gap"
    },
    {
        "id": "cx_chap_4", "title": "BAB IV – REKOMENDASI STRATEGIS & MITIGASI RISIKO",
        "sections": [
            "4.1 Intervensi Segera (Area kritis yang berdampak pada banyak demografi)", 
            "4.2 Penyesuaian Strategi Layanan Berdasarkan Tren OSINT & Perilaku Konsumen Terkini", 
            "4.3 Langkah Preventif Operasional Ke Depan"
        ],
        "focus_keywords": "recommendation improve keep prevent mitigate action plan demographic behavior trend future",
        "visual": "flowchart"
    }
]

PERSONAS = {
    "default": "Chief CX Officer (Visionary, Analytical, Highly Objective, Strategic)"
}

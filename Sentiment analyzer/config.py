import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

SUPPORTED_APP_MODES = {"demo", "hybrid"}
APP_MODE = os.getenv("APP_MODE", "demo").strip().lower()
if APP_MODE not in SUPPORTED_APP_MODES:
    APP_MODE = "demo"

DEMO_MODE = APP_MODE == "demo"
INTERNAL_DATA_MODE = "csv" if DEMO_MODE else "api"
EXTERNAL_DATA_MODE = "osint"

SERPER_API_KEY = os.getenv("SERPER_API_KEY", "YOUR_SERPER_API_KEY")
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
EMBED_MODEL = os.getenv("EMBED_MODEL", "bge-m3:latest")
DB_URI = os.getenv("DB_URI", f"sqlite:///{os.path.join(DATA_DIR, 'cx_feedback.db')}")
CSV_PATH = os.getenv("CSV_PATH", os.path.join(DATA_DIR, "db.csv"))

INTERNAL_API_BASE_URL = os.getenv("INTERNAL_API_BASE_URL", "").rstrip("/")
INTERNAL_API_KEY = os.getenv("INTERNAL_API_KEY", "")
INTERNAL_API_FEEDBACK_ENDPOINT = os.getenv(
    "INTERNAL_API_FEEDBACK_ENDPOINT",
    "/feedback",
)
INTERNAL_API_TIMEOUT_SECONDS = int(os.getenv("INTERNAL_API_TIMEOUT_SECONDS", "20"))
ENABLE_VECTOR_INDEX = os.getenv("ENABLE_VECTOR_INDEX", "0").strip().lower() in {
    "1",
    "true",
    "yes",
}

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

DATA_ACQUISITION_POLICY = {
    "demo": {
        "label": "Demo Mode",
        "internal_source": "Demo CSV dataset",
        "external_source": "OSINT",
        "internal_scope": [
            "Sample feedback records",
            "Sample stakeholder segments",
            "Sample service history",
        ],
        "external_scope": [
            "Market trends",
            "Public benchmarks",
            "Public sentiment",
        ],
    },
    "hybrid": {
        "label": "Hybrid Mode",
        "internal_source": "Company internal API",
        "external_source": "OSINT",
        "internal_scope": [
            "Customer feedback",
            "Operational service records",
            "Customer segmentation",
            "Performance and service outcomes",
        ],
        "external_scope": [
            "Market trends",
            "Competitor benchmarks",
            "Public reviews and media signals",
        ],
    },
}

CX_SENTIMENT_STRUCTURE = [
    {
        "id": "cx_chap_1", "title": "BAB I – DESCRIPTIVE ANALYTICS & FEEDBACK GOVERNANCE",
        "sections": [
            "1.1 Ringkasan Cakupan Feedback dan Tata Kelola",
            "1.2 Distribusi Sentimen, Rating, dan Volume",
            "1.3 Distribusi Stakeholder, Layanan, dan Kanal/Sumber"
        ],
        "focus_keywords": "feedback governance descriptive analytics rating stakeholder service channel source",
        "visual": "bar_chart"
    },
    {
        "id": "cx_chap_2", "title": "BAB II – DIAGNOSTIC ANALYTICS",
        "sections": [
            "2.1 Akar Masalah Utama dan Pain Point Dominan",
            "2.2 Kekuatan yang Konsisten dan Area yang Perlu Dijaga",
            "2.3 Bukti Verbatim, Kesenjangan Proses, dan Segmentasi Masalah"
        ],
        "focus_keywords": "diagnostic analytics root cause complaint praise service quality process gap"
    },
    {
        "id": "cx_chap_3", "title": "BAB III – PREDICTIVE ANALYTICS",
        "sections": [
            "3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut",
            "3.2 Prediksi Segmen dan Layanan yang Paling Rentan",
            "3.3 Tren Eksternal yang Berpotensi Memperbesar Risiko"
        ],
        "focus_keywords": "predictive analytics risk trend forecast segment service vulnerability"
    },
    {
        "id": "cx_chap_4", "title": "BAB IV – PRESCRIPTIVE ANALYTICS",
        "sections": [
            "4.1 Intervensi Prioritas 30 Hari",
            "4.2 Penguatan Tata Kelola Feedback dan Eskalasi",
            "4.3 Rencana Tindak Lanjut Lintas Fungsi"
        ],
        "focus_keywords": "prescriptive analytics recommendation action plan governance mitigation",
        "visual": "flowchart"
    }
]

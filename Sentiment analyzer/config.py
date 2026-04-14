import json
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
AUTH_DB_PATH = os.getenv("AUTH_DB_PATH", os.path.join(DATA_DIR, "auth.db"))
APP_SECRET_KEY = os.getenv("APP_SECRET_KEY", "change-this-secret-before-deployment")
ALLOW_SIGNUP = os.getenv("ALLOW_SIGNUP", "1").strip().lower() in {"1", "true", "yes"}
SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", "0").strip().lower() in {
    "1",
    "true",
    "yes",
}
REPORT_ARTIFACT_DIR = os.getenv(
    "REPORT_ARTIFACT_DIR",
    os.path.join(DATA_DIR, "generated_reports"),
)
JOB_STATE_PATH = os.getenv(
    "JOB_STATE_PATH",
    os.path.join(DATA_DIR, "report_jobs.json"),
)
REPORT_JOB_WORKERS = int(os.getenv("REPORT_JOB_WORKERS", "3"))
REPORT_MAX_PENDING_JOBS = int(os.getenv("REPORT_MAX_PENDING_JOBS", "24"))
REPORT_JOB_RETENTION_SECONDS = int(
    os.getenv("REPORT_JOB_RETENTION_SECONDS", "86400")
)

INTERNAL_API_BASE_URL = os.getenv("INTERNAL_API_BASE_URL", "").rstrip("/")
INTERNAL_API_KEY = os.getenv("INTERNAL_API_KEY", "")
INTERNAL_API_FEEDBACK_ENDPOINT = os.getenv(
    "INTERNAL_API_FEEDBACK_ENDPOINT",
    "/feedback",
)
INTERNAL_API_TIMEOUT_SECONDS = int(os.getenv("INTERNAL_API_TIMEOUT_SECONDS", "20"))
INTERNAL_API_AUTH_MODE = os.getenv("INTERNAL_API_AUTH_MODE", "api_key").strip().lower()
INTERNAL_API_AUTH_HEADER = os.getenv("INTERNAL_API_AUTH_HEADER", "Authorization").strip() or "Authorization"
INTERNAL_API_AUTH_PREFIX = os.getenv("INTERNAL_API_AUTH_PREFIX", "Bearer").strip()
INTERNAL_API_USERNAME = os.getenv("INTERNAL_API_USERNAME", "").strip()
INTERNAL_API_PASSWORD = os.getenv("INTERNAL_API_PASSWORD", "")
INTERNAL_API_SOURCE_URL = os.getenv("INTERNAL_API_SOURCE_URL", "").strip()
INTERNAL_API_SOURCE_METHOD = os.getenv("INTERNAL_API_SOURCE_METHOD", "GET").strip().upper() or "GET"
INTERNAL_API_SOURCE_BODY_MODE = os.getenv("INTERNAL_API_SOURCE_BODY_MODE", "json").strip().lower() or "json"
ENABLE_VECTOR_INDEX = os.getenv("ENABLE_VECTOR_INDEX", "0").strip().lower() in {
    "1",
    "true",
    "yes",
}


def _load_json_object(env_name, fallback):
    raw_value = os.getenv(env_name, "").strip()
    if not raw_value:
        return fallback
    try:
        parsed = json.loads(raw_value)
    except json.JSONDecodeError:
        return fallback
    return parsed if isinstance(parsed, dict) else fallback


def _load_internal_api_endpoints():
    defaults = {
        "feedback": {
            "path": INTERNAL_API_SOURCE_URL or INTERNAL_API_FEEDBACK_ENDPOINT,
            "method": INTERNAL_API_SOURCE_METHOD if INTERNAL_API_SOURCE_URL else "GET",
            "body_mode": INTERNAL_API_SOURCE_BODY_MODE,
            "record_keys": ["items", "data", "results", "records", "feedback"],
            "query_params": dict(INTERNAL_API_SOURCE_PARAMS),
            "headers": dict(INTERNAL_API_SOURCE_HEADERS),
            "auto_discover": True,
        },
        "services": {
            "path": "/services",
            "method": "GET",
            "body_mode": "json",
            "record_keys": ["items", "data", "results", "records", "services"],
            "query_params": {},
            "headers": {},
            "auto_discover": True,
        },
        "stakeholders": {
            "path": "/stakeholders",
            "method": "GET",
            "body_mode": "json",
            "record_keys": ["items", "data", "results", "records", "stakeholders"],
            "query_params": {},
            "headers": {},
            "auto_discover": True,
        },
        "operations": {
            "path": "/operations",
            "method": "GET",
            "body_mode": "json",
            "record_keys": ["items", "data", "results", "records", "operations"],
            "query_params": {},
            "headers": {},
            "auto_discover": True,
        },
    }

    overrides = _load_json_object("INTERNAL_API_ENDPOINTS_JSON", {})
    merged = {}
    for endpoint_name, default_spec in defaults.items():
        override_spec = overrides.get(endpoint_name, {})
        if isinstance(override_spec, dict):
            merged[endpoint_name] = {**default_spec, **override_spec}
        else:
            merged[endpoint_name] = dict(default_spec)

    for endpoint_name, override_spec in overrides.items():
        if endpoint_name in merged or not isinstance(override_spec, dict):
            continue
        merged[endpoint_name] = {
            "path": override_spec.get("path", f"/{endpoint_name}"),
            "method": override_spec.get("method", "GET"),
            "body_mode": override_spec.get("body_mode", "json"),
            "record_keys": override_spec.get(
                "record_keys",
                ["items", "data", "results", "records", endpoint_name],
            ),
            "query_params": override_spec.get("query_params", {}),
            "headers": override_spec.get("headers", {}),
            "auto_discover": bool(override_spec.get("auto_discover", True)),
        }

    return merged


INTERNAL_API_DEFAULT_HEADERS = _load_json_object(
    "INTERNAL_API_DEFAULT_HEADERS_JSON",
    {},
)
INTERNAL_API_SOURCE_HEADERS = _load_json_object(
    "INTERNAL_API_SOURCE_HEADERS_JSON",
    {},
)
INTERNAL_API_SOURCE_PARAMS = _load_json_object(
    "INTERNAL_API_SOURCE_PARAMS_JSON",
    {},
)
INTERNAL_API_ENDPOINTS = _load_internal_api_endpoints()

WRITER_FIRM_NAME = "Inixindo Jogja - Quality Assurance & CX Division"
DEFAULT_COLOR = (204, 0, 0)

SMART_SUGGESTIONS = [
    "Soroti area mana yang paling layak dijadikan prioritas perbaikan dan peluang pilot implementasi terlebih dahulu.",
    "Fokuskan analisis pada dampak bisnis, kesiapan data, dan siapa owner tindak lanjut tiap area.",
    "Jelaskan kontrol risiko, tata kelola, dan indikator kapan inisiatif perlu dilanjutkan, diubah, atau dihentikan.",
    "Tekankan pembelajaran organisasi, perubahan cara kerja, dan kebutuhan capability lintas fungsi."
]

SENTIMENT_OPTIONS = [
    {"id": "all", "label": "Semua Sentimen"},
    {"id": "positive", "label": "Positif"},
    {"id": "neutral", "label": "Netral"},
    {"id": "negative", "label": "Negatif"},
]

DEFAULT_SCORE_ENGINE = "experience_index"
SCORE_ENGINE_OPTIONS = [
    {
        "id": "learning_score",
        "label": "Learning Score",
        "description": "Menekankan kualitas pembelajaran, instruktur, materi, dan outcome belajar.",
    },
    {
        "id": "service_score",
        "label": "Service Score",
        "description": "Menekankan responsiveness, koordinasi, SLA, dan kualitas layanan secara umum.",
    },
    {
        "id": "facility_score",
        "label": "Facility Score",
        "description": "Menekankan fasilitas, ruang, jaringan, dan kesiapan operasional pendukung.",
    },
    {
        "id": "experience_index",
        "label": "Experience Index",
        "description": "Membaca pengalaman pelanggan secara menyeluruh lintas touchpoint dan tema.",
    },
]

SCORE_ENGINE_PROFILES = {
    "learning_score": {
        "label": "Learning Score",
        "summary_label": "kualitas pembelajaran",
        "narrative_focus": "kualitas instruktur, relevansi materi, kenyamanan belajar, dan hasil yang dirasakan peserta",
        "forecast_label": "Learning Score",
        "theme_weights": {
            "instructor": 1.5,
            "material": 1.4,
            "outcome": 1.25,
            "schedule": 0.9,
            "communication": 0.8,
            "facility": 0.7,
            "responsiveness": 0.6,
        },
    },
    "service_score": {
        "label": "Service Score",
        "summary_label": "kualitas layanan",
        "narrative_focus": "responsiveness, koordinasi, ketepatan tindak lanjut, dan kualitas eksekusi layanan",
        "forecast_label": "Service Score",
        "theme_weights": {
            "responsiveness": 1.45,
            "communication": 1.3,
            "schedule": 1.15,
            "outcome": 0.9,
            "instructor": 0.8,
            "material": 0.75,
            "facility": 0.65,
        },
    },
    "facility_score": {
        "label": "Facility Score",
        "summary_label": "kesiapan fasilitas",
        "narrative_focus": "fasilitas kelas, jaringan, ruang, sarana pendukung, dan kesiapan operasional sebelum delivery",
        "forecast_label": "Facility Score",
        "theme_weights": {
            "facility": 1.65,
            "schedule": 1.1,
            "communication": 0.8,
            "responsiveness": 0.7,
            "instructor": 0.55,
            "material": 0.45,
            "outcome": 0.45,
        },
    },
    "experience_index": {
        "label": "Experience Index",
        "summary_label": "pengalaman pelanggan end-to-end",
        "narrative_focus": "keseluruhan customer journey, dari koordinasi awal hingga outcome pasca-layanan",
        "forecast_label": "Experience Index",
        "theme_weights": {
            "responsiveness": 1.1,
            "communication": 1.1,
            "schedule": 1.0,
            "facility": 1.0,
            "instructor": 1.15,
            "material": 1.1,
            "outcome": 1.2,
        },
    },
}

CUSTOMER_JOURNEY_STAGES = [
    {
        "id": "pre_engagement",
        "label": "Pra-Layanan dan Ekspektasi",
        "theme_ids": ["communication", "responsiveness"],
        "description": "Tahap awal saat pelanggan membangun ekspektasi, meminta informasi, dan menilai kejelasan respons awal.",
    },
    {
        "id": "preparation_readiness",
        "label": "Persiapan dan Kesiapan Delivery",
        "theme_ids": ["schedule", "facility", "communication"],
        "description": "Tahap penyiapan jadwal, administrasi, ruang, perangkat, dan koordinasi operasional sebelum layanan berjalan.",
    },
    {
        "id": "delivery_experience",
        "label": "Pelaksanaan Layanan",
        "theme_ids": ["instructor", "material", "facility", "schedule"],
        "description": "Tahap inti ketika pelanggan merasakan kualitas fasilitator, materi, ritme sesi, dan kenyamanan eksekusi layanan.",
    },
    {
        "id": "follow_up_outcome",
        "label": "Tindak Lanjut dan Outcome",
        "theme_ids": ["outcome", "responsiveness", "communication"],
        "description": "Tahap pasca-layanan saat pelanggan menilai manfaat, penutupan isu, dan keberlanjutan tindak lanjut.",
    },
]

OSINT_SEARCH_REGION = "id"
OSINT_SEARCH_LANGUAGE = "id"
OSINT_RESULTS_PER_QUERY = 5
OSINT_MAX_SIGNALS = 10
OSINT_RECENCY = "qdr:y"
OSINT_CACHE_PATH = os.getenv(
    "OSINT_CACHE_PATH",
    os.path.join(DATA_DIR, "osint_cache.json"),
)
OSINT_CACHE_TTL_SECONDS = int(os.getenv("OSINT_CACHE_TTL_SECONDS", "21600"))
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

ADOPTION_READINESS_PILLARS = [
    {
        "id": "business_use_case",
        "title": "5.1 Prioritas Sasaran Bisnis",
        "guiding_question": "Masalah apa yang paling layak diprioritaskan dan apa dampaknya terhadap revenue, cost, atau risk?",
    },
    {
        "id": "data_model_foundation",
        "title": "5.2 Kesiapan Data dan Fondasi Analitik",
        "guiding_question": "Apakah data tersedia, cukup bersih, dan sudah jelas siapa owner serta standar pengelolaannya?",
    },
    {
        "id": "infrastructure_architecture",
        "title": "5.3 Kesiapan Arsitektur dan Operasionalisasi",
        "guiding_question": "Arsitektur seperti apa yang cukup aman, scalable, dan realistis untuk tahap implementasi saat ini?",
    },
    {
        "id": "people_capability",
        "title": "5.4 Peran, Kapabilitas, dan Kepemilikan Tindak Lanjut",
        "guiding_question": "Siapa yang perlu dilibatkan agar inisiatif ini benar-benar dekat dengan kebutuhan bisnis dan dapat dieksekusi?",
    },
    {
        "id": "governance",
        "title": "5.5 Kontrol Risiko dan Tata Kelola",
        "guiding_question": "Kontrol apa yang dibutuhkan agar risiko, kualitas rekomendasi, dan SOP tetap terjaga?",
    },
    {
        "id": "culture",
        "title": "5.6 Perubahan Kerja dan Pembelajaran Organisasi",
        "guiding_question": "Perubahan perilaku kerja apa yang perlu dibangun agar inisiatif ini menjadi kebiasaan kerja, bukan eksperimen sesaat?",
    },
]

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
    },
    {
        "id": "cx_chap_5", "title": "BAB V – REKOMENDASI IMPLEMENTASI DAN PENGUATAN ORGANISASI",
        "sections": [
            "5.1 Prioritas Sasaran Bisnis",
            "5.2 Kesiapan Data dan Fondasi Analitik",
            "5.3 Kesiapan Arsitektur dan Operasionalisasi",
            "5.4 Peran, Kapabilitas, dan Kepemilikan Tindak Lanjut",
            "5.5 Kontrol Risiko dan Tata Kelola",
            "5.6 Perubahan Kerja dan Pembelajaran Organisasi"
        ],
        "focus_keywords": "implementation readiness business priority data architecture capability governance learning culture pilot",
    }
]

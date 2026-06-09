import json
import logging
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
PROFILES_DIR = os.path.join(BASE_DIR, "profiles")

SUPPORTED_APP_MODES = {"demo", "hybrid"}
SUPPORTED_APP_PROFILES = {"demo", "production"}
APP_PROFILE = os.getenv("APP_PROFILE", "demo").strip().lower()
if APP_PROFILE not in SUPPORTED_APP_PROFILES:
    APP_PROFILE = "demo"

_PROFILE_DEFAULT_MODE = {
    "demo": "demo",
    "production": "hybrid",
}
APP_MODE = os.getenv("APP_MODE", _PROFILE_DEFAULT_MODE[APP_PROFILE]).strip().lower()
if APP_MODE not in SUPPORTED_APP_MODES:
    APP_MODE = _PROFILE_DEFAULT_MODE[APP_PROFILE]

DEMO_MODE = APP_MODE == "demo"
INTERNAL_DATA_MODE = "csv" if DEMO_MODE else "api"
EXTERNAL_DATA_MODE = "osint"


def _int_env(name: str, default: int) -> int:
    raw = os.getenv(name, str(default))
    try:
        return int(raw.strip())
    except ValueError:
        logging.warning("Invalid integer for %s=%r, using default %d", name, raw, default)
        return default


def _float_env(name: str, default: float) -> float:
    raw = os.getenv(name, str(default))
    try:
        return float(raw.strip())
    except ValueError:
        logging.warning("Invalid float for %s=%r, using default %f", name, raw, default)
        return default


def _load_csv_list(env_name, fallback):
    raw_value = os.getenv(env_name, "").strip()
    values = raw_value.split(",") if raw_value else fallback
    return [str(value).strip().lower() for value in values if str(value).strip()]


def _load_csv_set(env_name):
    return {value for value in _load_csv_list(env_name, []) if value}


SERPER_API_KEY = os.getenv("SERPER_API_KEY", "YOUR_SERPER_API_KEY")
OLLAMA_API_KEY = os.getenv("OLLAMA_API_KEY", "")
OLLAMA_WEB_SEARCH_URL = os.getenv("OLLAMA_WEB_SEARCH_URL", "https://ollama.com/api/web_search")
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
EMBED_MODEL = os.getenv("EMBED_MODEL", "bge-m3:latest")
DB_URI = os.getenv("DB_URI", f"sqlite:///{os.path.join(DATA_DIR, 'cx_feedback.db')}")
CSV_PATH = os.getenv("CSV_PATH", os.path.join(DATA_DIR, "db.csv"))
AUTH_DB_PATH = os.getenv("AUTH_DB_PATH", os.path.join(DATA_DIR, "auth.db"))

_DEFAULT_SECRET = "change-this-secret-before-deployment"
APP_SECRET_KEY = os.getenv("APP_SECRET_KEY", _DEFAULT_SECRET)

def _validate_secret_key():
    if APP_PROFILE != "demo" and APP_SECRET_KEY == _DEFAULT_SECRET:
        sys.exit(
            "\n[FATAL] APP_SECRET_KEY is not set.\n"
            "Set APP_SECRET_KEY in your profiles/*.env file or environment.\n"
            "Generate one with: python3 -c \"import secrets; print(secrets.token_urlsafe(48))\"\n"
        )

_validate_secret_key()

_signup_default = "1" if APP_PROFILE == "demo" else "0"
ALLOW_SIGNUP = os.getenv("ALLOW_SIGNUP", _signup_default).strip().lower() in {"1", "true", "yes"}

_approval_default = "0" if APP_PROFILE == "demo" else "1"
SIGNUP_REQUIRES_APPROVAL = os.getenv("SIGNUP_REQUIRES_APPROVAL", _approval_default).strip().lower() in {
    "1",
    "true",
    "yes",
}

SIGNUP_ALLOWED_EMAIL_DOMAIN = os.getenv(
    "SIGNUP_ALLOWED_EMAIL_DOMAIN",
    "@company.example",
).strip().lower()

REFERENCE_INTERNAL_ACCOUNT_LOOKUP_MODE = os.getenv(
    "REFERENCE_INTERNAL_ACCOUNT_LOOKUP_MODE",
    "test_double" if APP_PROFILE == "demo" else "api",
).strip().lower()
REFERENCE_INTERNAL_ACCOUNT_TEST_EMAILS = _load_csv_set("REFERENCE_INTERNAL_ACCOUNT_TEST_EMAILS")
REFERENCE_INTERNAL_ACCOUNT_LOOKUP_URL = os.getenv(
    "REFERENCE_INTERNAL_ACCOUNT_LOOKUP_URL",
    "https://inworx.inixindojogja.co.id/api/Resource/dataset",
).strip()
REFERENCE_INTERNAL_ACCOUNT_LOOKUP_USERNAME = os.getenv(
    "REFERENCE_INTERNAL_ACCOUNT_LOOKUP_USERNAME",
    os.getenv("INTERNAL_API_USERNAME", ""),
).strip()
REFERENCE_INTERNAL_ACCOUNT_LOOKUP_PASSWORD = os.getenv(
    "REFERENCE_INTERNAL_ACCOUNT_LOOKUP_PASSWORD",
    os.getenv("INTERNAL_API_PASSWORD", ""),
)
REFERENCE_INTERNAL_ACCOUNT_LOOKUP_TIMEOUT_SECONDS = _int_env(
    "REFERENCE_INTERNAL_ACCOUNT_LOOKUP_TIMEOUT_SECONDS",
    10,
)

AUTH_SIGNUP_VERIFICATION_DELIVERY_MODE = os.getenv(
    "AUTH_SIGNUP_VERIFICATION_DELIVERY_MODE",
    "capture" if APP_PROFILE == "demo" else "webhook",
).strip().lower()
AUTH_SIGNUP_VERIFICATION_WEBHOOK_URL = os.getenv(
    "AUTH_SIGNUP_VERIFICATION_WEBHOOK_URL",
    "",
).strip()
AUTH_SIGNUP_VERIFICATION_TIMEOUT_SECONDS = _int_env(
    "AUTH_SIGNUP_VERIFICATION_TIMEOUT_SECONDS",
    10,
)

_secure_default = "1" if APP_PROFILE != "demo" else "0"
SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", _secure_default).strip().lower() in {
    "1",
    "true",
    "yes",
}

SESSION_IDLE_TIMEOUT_SECONDS = _int_env("SESSION_IDLE_TIMEOUT_SECONDS", 1800)
SESSION_ACTIVITY_TOUCH_SECONDS = _int_env("SESSION_ACTIVITY_TOUCH_SECONDS", 60)
SESSION_MAX_ACTIVE_PER_USER = _int_env("SESSION_MAX_ACTIVE_PER_USER", 1)
SESSION_MAX_ACTIVE_TOTAL = _int_env("SESSION_MAX_ACTIVE_TOTAL", 24)
REPORT_ARTIFACT_DIR = os.getenv(
    "REPORT_ARTIFACT_DIR",
    os.path.join(DATA_DIR, "generated_reports"),
)
JOB_STATE_PATH = os.getenv(
    "JOB_STATE_PATH",
    os.path.join(DATA_DIR, "report_jobs.json"),
)
REPORT_JOB_WORKERS = _int_env("REPORT_JOB_WORKERS", 3)
REPORT_MAX_PENDING_JOBS = _int_env("REPORT_MAX_PENDING_JOBS", 24)
REPORT_JOB_RETENTION_SECONDS = _int_env("REPORT_JOB_RETENTION_SECONDS", 86400)

CORS_ALLOWED_ORIGINS = os.getenv("CORS_ALLOWED_ORIGINS", "")

INTERNAL_API_BASE_URL = os.getenv("INTERNAL_API_BASE_URL", "").rstrip("/")
INTERNAL_API_KEY = os.getenv("INTERNAL_API_KEY", "")
INTERNAL_API_FEEDBACK_ENDPOINT = os.getenv(
    "INTERNAL_API_FEEDBACK_ENDPOINT",
    "/feedback",
)
INTERNAL_API_TIMEOUT_SECONDS = _int_env("INTERNAL_API_TIMEOUT_SECONDS", 20)
INTERNAL_API_AUTH_MODE = os.getenv("INTERNAL_API_AUTH_MODE", "api_key").strip().lower()
INTERNAL_API_AUTH_HEADER = os.getenv("INTERNAL_API_AUTH_HEADER", "Authorization").strip() or "Authorization"
INTERNAL_API_AUTH_PREFIX = os.getenv("INTERNAL_API_AUTH_PREFIX", "Bearer").strip()
INTERNAL_API_USERNAME = os.getenv("INTERNAL_API_USERNAME", "").strip()
INTERNAL_API_PASSWORD = os.getenv("INTERNAL_API_PASSWORD", "")
INTERNAL_API_SOURCE_URL = os.getenv("INTERNAL_API_SOURCE_URL", "").strip()
INTERNAL_API_SOURCE_METHOD = os.getenv("INTERNAL_API_SOURCE_METHOD", "GET").strip().upper() or "GET"
INTERNAL_API_SOURCE_BODY_MODE = os.getenv("INTERNAL_API_SOURCE_BODY_MODE", "json").strip().lower() or "json"
INTERNAL_CONNECTOR_PATH = os.getenv(
    "INTERNAL_CONNECTOR_PATH",
    os.path.join(BASE_DIR, "internal_connector.production.json"),
)
ENABLE_VECTOR_INDEX = os.getenv("ENABLE_VECTOR_INDEX", "0").strip().lower() in {
    "1",
    "true",
    "yes",
}

# ── Score Engine Formula Parameters ──
# Safely parse float environmental variables to prevent crashes at import time
SCORE_BASE_WEIGHT = _float_env("SCORE_BASE_WEIGHT", 0.72)
SCORE_BALANCE_WEIGHT = _float_env("SCORE_BALANCE_WEIGHT", 0.28)
SCORE_POS_FACTOR = _float_env("SCORE_POS_FACTOR", 6.0)
SCORE_NEG_FACTOR = _float_env("SCORE_NEG_FACTOR", 11.0)
SCORE_RISK_PENALTY_SCALE = _float_env("SCORE_RISK_PENALTY_SCALE", 4.0)
SCORE_RISK_PENALTY_MAX = _float_env("SCORE_RISK_PENALTY_MAX", 3.0)
SCORE_DIRECTION_THRESHOLD = _float_env("SCORE_DIRECTION_THRESHOLD", 0.6)


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

WRITER_FIRM_NAME = "Inixindo Jogja - Divisi Penjaminan Mutu dan Pengalaman Pelanggan"
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
    {"id": "mixed", "label": "Kritik Konstruktif"},
    {"id": "neutral", "label": "Netral"},
    {"id": "negative", "label": "Negatif"},
    {"id": "weak_negative", "label": "Negatif Bukti Lemah"},
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
        "description": "Membaca gabungan touchpoint pelanggan, pengalaman yang dirasakan, dan perjalanan peserta saat mengikuti agenda perusahaan.",
    },
]

SCORE_ENGINE_PARAMETER_SOURCE = "Feedback Score.xlsx (2026-04-20)"
SCORE_ENGINE_PARAMETER_TABLES = {
    "learning_score": [
        {"dimension": "Instructor Delivery", "indicator": "Kejelasan penyampaian", "weight_pct": 7.5},
        {"dimension": "Instructor Delivery", "indicator": "Struktur materi", "weight_pct": 7.5},
        {"dimension": "Instructor Delivery", "indicator": "Kemampuan menjawab pertanyaan", "weight_pct": 7.5},
        {"dimension": "Instructor Delivery", "indicator": "Penguasaan topik", "weight_pct": 7.5},
        {"dimension": "Engagement", "indicator": "Interaksi dengan peserta", "weight_pct": 6.25},
        {"dimension": "Engagement", "indicator": "Menjaga perhatian", "weight_pct": 6.25},
        {"dimension": "Engagement", "indicator": "Diskusi & partisipasi", "weight_pct": 6.25},
        {"dimension": "Engagement", "indicator": "Keterlibatan aktif", "weight_pct": 6.25},
        {"dimension": "Relevance (Sebelum & saat training)", "indicator": "Materi sesuai pekerjaan (silabus / studi kasus / TNA)", "weight_pct": 12.5},
        {"dimension": "Relevance (Sebelum & saat training)", "indicator": "Problem Fit (membantu menyelesaikan masalah)", "weight_pct": 12.5},
        {"dimension": "Learning Outcome (Setelah training)", "indicator": "Peningkatan skill", "weight_pct": 10.0},
        {"dimension": "Learning Outcome (Setelah training)", "indicator": "Kemampuan penerapan dalam pekerjaan (capability)", "weight_pct": 10.0},
    ],
    "service_score": [
        {"dimension": "Attitude", "indicator": "Keramahan", "weight_pct": 10.0},
        {"dimension": "Attitude", "indicator": "Empati", "weight_pct": 10.0},
        {"dimension": "Attitude", "indicator": "Kesopanan", "weight_pct": 10.0},
        {"dimension": "Responsiveness", "indicator": "Kecepatan respon", "weight_pct": 10.0},
        {"dimension": "Responsiveness", "indicator": "Ketersediaan bantuan", "weight_pct": 10.0},
        {"dimension": "Competence", "indicator": "Pengetahuan staff", "weight_pct": 10.0},
        {"dimension": "Competence", "indicator": "Problem solving", "weight_pct": 10.0},
        {"dimension": "Transport (antar jemput)", "indicator": "Ketepatan waktu", "weight_pct": 7.5},
        {"dimension": "Transport (antar jemput)", "indicator": "Kenyamanan perjalanan", "weight_pct": 7.5},
        {"dimension": "Souvenir", "indicator": "Relevansi", "weight_pct": 7.5},
        {"dimension": "Souvenir", "indicator": "Kualitas", "weight_pct": 7.5},
    ],
    "facility_score": [
        {"dimension": "Classroom Comfort", "indicator": "Kursi & meja", "weight_pct": 7.5},
        {"dimension": "Classroom Comfort", "indicator": "Pencahayaan", "weight_pct": 7.5},
        {"dimension": "Classroom Comfort", "indicator": "Suhu ruangan", "weight_pct": 7.5},
        {"dimension": "Classroom Comfort", "indicator": "Kebersihan", "weight_pct": 7.5},
        {"dimension": "Equipment", "indicator": "Proyektor / screen", "weight_pct": 6.25},
        {"dimension": "Equipment", "indicator": "Audio / Noise", "weight_pct": 6.25},
        {"dimension": "Equipment", "indicator": "Internet", "weight_pct": 6.25},
        {"dimension": "Equipment", "indicator": "Perangkat lab", "weight_pct": 6.25},
        {"dimension": "Supporting Facilities", "indicator": "Toilet", "weight_pct": 6.25},
        {"dimension": "Supporting Facilities", "indicator": "Musholla", "weight_pct": 6.25},
        {"dimension": "Supporting Facilities", "indicator": "Area istirahat", "weight_pct": 6.25},
        {"dimension": "Supporting Facilities", "indicator": "Konsumsi", "weight_pct": 6.25},
        {"dimension": "Accessibility", "indicator": "Parkir", "weight_pct": 5.0},
        {"dimension": "Accessibility", "indicator": "Akses lokasi", "weight_pct": 5.0},
        {"dimension": "Accessibility", "indicator": "Signage", "weight_pct": 10.0},
    ],
    "experience_index": [
        {"dimension": "Component", "indicator": "Learning Score", "weight_pct": 50.0},
        {"dimension": "Component", "indicator": "Service Score", "weight_pct": 30.0},
        {"dimension": "Component", "indicator": "Facility Score", "weight_pct": 20.0},
    ],
}

SCORE_ENGINE_PROFILES = {
    "learning_score": {
        "label": "Learning Score",
        "summary_label": "kualitas pembelajaran",
        "narrative_focus": "kualitas instruktur, relevansi materi, kenyamanan belajar, dan hasil yang dirasakan peserta",
        "forecast_label": "Learning Score",
        "parameter_source": SCORE_ENGINE_PARAMETER_SOURCE,
        "theme_weights": {
            "instructor": 0.475,
            "material": 0.15,
            "outcome": 0.30,
            "communication": 0.075,
            "schedule": 0.0,
            "facility": 0.0,
            "responsiveness": 0.0,
        },
    },
    "service_score": {
        "label": "Service Score",
        "summary_label": "kualitas layanan",
        "narrative_focus": "responsiveness, koordinasi, ketepatan tindak lanjut, dan kualitas eksekusi layanan",
        "forecast_label": "Service Score",
        "parameter_source": SCORE_ENGINE_PARAMETER_SOURCE,
        "theme_weights": {
            "responsiveness": 0.35,
            "communication": 0.25,
            "outcome": 0.10,
            "schedule": 0.10,
            "facility": 0.10,
            "material": 0.10,
            "instructor": 0.0,
        },
    },
    "facility_score": {
        "label": "Facility Score",
        "summary_label": "kesiapan fasilitas",
        "narrative_focus": "fasilitas kelas, jaringan, ruang, sarana pendukung, dan kesiapan operasional sebelum delivery",
        "forecast_label": "Facility Score",
        "parameter_source": SCORE_ENGINE_PARAMETER_SOURCE,
        "theme_weights": {
            "facility": 0.95,
            "schedule": 0.05,
            "communication": 0.0,
            "responsiveness": 0.0,
            "instructor": 0.0,
            "material": 0.0,
            "outcome": 0.0,
        },
    },
    "experience_index": {
        "label": "Experience Index",
        "summary_label": "pengalaman pelanggan lintas touchpoint, rasa layanan, dan perjalanan agenda",
        "narrative_focus": "gabungan touchpoint pelanggan, pengalaman yang dirasakan, dan perjalanan peserta saat mengikuti agenda perusahaan dari pra-layanan hingga outcome pasca-layanan",
        "forecast_label": "Experience Index",
        "parameter_source": SCORE_ENGINE_PARAMETER_SOURCE,
        "component_weights": {
            "learning_score": 0.50,
            "service_score": 0.30,
            "facility_score": 0.20,
        },
        "theme_weights": {
            "responsiveness": 0.105,
            "communication": 0.1125,
            "schedule": 0.04,
            "facility": 0.22,
            "instructor": 0.2375,
            "material": 0.105,
            "outcome": 0.18,
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
OSINT_QUERY_WORKERS = int(os.getenv("OSINT_QUERY_WORKERS", "4"))
OSINT_DEEP_SCRAPE_MAX_CHARS = int(os.getenv("OSINT_DEEP_SCRAPE_MAX_CHARS", "5000"))
OSINT_TRUSTED_DOMAINS = tuple(
    _load_csv_list(
        "OSINT_TRUSTED_DOMAINS",
        [
            "bps.go.id",
            "kominfo.go.id",
            "worldbank.org",
            "mckinsey.com",
            "deloitte.com",
            "pwc.com",
            "idc.com",
            "gartner.com",
            "coursera.org",
            "skillsoft.com",
            "trainingindustry.com",
        ],
    )
)
OSINT_BLOCKED_DOMAINS = tuple(
    _load_csv_list(
        "OSINT_BLOCKED_DOMAINS",
        [
            "facebook.com",
            "instagram.com",
            "tiktok.com",
            "pinterest.com",
            "shopee.co.id",
            "tokopedia.com",
        ],
    )
)
OSINT_CACHE_PATH = os.getenv(
    "OSINT_CACHE_PATH",
    os.path.join(DATA_DIR, "osint_cache.json"),
)
OSINT_CACHE_DIR = os.getenv(
    "OSINT_CACHE_DIR",
    os.path.join(os.path.dirname(OSINT_CACHE_PATH), ".osint_cache"),
)
OSINT_CACHE_TTL_SECONDS = int(os.getenv("OSINT_CACHE_TTL_SECONDS", "21600"))
OSINT_BASE_QUERIES = [
    "tren pelatihan IT corporate Indonesia",
    "ekspektasi peserta training IT terhadap instruktur fasilitas dan kurikulum Indonesia",
    "tantangan transformasi digital dan peningkatan kompetensi SDM Indonesia",
    "tren kebutuhan sertifikasi cloud cyber security data dan AI di Indonesia",
    "benchmark corporate learning customer experience Indonesia",
    "digital talent skill gap Indonesia enterprise training",
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
        "id": "cx_chap_1", "title": "BAB I – ANALITIK DESKRIPTIF DAN TATA KELOLA UMPAN BALIK",
        "sections": [
            "1.1 Ringkasan Cakupan Umpan Balik dan Tata Kelola",
            "1.2 Distribusi Sentimen, Penilaian, dan Volume",
            "1.3 Distribusi Pemangku Kepentingan, Layanan, dan Kanal/Sumber"
        ],
        "focus_keywords": "feedback governance descriptive analytics rating stakeholder service channel source",
        "visual": "bar_chart"
    },
    {
        "id": "cx_chap_2", "title": "BAB II – ANALITIK DIAGNOSTIK",
        "sections": [
            "2.1 Akar Masalah Utama dan Titik Keluhan Dominan",
            "2.2 Kekuatan yang Konsisten dan Area yang Perlu Dijaga",
            "2.3 Bukti Verbatim, Kesenjangan Proses, dan Segmentasi Masalah"
        ],
        "focus_keywords": "diagnostic analytics root cause complaint praise service quality process gap"
    },
    {
        "id": "cx_chap_3", "title": "BAB III – ANALITIK PREDIKTIF",
        "sections": [
            "3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut",
            "3.2 Prediksi Segmen dan Layanan yang Paling Rentan",
            "3.3 Tren Eksternal yang Berpotensi Memperbesar Risiko"
        ],
        "focus_keywords": "predictive analytics risk trend forecast segment service vulnerability"
    },
    {
        "id": "cx_chap_4", "title": "BAB IV – ANALITIK PRESKRIPTIF",
        "sections": [
            "4.1 Intervensi Prioritas 30 Hari",
            "4.2 Penguatan Tata Kelola Umpan Balik dan Eskalasi",
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

CANONICAL_INTERNAL_COLUMNS = (
    "Record ID",
    "Sumber Feedback",
    "Kanal Feedback",
    "Tanggal Feedback",
    "Tipe Stakeholder",
    "Layanan",
    "Lokasi",
    "Tipe Instruktur",
    "Rentang Waktu",
    "Rating",
    "Komentar",
    "Customer Journey Hint",
)

APIDOG_TIMELESS_TIMEFRAME = "Semua Data APIDog (tanggal tidak tersedia)"
APIDOG_UNKNOWN_DATE = "Tanggal tidak tersedia"
APIDOG_CLASS_REPORT_CHANNEL = "Evaluasi Kelas Internal"

CLASS_REPORT_METADATA_COLUMNS = (
    "Raw Response Count",
    "Rating Response Count",
    "Text Response Count",
    "Rating Distribution",
    "Representative Why",
)

COLUMN_ALIASES = {
    "Record ID": ("record_id", "id", "feedback_id", "ticket_id", "case_id", "uuid", "kode", "no_tiket", "nomor_tiket"),
    "Sumber Feedback": ("sumber feedback", "source", "feedback_source", "origin", "source_name", "sumber", "asal_data", "source_type"),
    "Kanal Feedback": ("kanal feedback", "channel", "feedback_channel", "touchpoint", "platform", "kanal", "media", "channel_name"),
    "Tanggal Feedback": ("tanggal feedback", "feedback_date", "created_at", "submitted_at", "date", "tanggal", "tanggal_submit", "tgl_feedback", "created_date"),
    "Tipe Stakeholder": ("tipe stakeholder", "stakeholder_type", "stakeholder", "customer_segment", "customer_type", "segment", "segmen", "tipe_pelanggan", "jenis_pelanggan", "kategori_peserta", "instansi_type"),
    "Layanan": ("layanan", "service", "service_name", "product", "offering", "service_type", "nama_layanan", "program", "course", "training_name", "kelas", "judul_pelatihan"),
    "Lokasi": ("lokasi", "location", "training_location", "city", "kota", "venue_location", "tempat", "cabang", "venue"),
    "Tipe Instruktur": ("tipe instruktur", "instructor_type", "trainer_type", "coach_type", "internal_ol", "internal_or_ol", "trainer_origin", "jenis_instruktur", "tipe_trainer", "pengajar_type"),
    "Rentang Waktu": ("rentang waktu", "timeframe", "periode", "period", "reporting_period", "bulan", "semester", "tahun", "periode_laporan"),
    "Rating": ("rating", "score", "csat", "sentiment_score", "nilai", "skor", "bintang", "kepuasan", "satisfaction_score"),
    "Komentar": ("komentar", "comment", "feedback", "feedback_text", "review", "notes", "complaint_text", "customer_comment", "ulasan", "saran", "kritik", "pesan", "testimoni", "isi_feedback"),
    "Customer Journey Hint": ("customer_journey_hint", "journey_hint", "journey_stage", "customer_journey_stage", "touchpoint_stage", "tahap_journey", "fase_layanan"),
}

DATE_COLUMN_ALIASES = (
    "tanggal feedback",
    "tanggal",
    "date",
    "created_at",
    "submitted_at",
    "feedback_date",
)

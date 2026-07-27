"""Bounded feedback policy for UC2 report generation."""
from adaptive_feedback import FeedbackStore, build_feedback_policy
from config import LEARNING_FEEDBACK_DB_PATH

UC2_FEEDBACK_POLICY = build_feedback_policy("uc2_feedback", {
    "hasil_membantu": {"label": "Laporan membantu", "guidance": "", "mode": "positive"},
    "ringkasan_kurang_jelas": {"label": "Ringkasan kurang jelas", "guidance": "Buat ringkasan dan implikasi utama lebih langsung, spesifik, dan mudah ditindaklanjuti.", "mode": "adapt"},
    "bahasa_terlalu_teknis": {"label": "Bahasa terlalu teknis", "guidance": "Gunakan Bahasa Indonesia yang lebih sederhana tanpa menghilangkan angka, bukti, batasan, atau makna analitis.", "mode": "adapt"},
    "rekomendasi_kurang_praktis": {"label": "Rekomendasi kurang praktis", "guidance": "Jelaskan tindakan, penanggung jawab, ukuran selesai, dan kaitannya dengan bukti secara lebih praktis.", "mode": "adapt"},
    "fokus_kurang_sesuai": {"label": "Fokus laporan kurang sesuai", "guidance": "Prioritaskan tema yang paling relevan dengan ruang lingkup laporan tanpa mengubah fakta atau hasil analisis.", "mode": "adapt"},
    "data_perlu_diperiksa": {"label": "Data atau angka perlu diperiksa", "guidance": "", "mode": "review"},
})

def create_feedback_store():
    return FeedbackStore(LEARNING_FEEDBACK_DB_PATH, UC2_FEEDBACK_POLICY)

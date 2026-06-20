"""Distilled exemplar guidance for feedback-report writing.

The source PDFs are not application evidence. This module keeps only compact
style and analysis guidance so generated UC2 reports can stay coherent without
copying source text, client names, or document-specific facts.
"""
from __future__ import annotations

import copy
from typing import Any


UC2_FEEDBACK_EXEMPLAR_PROFILE_VERSION = "uc2-feedback-exemplar-profile-v1"


_UC2_FEEDBACK_EXEMPLAR_PROFILE: dict[str, Any] = {
    "version": UC2_FEEDBACK_EXEMPLAR_PROFILE_VERSION,
    "hardcoded_structure_policy": "use_existing_report_structure_only",
    "source_mix": [
        {
            "kind": "customer_intelligence_report",
            "calibration": "Gunakan alur ringkasan, cakupan, metode, temuan, rekomendasi, dan batas data secara eksplisit.",
        },
        {
            "kind": "theme_and_sentiment_analysis_deck",
            "calibration": "Susun tema, kategori, volume, sentimen, pendorong utama, anomali, dan langkah berikutnya sebagai satu rantai keputusan.",
        },
        {
            "kind": "indonesian_statistical_bulletin",
            "calibration": "Tulis perubahan angka secara padat dengan konteks periode, kelompok responden, arah pergerakan, dan batas interpretasi.",
        },
        {
            "kind": "customer_service_compliance_report",
            "calibration": "Jaga kaitan layanan, keluhan, kepuasan, tanggung jawab, dan pemantauan tindak lanjut.",
        },
    ],
    "analysis_moves": [
        "executive_snapshot",
        "scope_and_method",
        "data_overview",
        "theme_bank",
        "sentiment_volume_movement",
        "segment_or_demographic_reading",
        "key_driver_and_anomaly_check",
        "service_gap_and_data_gap",
        "recommendation_and_next_steps",
        "measurement_obligation",
    ],
    "theme_sentiment_rules": [
        "Tema harus dibaca bersama volume, intensitas sentimen, dan contoh komentar yang mewakili pola.",
        "Kenaikan volume komentar negatif belum otomatis berarti penyebab layanan; periksa cakupan, segmen, dan periode.",
        "Kategori tema harus membantu keputusan, bukan sekadar mengulang kata yang sering muncul.",
        "Anomali boleh disebut hanya jika ada pembanding internal yang jelas atau cukup dijelaskan sebagai sinyal awal.",
    ],
    "indonesian_language_rules": [
        "Gunakan Bahasa Indonesia laporan yang lugas: tercatat, meningkat, menurun, relatif stabil, dibandingkan periode sebelumnya.",
        "Jelaskan angka dengan satuan dan basis responden sebelum menarik makna manajerial.",
        "Hindari rasa terjemahan langsung; tulis sebagai analisis layanan yang memang dibuat untuk pembaca Indonesia.",
        "Variasikan transisi antarbagian agar bab terasa menyambung, bukan seperti template yang diulang.",
    ],
    "factual_boundaries": [
        "Contoh dokumen adalah kalibrasi gaya dan struktur analisis, bukan bukti untuk laporan.",
        "Gunakan hanya data feedback UC2, filter aktif, komentar peserta, dan konteks internal yang tersedia.",
        "Jangan memindahkan nama organisasi, judul sumber, angka sumber, kutipan sumber, atau kerangka sumber ke laporan.",
        "Jika pembanding eksternal tidak tersedia, sebutkan batasnya dan jangan membuat klaim pasar atau benchmark.",
    ],
    "coherence_rules": [
        "Ringkasan eksekutif harus menyiapkan masalah utama yang kemudian dibuktikan oleh bab analisis.",
        "Bab diagnostik harus menjelaskan mengapa pola deskriptif penting, bukan membuka isu baru tanpa jembatan.",
        "Rekomendasi harus menjawab temuan sebelumnya dengan pemilik, ukuran perubahan, dan waktu tinjauan.",
        "Lampiran menyimpan metodologi dan batas data agar narasi utama tetap terbaca.",
    ],
}


def build_uc2_feedback_exemplar_profile() -> dict[str, Any]:
    """Return a defensive copy so callers cannot mutate the shared profile."""
    return copy.deepcopy(_UC2_FEEDBACK_EXEMPLAR_PROFILE)

#!/usr/bin/env python3
import sqlite3
from datetime import date, timedelta
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CSV_PATH = DATA_DIR / "db.csv"
DB_PATH = DATA_DIR / "cx_feedback.db"

TIMEFRAME_SPECS = [
    {"label": "1 Minggu Terakhir (Weekly)", "age_window": (1, 7)},
    {"label": "1 Bulan Terakhir (Monthly)", "age_window": (8, 30)},
    {"label": "6 Bulan Terakhir (Biannual)", "age_window": (31, 180)},
    {"label": "1 Tahun Terakhir (Yearly)", "age_window": (181, 360)},
]

SEGMENT_SPECS = [
    {
        "label": "Instansi Pemerintah (Gov)",
        "short_label": "instansi pemerintah",
        "locations": ["Yogyakarta", "Jakarta"],
        "services": [
            "Pelatihan IT Security",
            "Audit SPBE",
            "Workshop Tata Kelola Data",
        ],
        "owner": "Public Sector Account Lead",
    },
    {
        "label": "BUMN / Corporate",
        "short_label": "BUMN / corporate",
        "locations": ["Jakarta", "Bandung", "Surabaya"],
        "services": [
            "Konsultasi IT Masterplan",
            "Pelatihan Cloud Native (AWS)",
            "Workshop AI Adoption Readiness",
        ],
        "owner": "Corporate Delivery Lead",
    },
    {
        "label": "Mahasiswa / Personal",
        "short_label": "personal",
        "locations": ["Yogyakarta", "Online Live"],
        "services": [
            "Sertifikasi Cisco (CCNA)",
            "Pelatihan Web Dev",
            "Bootcamp Data Science",
        ],
        "owner": "Retail Program Coordinator",
    },
    {
        "label": "Sekolah / Universitas",
        "short_label": "sekolah / universitas",
        "locations": ["Semarang", "Yogyakarta"],
        "services": [
            "Program Magang / Kunjungan Industri",
            "Pelatihan Dasar Cybersecurity",
            "Bootcamp Data untuk Dosen",
        ],
        "owner": "Education Partnership Manager",
    },
    {
        "label": "Vendor / Partner",
        "short_label": "vendor / partner",
        "locations": ["Jakarta", "Online Live"],
        "services": [
            "Kerja Sama Event",
            "Enablement Partner Teknologi",
            "Technical Sales Workshop",
        ],
        "owner": "Alliance Manager",
    },
]

THEME_SPECS = [
    {
        "id": "responsiveness",
        "source": "Account Review",
        "channels": ["WhatsApp", "Email"],
        "journey": "Pra-Layanan dan Ekspektasi",
        "positive": (
            "Tim account untuk {service} cepat merespons pertanyaan dan follow up sangat jelas, "
            "sehingga koordinasi dengan {segment} berjalan lancar sejak awal."
        ),
        "neutral": (
            "Respons admin untuk {service} cukup cepat, tetapi follow up revisi dan update status "
            "masih perlu dibuat lebih konsisten untuk {segment}."
        ),
        "negative": (
            "Tim account untuk {service} lambat merespons permintaan revisi dan follow up dari {segment}, "
            "sehingga timeline terasa mundur dari SLA yang dijanjikan."
        ),
    },
    {
        "id": "schedule",
        "source": "Project Coordination Log",
        "channels": ["Email", "Google Form"],
        "journey": "Persiapan dan Kesiapan Delivery",
        "positive": (
            "Jadwal {service} tertata rapi, durasi sesi pas, dan jeda antarsesi cukup untuk menjaga fokus peserta {segment}."
        ),
        "neutral": (
            "Jadwal {service} cukup sesuai, tetapi durasi sesi siang terasa padat dan jeda masih bisa diperbaiki untuk peserta {segment}."
        ),
        "negative": (
            "Jadwal {service} terlalu padat, perubahan waktu mendadak, dan jeda antarsesi kurang memadai bagi peserta {segment}."
        ),
    },
    {
        "id": "facility",
        "source": "Operational Readiness Check",
        "channels": ["On-site Form", "Google Form"],
        "journey": "Persiapan dan Kesiapan Delivery",
        "positive": (
            "Fasilitas ruang untuk {service} nyaman, lab siap dipakai, dan wifi stabil sepanjang sesi untuk {segment}."
        ),
        "neutral": (
            "Ruang {service} cukup nyaman, namun jaringan sesekali melambat saat peserta {segment} mengakses lab dan package."
        ),
        "negative": (
            "Fasilitas ruang untuk {service} kurang siap; wifi sering putus, AC tidak stabil, dan setup lab terlambat saat sesi dimulai untuk {segment}."
        ),
    },
    {
        "id": "instructor",
        "source": "Post-Class Survey",
        "channels": ["Google Form", "LMS"],
        "journey": "Pelaksanaan Layanan",
        "positive": (
            "Instruktur {service} sangat kompeten, penjelasannya jelas, dan contoh kasusnya dekat dengan kebutuhan {segment}."
        ),
        "neutral": (
            "Instruktur {service} cukup menguasai materi, tetapi penyampaian contoh kasus belum selalu relevan dengan kebutuhan {segment}."
        ),
        "negative": (
            "Instruktur {service} kurang konsisten menjawab pertanyaan lanjutan dan beberapa penjelasan terasa terlalu umum untuk kebutuhan {segment}."
        ),
    },
    {
        "id": "material",
        "source": "Learning Evaluation",
        "channels": ["LMS", "Google Form"],
        "journey": "Pelaksanaan Layanan",
        "positive": (
            "Materi {service} relevan, modulnya terstruktur, dan contoh implementasinya mudah diterapkan setelah sesi selesai oleh {segment}."
        ),
        "neutral": (
            "Materi {service} cukup jelas, tetapi beberapa modul masih perlu diperbarui agar lebih relevan dengan kondisi terbaru untuk {segment}."
        ),
        "negative": (
            "Materi {service} terasa terlalu basic, beberapa modul out of date, dan contoh implementasinya belum sesuai konteks {segment}."
        ),
    },
    {
        "id": "communication",
        "source": "Stakeholder Debrief",
        "channels": ["Email", "CRM"],
        "journey": "Pra-Layanan dan Ekspektasi",
        "positive": (
            "Komunikasi pra-delivery {service} rapi, brief jelas, dan semua update status diterima tepat waktu oleh tim {segment}."
        ),
        "neutral": (
            "Komunikasi untuk {service} cukup jelas, namun update perubahan kecil belum selalu dibagikan ke seluruh PIC {segment}."
        ),
        "negative": (
            "Komunikasi {service} kurang rapi; informasi perubahan jadwal, brief teknis, dan update status tidak selalu tersampaikan ke semua PIC {segment}."
        ),
    },
    {
        "id": "outcome",
        "source": "Outcome Review",
        "channels": ["Interview", "CRM"],
        "journey": "Tindak Lanjut dan Outcome",
        "positive": (
            "Hasil {service} terasa actionable, membantu tim {segment} menindaklanjuti pekerjaan, dan output akhirnya bisa langsung dipakai."
        ),
        "neutral": (
            "Hasil {service} cukup membantu, tetapi tindak lanjut pasca-sesi masih perlu pendampingan agar manfaatnya lebih terasa untuk {segment}."
        ),
        "negative": (
            "Setelah {service} selesai, manfaat yang dirasakan belum kuat, tindak lanjut kurang jelas, dan output akhirnya belum mudah dipakai oleh tim {segment}."
        ),
    },
]

SENTIMENT_PATTERNS = {
    "1 Minggu Terakhir (Weekly)": [
        "positive",
        "negative",
        "negative",
        "neutral",
        "positive",
        "neutral",
        "negative",
    ],
    "1 Bulan Terakhir (Monthly)": [
        "positive",
        "neutral",
        "negative",
        "positive",
        "negative",
        "neutral",
        "positive",
    ],
    "6 Bulan Terakhir (Biannual)": [
        "positive",
        "positive",
        "neutral",
        "negative",
        "positive",
        "neutral",
        "negative",
    ],
    "1 Tahun Terakhir (Yearly)": [
        "positive",
        "positive",
        "negative",
        "positive",
        "neutral",
        "positive",
        "negative",
    ],
}

RATING_MAP = {
    "positive": [4, 5],
    "neutral": [3],
    "negative": [1, 2],
}


def build_feedback_records():
    today = date.today()
    records = []
    record_counter = 1

    for segment_index, segment in enumerate(SEGMENT_SPECS):
        for timeframe_index, timeframe in enumerate(TIMEFRAME_SPECS):
            label = timeframe["label"]
            min_age, max_age = timeframe["age_window"]
            sentiment_pattern = SENTIMENT_PATTERNS[label]

            for theme_index, theme in enumerate(THEME_SPECS):
                sentiment = sentiment_pattern[(segment_index + theme_index) % len(sentiment_pattern)]
                service = segment["services"][(timeframe_index + theme_index) % len(segment["services"])]
                rating_options = RATING_MAP[sentiment]
                rating = rating_options[(segment_index + timeframe_index + theme_index) % len(rating_options)]
                age_span = max_age - min_age + 1
                age_days = min_age + ((segment_index * 13 + timeframe_index * 7 + theme_index * 3) % age_span)
                feedback_date = today - timedelta(days=age_days)
                channel = theme["channels"][(segment_index + timeframe_index) % len(theme["channels"])]
                comment_template = theme[sentiment]
                location = segment["locations"][
                    (segment_index + timeframe_index + theme_index) % len(segment["locations"])
                ]
                instructor_type = (
                    "Internal"
                    if (segment_index + timeframe_index + theme_index) % 3 != 0
                    else "OL"
                )

                records.append(
                    {
                        "Record ID": f"DEMO-{today.strftime('%Y%m%d')}-{record_counter:04d}",
                        "Sumber Feedback": theme["source"],
                        "Kanal Feedback": channel,
                        "Tanggal Feedback": feedback_date.isoformat(),
                        "Tipe Stakeholder": segment["label"],
                        "Layanan": service,
                        "Rentang Waktu": label,
                        "Rating": rating,
                        "Komentar": comment_template.format(
                            service=service,
                            segment=segment["short_label"],
                        ),
                        "Tema Feedback": theme["id"],
                        "Customer Journey Hint": theme["journey"],
                        "PIC Layanan": segment["owner"],
                        "Lokasi": location,
                        "Tipe Instruktur": instructor_type,
                    }
                )
                record_counter += 1

            records.append(
                {
                    "Record ID": f"DEMO-{today.strftime('%Y%m%d')}-{record_counter:04d}",
                    "Sumber Feedback": "Executive Follow-up",
                    "Kanal Feedback": "Meeting Notes",
                    "Tanggal Feedback": (today - timedelta(days=min_age)).isoformat(),
                    "Tipe Stakeholder": segment["label"],
                    "Layanan": segment["services"][timeframe_index % len(segment["services"])],
                    "Rentang Waktu": label,
                    "Rating": 4 if timeframe_index % 2 == 0 else 3,
                    "Komentar": (
                        "Secara umum pengalaman layanan sudah cukup baik, tetapi organisasi masih meminta komunikasi, "
                        "jadwal, dan tindak lanjut hasil dibuat lebih konsisten agar manfaatnya terasa merata."
                    ),
                    "Tema Feedback": "cross_theme",
                    "Customer Journey Hint": "Lintas Tahap",
                    "PIC Layanan": segment["owner"],
                    "Lokasi": segment["locations"][timeframe_index % len(segment["locations"])],
                    "Tipe Instruktur": "Internal" if timeframe_index % 2 == 0 else "OL",
                }
            )
            record_counter += 1

    return records


def save_seed_data(records):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    dataframe = pd.DataFrame(records)
    dataframe = dataframe.sort_values(
        by=["Tanggal Feedback", "Tipe Stakeholder", "Layanan", "Record ID"],
        ascending=[False, True, True, True],
    ).reset_index(drop=True)
    dataframe.to_csv(CSV_PATH, index=False)

    with sqlite3.connect(DB_PATH) as connection:
        dataframe.to_sql("feedback", connection, if_exists="replace", index=False)

    return dataframe


def main():
    records = build_feedback_records()
    dataframe = save_seed_data(records)

    print("Demo seed data refreshed.")
    print(f"CSV   : {CSV_PATH}")
    print(f"SQLite: {DB_PATH}")
    print(f"Rows  : {len(dataframe)}")
    print("Columns:")
    print(", ".join(dataframe.columns.tolist()))


if __name__ == "__main__":
    main()

import re

import pandas as pd

from data_contract import (
    APIDOG_CLASS_REPORT_CHANNEL,
    APIDOG_TIMELESS_TIMEFRAME,
    APIDOG_UNKNOWN_DATE,
    CLASS_REPORT_METADATA_COLUMNS,
)


CLASS_REPORT_LABEL_OVERRIDES = {
    "KESESUAIAN MATERIAL BAHAN AJAR": "Kesesuaian materi bahan ajar",
    "KUALITAS PENYAMPAIAN INSTRUKTUR": "Kualitas penyampaian instruktur",
    "PENGUASAAN MATERI INSTRUKTUR": "Penguasaan materi instruktur",
    "KEMAMPUAN INSTRUKTUR MENJAWAB PERTANYAAN": "Kemampuan instruktur menjawab pertanyaan",
    "FASILITAS RUANG KELAS": "Fasilitas ruang kelas",
    "KUALITAS KONSUMSI": "Kualitas konsumsi",
    "KOMENTAR INSTRUKTUR": "Komentar instruktur",
    "SARAN": "Saran peserta",
}

CLASS_REPORT_JOURNEY_RULES = (
    (("materi", "bahan ajar", "kurikulum", "modul"), "Materi dan kurikulum", "Pelaksanaan Layanan"),
    (("instruktur", "trainer", "pengajar", "penyampaian"), "Kinerja instruktur", "Pelaksanaan Layanan"),
    (("fasilitas", "ruang", "kelas", "lab", "lokasi"), "Fasilitas pelatihan", "Persiapan dan Kesiapan Delivery"),
    (("konsumsi", "makan", "snack", "coffee"), "Hospitality pelatihan", "Persiapan dan Kesiapan Delivery"),
    (("pendaftaran", "administrasi", "sertifikat"), "Administrasi pelatihan", "Tindak Lanjut dan Outcome"),
)


class ClassReportAdapter:
    @staticmethod
    def looks_like_class_report(dataframe):
        columns = {str(column).strip() for column in dataframe.columns}
        return {"response_id", "response_name", "response_answer"}.issubset(columns)

    @staticmethod
    def looks_like_reference_class_report(dataframe):
        columns = {str(column).strip() for column in dataframe.columns}
        return (
            {"response_id", "response_name", "response_type"}.issubset(columns)
            and "response_answer" not in columns
        )

    @staticmethod
    def clean_label(value):
        raw_label = re.sub(r"\s+", " ", str(value or "")).strip(" :-")
        if not raw_label:
            return "Evaluasi kelas"
        override = CLASS_REPORT_LABEL_OVERRIDES.get(raw_label.upper())
        if override:
            return override
        if raw_label.isupper():
            return raw_label.lower().capitalize()
        return raw_label[:1].upper() + raw_label[1:]

    @classmethod
    def question_lookup(cls, dataframe):
        lookup = {}
        if dataframe is None or dataframe.empty:
            return lookup
        for _, row in dataframe.iterrows():
            response_id = str(row.get("response_id") or "").strip()
            if not response_id:
                continue
            lookup[response_id] = {
                "label": cls.clean_label(row.get("response_name")),
                "type": str(row.get("response_type") or "").strip(),
                "parent_id": str(row.get("response_parent_id") or "").strip(),
            }
        return lookup

    @staticmethod
    def semantics(clean_label):
        lowered = str(clean_label or "").lower()
        for keywords, service_label, journey_hint in CLASS_REPORT_JOURNEY_RULES:
            if any(keyword in lowered for keyword in keywords):
                return service_label, journey_hint
        return clean_label or "Evaluasi kelas", "Pelaksanaan Layanan"

    @staticmethod
    def is_rating_response(row):
        response_type = str(row.get("response_type") or "").strip().lower()
        if response_type.startswith("rating"):
            return True
        answer = str(row.get("response_answer") or "").strip()
        return bool(answer) and pd.notna(pd.to_numeric(answer, errors="coerce"))

    @staticmethod
    def is_text_response(row):
        response_type = str(row.get("response_type") or "").strip().lower()
        return response_type == "text"

    @staticmethod
    def format_rating_value(value):
        if pd.isna(value):
            return ""
        rounded = round(float(value), 2)
        if rounded.is_integer():
            return str(int(rounded))
        return str(rounded).rstrip("0").rstrip(".")

    @classmethod
    def question_label(cls, row, reference_lookup=None):
        reference_lookup = reference_lookup or {}
        response_id = str(row.get("response_id") or "").strip()
        reference = reference_lookup.get(response_id, {})
        return cls.clean_label(reference.get("label") or row.get("response_name"))

    @staticmethod
    def dedupe_texts(values, limit=5):
        deduped = []
        seen = set()
        for value in values:
            cleaned = re.sub(r"\s+", " ", str(value or "").strip(" .:-"))
            key = cleaned.lower()
            if not cleaned or key in seen:
                continue
            seen.add(key)
            deduped.append(cleaned)
            if len(deduped) >= limit:
                break
        return deduped

    @classmethod
    def build_row(cls, endpoint_name, row_index, question, rating_value, explanation_texts):
        service_label, journey_hint = cls.semantics(question)
        average_text = f"Rata-rata rating {question}: {round(float(rating_value), 2)} dari 5" if pd.notna(rating_value) else question
        explanations = cls.dedupe_texts(explanation_texts, limit=5)
        why_text = f"Mengapa: {'; '.join(explanations)}" if explanations else "Mengapa: belum ada komentar teks yang terhubung ke rating ini."
        return {
            "Record ID": f"{endpoint_name}-{row_index + 1:05d}",
            "Sumber Feedback": endpoint_name,
            "Kanal Feedback": APIDOG_CLASS_REPORT_CHANNEL,
            "Tanggal Feedback": APIDOG_UNKNOWN_DATE,
            "Tipe Stakeholder": "Peserta Kelas",
            "Layanan": service_label,
            "Lokasi": "",
            "Tipe Instruktur": "",
            "Rentang Waktu": APIDOG_TIMELESS_TIMEFRAME,
            "Rating": cls.format_rating_value(rating_value),
            "Komentar": f"{average_text}. {why_text}",
            "Customer Journey Hint": journey_hint,
            "Raw Response Count": "",
            "Rating Response Count": "",
            "Text Response Count": "",
            "Rating Distribution": "",
            "Representative Why": "; ".join(explanations),
        }

    @classmethod
    def normalize(cls, dataframe, endpoint_name, reference_lookup=None):
        reference_lookup = reference_lookup or {}
        rating_groups = {}
        text_by_parent = {}
        orphan_text_rows = []
        for index, row in dataframe.iterrows():
            response_id = str(row.get("response_id") or "").strip()
            answer = str(row.get("response_answer") or "").strip()
            if not response_id and not answer:
                continue
            if cls.is_rating_response(row):
                rating = pd.to_numeric(answer, errors="coerce")
                if pd.isna(rating):
                    continue
                group = rating_groups.setdefault(
                    response_id,
                    {
                        "first_index": index,
                        "question": cls.question_label(row, reference_lookup),
                        "ratings": [],
                    },
                )
                group["ratings"].append(float(rating))
                continue
            if cls.is_text_response(row):
                parent_id = str(row.get("response_parent_id") or "").strip()
                if parent_id:
                    text_by_parent.setdefault(parent_id, []).append(answer)
                else:
                    orphan_text_rows.append((index, cls.question_label(row, reference_lookup), answer))

        rows = []
        for response_id, group in sorted(rating_groups.items(), key=lambda item: item[1]["first_index"]):
            ratings = group["ratings"]
            average_rating = sum(ratings) / len(ratings) if ratings else float("nan")
            explanations = text_by_parent.get(response_id, [])
            distribution = {}
            for value in ratings:
                label = cls.format_rating_value(value)
                distribution[label] = distribution.get(label, 0) + 1
            row_payload = cls.build_row(endpoint_name, len(rows), group["question"], average_rating, explanations)
            row_payload.update(
                {
                    "Raw Response Count": str(len(ratings) + len(explanations)),
                    "Rating Response Count": str(len(ratings)),
                    "Text Response Count": str(len(explanations)),
                    "Rating Distribution": "; ".join(
                        f"{key}: {distribution[key]}" for key in sorted(distribution, key=lambda item: float(item))
                    ),
                }
            )
            rows.append(row_payload)

        for _, question, answer in orphan_text_rows:
            service_label, journey_hint = cls.semantics(question)
            rows.append(
                {
                    "Record ID": f"{endpoint_name}-{len(rows) + 1:05d}",
                    "Sumber Feedback": endpoint_name,
                    "Kanal Feedback": APIDOG_CLASS_REPORT_CHANNEL,
                    "Tanggal Feedback": APIDOG_UNKNOWN_DATE,
                    "Tipe Stakeholder": "Peserta Kelas",
                    "Layanan": service_label,
                    "Lokasi": "",
                    "Tipe Instruktur": "",
                    "Rentang Waktu": APIDOG_TIMELESS_TIMEFRAME,
                    "Rating": "",
                    "Komentar": f"{question}: {answer}".strip(": "),
                    "Customer Journey Hint": journey_hint,
                    "Raw Response Count": "1",
                    "Rating Response Count": "0",
                    "Text Response Count": "1",
                    "Rating Distribution": "",
                    "Representative Why": answer,
                }
            )

        output = pd.DataFrame(rows)
        for column_name in CLASS_REPORT_METADATA_COLUMNS:
            if column_name not in output.columns:
                output[column_name] = ""
            output[column_name] = output[column_name].fillna("").astype(str).str.strip()
        return output

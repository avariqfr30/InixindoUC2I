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
    (("brand equity", "mengapa inixindo", "menjadi pilihan", "alasan memilih"), "Reputasi dan alasan memilih Inixindo", "Tindak Lanjut dan Outcome"),
    (("materi", "bahan ajar", "kurikulum", "modul"), "Materi dan kurikulum", "Pelaksanaan Layanan"),
    (("instruktur", "trainer", "pengajar", "penyampaian"), "Kinerja instruktur", "Pelaksanaan Layanan"),
    (("fasilitas", "ruang", "kelas", "lab", "lokasi"), "Fasilitas pelatihan", "Persiapan dan Kesiapan Delivery"),
    (("konsumsi", "makan", "snack", "coffee"), "Hospitality pelatihan", "Persiapan dan Kesiapan Delivery"),
    (("pendaftaran", "administrasi", "sertifikat"), "Administrasi pelatihan", "Tindak Lanjut dan Outcome"),
)


class ClassReportAdapter:
    START_DATE_FIELDS = ("class_start_date", "start_date", "tanggal_mulai", "tanggal_awal")
    END_DATE_FIELDS = ("class_end_date", "end_date", "tanggal_selesai", "tanggal_akhir", "created_at", "submitted_at")
    RAW_PROMPT_PATTERNS = (
        r"\bpilih\s+\d+\s+bintang\b",
        r"\buntuk mengisi\b",
        r"\?.*\)",
        r"^[A-Z0-9 _/-]{8,}\s*\(",
    )

    @staticmethod
    def clean_scalar(value):
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = str(value).strip()
        return "" if text.lower() in {"nan", "none", "null", "nat", "<na>"} else text

    @staticmethod
    def looks_like_class_report(dataframe):
        columns = {str(column).strip() for column in dataframe.columns}
        return {"response_id", "response_name", "response_answer"}.issubset(columns)

    @staticmethod
    def looks_like_reference_class_report(dataframe, endpoint_name="", dataset_code=""):
        columns = {str(column).strip() for column in dataframe.columns}
        endpoint_token = str(endpoint_name or "").strip().lower()
        dataset_token = str(dataset_code or "").strip().lower()
        explicitly_reference = "reference_class_report" in endpoint_token or dataset_token == "referenceclassreport"
        has_answer_values = False
        if "response_answer" in dataframe.columns:
            answers = dataframe["response_answer"].fillna("").astype(str).str.strip()
            has_answer_values = answers.ne("").any()
        return (
            {"response_id", "response_name", "response_type"}.issubset(columns)
            and not has_answer_values
            and (explicitly_reference or "response_answer" not in columns)
        )

    @staticmethod
    def clean_label(value):
        raw_label = re.sub(r"\s+", " ", str(value or "")).strip(" :-")
        raw_label = re.sub(r"\bPilih\s+\d+\s+Bintang\s+untuk\s+mengisi\b", "", raw_label, flags=re.IGNORECASE).strip(" :-")
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
            response_id = cls.clean_scalar(row.get("response_id"))
            if not response_id:
                continue
            current = lookup.setdefault(
                response_id,
                {
                    "label": cls.clean_label(row.get("response_name")),
                    "type": cls.clean_scalar(row.get("response_type")),
                    "parent_id": cls.clean_scalar(row.get("response_parent_id")),
                    "class_start_dates": [],
                    "class_end_dates": [],
                },
            )
            if not current.get("label"):
                current["label"] = cls.clean_label(row.get("response_name"))
            for field_name, target_name in (
                ("class_start_date", "class_start_dates"),
                ("class_end_date", "class_end_dates"),
            ):
                value = cls.clean_scalar(row.get(field_name))
                if value and value not in current[target_name]:
                    current[target_name].append(value)
            current.update({
                "label": cls.clean_label(row.get("response_name")),
                "type": cls.clean_scalar(row.get("response_type")),
                "parent_id": cls.clean_scalar(row.get("response_parent_id")),
            })
        return lookup

    @staticmethod
    def clean_date_value(value):
        parsed = pd.to_datetime(str(value or "").strip(), errors="coerce")
        if pd.isna(parsed):
            return ""
        return parsed.strftime("%Y-%m-%d")

    @classmethod
    def collect_date_context(cls, row):
        start_dates = []
        end_dates = []
        for field_name in cls.START_DATE_FIELDS:
            value = cls.clean_date_value(row.get(field_name))
            if value and value not in start_dates:
                start_dates.append(value)
        for field_name in cls.END_DATE_FIELDS:
            value = cls.clean_date_value(row.get(field_name))
            if value and value not in end_dates:
                end_dates.append(value)
        return {"class_start_dates": start_dates, "class_end_dates": end_dates}

    @staticmethod
    def merge_date_context(target, source):
        for field_name in ("class_start_dates", "class_end_dates"):
            target_values = target.setdefault(field_name, [])
            for value in source.get(field_name, []):
                if value and value not in target_values:
                    target_values.append(value)

    @classmethod
    def date_context_key(cls, row):
        context = cls.collect_date_context(row)
        start_dates = tuple(sorted(context.get("class_start_dates", [])))
        end_dates = tuple(sorted(context.get("class_end_dates", [])))
        if not start_dates and not end_dates:
            return "__timeless__", context
        return (start_dates, end_dates), context

    @staticmethod
    def date_context(reference=None):
        reference = reference or {}
        start_dates = sorted(
            date_value
            for item in reference.get("class_start_dates", [])
            for date_value in [ClassReportAdapter.clean_date_value(item)]
            if date_value
        )
        end_dates = sorted(
            date_value
            for item in reference.get("class_end_dates", [])
            for date_value in [ClassReportAdapter.clean_date_value(item)]
            if date_value
        )
        if not start_dates and not end_dates:
            return APIDOG_UNKNOWN_DATE, APIDOG_TIMELESS_TIMEFRAME
        start_date = start_dates[0] if start_dates else end_dates[0]
        end_date = end_dates[-1] if end_dates else start_dates[-1]
        if start_date == end_date:
            return end_date, end_date
        return end_date, f"{start_date} sampai {end_date}"

    @staticmethod
    def semantics(clean_label):
        lowered = str(clean_label or "").lower()
        for keywords, service_label, journey_hint in CLASS_REPORT_JOURNEY_RULES:
            if any(keyword in lowered for keyword in keywords):
                return service_label, journey_hint
        if ClassReportAdapter.looks_like_raw_prompt(clean_label):
            return "Evaluasi umum kelas", "Pelaksanaan Layanan"
        return clean_label or "Evaluasi kelas", "Pelaksanaan Layanan"

    @staticmethod
    def looks_like_raw_prompt(value):
        text = re.sub(r"\s+", " ", str(value or "")).strip()
        if len(text) > 72 and any(mark in text for mark in ("?", "(", ")")):
            return True
        return any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in ClassReportAdapter.RAW_PROMPT_PATTERNS)

    @staticmethod
    def is_rating_response(row):
        response_type = ClassReportAdapter.clean_scalar(row.get("response_type")).lower()
        if response_type.startswith("rating"):
            return True
        answer = ClassReportAdapter.clean_scalar(row.get("response_answer"))
        return bool(answer) and pd.notna(pd.to_numeric(answer, errors="coerce"))

    @staticmethod
    def is_text_response(row):
        response_type = ClassReportAdapter.clean_scalar(row.get("response_type")).lower()
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
        response_id = cls.clean_scalar(row.get("response_id"))
        reference = reference_lookup.get(response_id, {})
        return cls.clean_label(reference.get("label") or row.get("response_name"))

    @classmethod
    def parent_id_for_row(cls, row, reference_lookup=None):
        explicit_parent_id = cls.clean_scalar(row.get("response_parent_id"))
        if explicit_parent_id:
            return explicit_parent_id
        response_id = cls.clean_scalar(row.get("response_id"))
        reference = (reference_lookup or {}).get(response_id, {})
        return cls.clean_scalar(reference.get("parent_id"))

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
    def build_row(cls, endpoint_name, row_index, question, rating_value, explanation_texts, date_context=None):
        service_label, journey_hint = cls.semantics(question)
        if pd.notna(rating_value):
            rating_prefixes = (
                "Skor ringkas",
                "Bacaan nilai",
                "Sinyal rating",
                "Ukuran pengalaman",
                "Catatan skor",
            )
            prefix = rating_prefixes[row_index % len(rating_prefixes)]
            average_text = f"{prefix} untuk {question}: {round(float(rating_value), 2)} dari 5"
        else:
            average_text = question
        explanations = cls.dedupe_texts(explanation_texts, limit=5)
        why_text = f"Mengapa: {'; '.join(explanations)}" if explanations else "Mengapa: belum ada komentar teks yang terhubung ke rating ini."
        feedback_date, timeframe = date_context or (APIDOG_UNKNOWN_DATE, APIDOG_TIMELESS_TIMEFRAME)
        return {
            "Record ID": f"{endpoint_name}-{row_index + 1:05d}",
            "Sumber Feedback": endpoint_name,
            "Kanal Feedback": APIDOG_CLASS_REPORT_CHANNEL,
            "Tanggal Feedback": feedback_date,
            "Tipe Stakeholder": "Peserta Kelas",
            "Layanan": service_label,
            "Lokasi": "",
            "Tipe Instruktur": "",
            "Rentang Waktu": timeframe,
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
            response_id = cls.clean_scalar(row.get("response_id"))
            answer = cls.clean_scalar(row.get("response_answer"))
            if not response_id and not answer:
                continue
            date_key, row_date_context = cls.date_context_key(row)
            if cls.is_rating_response(row):
                rating = pd.to_numeric(answer, errors="coerce")
                if pd.isna(rating):
                    continue
                group_key = (response_id, date_key)
                group = rating_groups.setdefault(
                    group_key,
                    {
                        "first_index": index,
                        "question": cls.question_label(row, reference_lookup),
                        "response_id": response_id,
                        "ratings": [],
                        "class_start_dates": [],
                        "class_end_dates": [],
                    },
                )
                group["ratings"].append(float(rating))
                cls.merge_date_context(group, row_date_context)
                continue
            if cls.is_text_response(row):
                parent_id = cls.parent_id_for_row(row, reference_lookup)
                if parent_id:
                    text_by_parent.setdefault((parent_id, date_key), []).append(answer)
                    group = rating_groups.get((parent_id, date_key))
                    if group is not None:
                        cls.merge_date_context(group, row_date_context)
                else:
                    orphan_text_rows.append((index, cls.question_label(row, reference_lookup), answer, row_date_context))

        rows = []
        for group_key, group in sorted(rating_groups.items(), key=lambda item: item[1]["first_index"]):
            response_id = group["response_id"]
            _, date_key = group_key
            ratings = group["ratings"]
            average_rating = sum(ratings) / len(ratings) if ratings else float("nan")
            explanations = text_by_parent.get((response_id, date_key), [])
            distribution = {}
            for value in ratings:
                label = cls.format_rating_value(value)
                distribution[label] = distribution.get(label, 0) + 1
            reference = reference_lookup.get(response_id, {})
            merged_reference = {**reference}
            cls.merge_date_context(merged_reference, group)
            row_payload = cls.build_row(
                endpoint_name,
                len(rows),
                group["question"],
                average_rating,
                explanations,
                date_context=cls.date_context(merged_reference),
            )
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

        for _, question, answer, row_date_context in orphan_text_rows:
            service_label, journey_hint = cls.semantics(question)
            feedback_date, timeframe = cls.date_context(row_date_context)
            rows.append(
                {
                    "Record ID": f"{endpoint_name}-{len(rows) + 1:05d}",
                    "Sumber Feedback": endpoint_name,
                    "Kanal Feedback": APIDOG_CLASS_REPORT_CHANNEL,
                    "Tanggal Feedback": feedback_date,
                    "Tipe Stakeholder": "Peserta Kelas",
                    "Layanan": service_label,
                    "Lokasi": "",
                    "Tipe Instruktur": "",
                    "Rentang Waktu": timeframe,
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

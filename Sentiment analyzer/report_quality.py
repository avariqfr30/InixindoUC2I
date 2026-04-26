import re

class ReportQualityValidator:
    REQUIRED_CHAPTER_IDS = {"cx_chap_1": "Descriptive chapter tersedia", "cx_chap_2": "Diagnostic chapter tersedia", "cx_chap_3": "Predictive chapter tersedia", "cx_chap_4": "Prescriptive chapter tersedia", "cx_chap_5": "Implementation readiness chapter tersedia"}

    @staticmethod
    def _plain_text(value):
        text = str(value or "")
        text = re.sub(r"\[\[(?:CHART|PIE|FLOW):.*?\]\]", " ", text)
        text = re.sub(r"[#*`>|_]", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def _section_map(report_sections):
        return {section.get("id"): section.get("content", "") for section in report_sections}

    @staticmethod
    def _check(checks, label, passed):
        checks.append({"label": label, "passed": bool(passed)})

    @classmethod
    def evaluate(cls, document, executive_snapshot, report_sections, score_label):
        checks = []
        section_map = cls._section_map(report_sections)
        plain_combined = cls._plain_text("\n".join([executive_snapshot or "", "\n".join(section_map.values())])).lower()

        cls._check(checks, "Executive snapshot substantif", len(cls._plain_text(executive_snapshot)) >= 400)
        for section_id, label in cls.REQUIRED_CHAPTER_IDS.items(): cls._check(checks, label, len(cls._plain_text(section_map.get(section_id, ""))) >= 250)

        cls._check(checks, "Score engine POV tercermin", score_label.lower() in plain_combined)
        cls._check(checks, "Customer journey teridentifikasi", "customer journey" in plain_combined or "tahap customer journey" in plain_combined)
        cls._check(checks, "Lokasi pelatihan tercantum", "lokasi pelatihan" in plain_combined or " lokasi " in f" {plain_combined} ")
        cls._check(checks, "Tipe instruktur tercantum", "tipe instruktur" in plain_combined or "instruktur" in plain_combined)
        cls._check(checks, "Prediksi menggunakan bahasa manusia", bool(re.search(r"diproyeksikan (turun|naik|relatif stabil)", plain_combined)))
        cls._check(checks, "Prediksi menyebut horizon waktu", bool(re.search(r"(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)\s+\d{4}|pada tahun \d{4}|1-2 bulan ke depan|1-2 minggu ke depan|semester berikutnya", plain_combined)))

        nonempty_paragraphs = sum(1 for paragraph in document.paragraphs if paragraph.text.strip())
        table_count, visual_count = len(document.tables), len(document.inline_shapes)

        cls._check(checks, "Dokumen memiliki paragraf yang memadai", nonempty_paragraphs >= 80)
        cls._check(checks, "Dokumen memiliki tabel pendukung", table_count >= 8)
        cls._check(checks, "Dokumen memiliki visual pendukung", visual_count >= 3)

        passed_checks = sum(1 for check in checks if check["passed"])
        total_checks = len(checks)
        completeness_score = round((passed_checks / total_checks) * 100, 1) if total_checks else 0.0
        verified_complete = completeness_score >= 80.0

        return {
            "verification_status": "verified" if verified_complete else "needs_review", "verified_complete": verified_complete,
            "completeness_score": completeness_score, "passed_checks": passed_checks, "total_checks": total_checks,
            "missing_checks": [check["label"] for check in checks if not check["passed"]],
            "document_stats": {"paragraph_count": nonempty_paragraphs, "table_count": table_count, "visual_count": visual_count},
            "summary": f"{passed_checks}/{total_checks} checks passed. Completeness score {completeness_score}%.",
        }


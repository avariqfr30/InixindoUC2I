import re
from reasoning_policy import FeedbackHotsReasoningPolicy
from editorial_intelligence import evaluate_feedback_document_spine

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

    @staticmethod
    def _has_raw_source_label(text):
        return bool(
            re.search(
                r"\b(APIDog|Internal API|endpoint|source\s*=|Evidence Ledger)\b|/api/Resource/dataset",
                str(text or ""),
                flags=re.IGNORECASE,
            )
        )

    @classmethod
    def _missing_action_contract_fields(cls, report_sections):
        section_map = cls._section_map(report_sections or [])
        prescriptive = cls._plain_text(section_map.get("cx_chap_4", "")).lower()
        required = {
            "fokus masalah": (r"\bfokus\b", r"\bisu\b"),
            "tindakan": (r"\btindakan\b", r"\baksi\b", r"\bintervensi\b"),
            "penanggung jawab": (r"\bowner\b", r"penanggung jawab"),
            "batas waktu": (r"batas waktu", r"\b\d+\s*hari\b", r"\bminggu\s+\d+"),
            "indikator keberhasilan": (r"indikator keberhasilan", r"kriteria penerimaan", r"\btarget\b"),
            "dampak yang diharapkan": (r"dampak yang diharapkan", r"hasil yang diharapkan"),
        }
        return [
            label
            for label, patterns in required.items()
            if not any(re.search(pattern, prescriptive, flags=re.IGNORECASE) for pattern in patterns)
        ]

    @classmethod
    def evaluate_narrative(
        cls,
        executive_snapshot,
        report_sections,
        deliberation_contract=None,
        appendix_content="",
    ):
        categories = set()
        findings = []
        if len(cls._plain_text(executive_snapshot)) < 80:
            categories.add("thin_executive_summary")
            findings.append("Ringkasan eksekutif terlalu tipis untuk dirender.")
        combined_text = "\n".join([executive_snapshot or "", *[section.get("content", "") for section in report_sections or []]])
        if FeedbackHotsReasoningPolicy.find_visible_reasoning(combined_text):
            categories.add("visible_reasoning")
            findings.append("Narasi masih memuat proses penalaran internal.")
        if FeedbackHotsReasoningPolicy.has_uncalibrated_feedback_claim(combined_text):
            categories.add("uncalibrated_feedback_claim")
            findings.append("Narasi menyatakan akar masalah terlalu pasti tanpa bukti, countercheck, atau batasan.")
        spine_result = evaluate_feedback_document_spine(executive_snapshot, report_sections)
        if not spine_result["passes"]:
            categories.update(spine_result["categories"])
            findings.extend(spine_result["findings"])
        missing_action_fields = cls._missing_action_contract_fields(report_sections)
        if missing_action_fields:
            categories.add("missing_action_contract")
            findings.append(
                "Rekomendasi belum lengkap dari sisi " + ", ".join(missing_action_fields) + "."
            )
        for section in report_sections or []:
            title = section.get("title") or section.get("id") or "section"
            content = section.get("content", "")
            if len(cls._plain_text(content)) < 80:
                categories.add("empty_section")
                findings.append(f"{title} kosong atau terlalu tipis.")
            if cls._has_raw_source_label(content):
                categories.add("raw_source_label")
                findings.append(f"{title} masih memuat label sumber mentah.")
        contract = deliberation_contract or {}
        if deliberation_contract is not None:
            required = {
                "evidence_dossier", "research_plan", "document_thesis", "chapter_contracts",
                "claim_ledger", "data_gap_register", "editorial_contract", "appendix_manifest",
            }
            if not required.issubset(contract):
                categories.add("missing_deliberation_contract")
                findings.append("Kontrak deliberasi laporan feedback belum lengkap.")
            appendix = str(appendix_content or "")
            if (
                "## A. Cakupan dan Metodologi" not in appendix
                or "## B. Matriks Temuan dan Pengukuran" not in appendix
                or "## C. Kesenjangan Data" not in appendix
            ):
                categories.add("missing_tiered_appendix")
                findings.append("Lampiran metodologi, pengukuran, dan kesenjangan data belum lengkap.")
            if re.search(r"\b(?:ClassReport|ReferenceClassReport|SECTION_PLAN_JSON|DOCUMENT_DELIBERATION)\b", appendix):
                categories.add("raw_source_label")
                findings.append("Lampiran masih memuat label sumber atau perencanaan internal.")
        return {"passes": not categories, "categories": sorted(categories), "findings": findings}

    @classmethod
    def evaluate(cls, document, executive_snapshot, report_sections, score_label):
        checks = []
        section_map = cls._section_map(report_sections)
        plain_combined = cls._plain_text("\n".join([executive_snapshot or "", "\n".join(section_map.values())])).lower()

        cls._check(checks, "Executive snapshot substantif", len(cls._plain_text(executive_snapshot)) >= 400)
        for section_id, label in cls.REQUIRED_CHAPTER_IDS.items(): cls._check(checks, label, len(cls._plain_text(section_map.get(section_id, ""))) >= 250)

        score_terms = {score_label.lower()}
        if score_label.lower() == "experience index":
            score_terms.add("indeks pengalaman")
        cls._check(checks, "Perspektif skor tercermin", any(term in plain_combined for term in score_terms))
        cls._check(
            checks,
            "Perjalanan pelanggan teridentifikasi",
            "customer journey" in plain_combined
            or "tahap customer journey" in plain_combined
            or "perjalanan pelanggan" in plain_combined
            or "tahap perjalanan pelanggan" in plain_combined,
        )
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

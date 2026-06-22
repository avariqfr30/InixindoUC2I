"""Deterministic factuality packet for feedback-report generation."""
from __future__ import annotations

import hashlib
import json
import re
from collections import Counter
from typing import Any


class ReportFactRegistry:
    VERSION = "feedback-fact-registry-v1"
    SNAPSHOT_COLUMNS = (
        "Record ID",
        "Tanggal Feedback",
        "Tipe Stakeholder",
        "Layanan",
        "Lokasi",
        "Tipe Instruktur",
        "Rentang Waktu",
        "Rating Numeric",
        "Komentar",
        "Sentiment Label",
        "Customer Journey Hint",
        "Raw Response Count",
        "Rating Response Count",
        "Text Response Count",
    )
    THEME_FACT_TOKENS = {
        "responsiveness": "RESPONS",
        "schedule": "JADWAL",
        "facility": "FASILITAS",
        "instructor": "INSTRUKTUR",
        "material": "MATERI",
        "communication": "KOMUNIKASI",
        "outcome": "HASIL",
    }

    @staticmethod
    def _clean_number(value: Any) -> int | float:
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return 0
        return int(numeric) if numeric.is_integer() else round(numeric, 2)

    @classmethod
    def _snapshot_fingerprint(cls, dataframe: Any, scope: dict[str, Any]) -> str:
        columns = [column for column in cls.SNAPSHOT_COLUMNS if column in getattr(dataframe, "columns", [])]
        rows = []
        if dataframe is not None and not getattr(dataframe, "empty", True):
            for _, row in dataframe[columns].fillna("").iterrows():
                rows.append([str(row.get(column, "")).strip() for column in columns])
        payload = {
            "columns": columns,
            "rows": sorted(rows),
            "scope": {key: str(scope.get(key, "")) for key in sorted(scope)},
            "version": cls.VERSION,
        }
        raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return "sha256:" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]

    @staticmethod
    def _fact(fact_id: str, metric: str, value: Any, unit: str, basis: str) -> dict[str, Any]:
        return {
            "fact_id": fact_id,
            "metric": metric,
            "value": value,
            "unit": unit,
            "basis": basis,
        }

    @classmethod
    def theme_fact_id(cls, theme_id: Any) -> str:
        normalized = re.sub(r"[^a-z0-9]+", "-", str(theme_id or "").lower()).strip("-")
        token = cls.THEME_FACT_TOKENS.get(normalized)
        if not token:
            token = hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:8].upper()
        return "F-TEMA-" + token

    @classmethod
    def build(
        cls,
        dataframe: Any,
        analysis_context: dict[str, Any],
        governance: dict[str, Any],
        contradiction_review: dict[str, Any],
        scope: dict[str, Any],
    ) -> dict[str, Any]:
        response_count = int(governance.get("total_rows", 0) or 0)
        rating_count = int(governance.get("rating_response_count", 0) or 0)
        text_count = int(governance.get("text_response_count", 0) or 0)
        completeness = cls._clean_number(governance.get("completeness_pct", 0))
        score_metrics = analysis_context.get("score_metrics") or {}
        score_label = str((analysis_context.get("score_profile") or {}).get("label") or "Skor utama")

        facts = [
            cls._fact("F-RESPONSES", "Respons terolah", response_count, "respons", "filter aktif"),
            cls._fact("F-RATINGS", "Respons dengan rating", rating_count, "respons", "filter aktif"),
            cls._fact("F-COMMENTS", "Respons dengan komentar", text_count, "respons", "filter aktif"),
            cls._fact("F-COMPLETENESS", "Kelengkapan field inti", completeness, "%", "field tata kelola"),
            cls._fact(
                "F-SCORE-CURRENT",
                f"{score_label} saat ini",
                cls._clean_number(score_metrics.get("current_score", 0)),
                "poin",
                "score engine aktif",
            ),
            cls._fact(
                "F-SCORE-PROJECTED",
                f"{score_label} proyeksi",
                cls._clean_number(score_metrics.get("projected_score", 0)),
                "poin",
                "early-warning deterministik",
            ),
        ]

        theme_evidence_ids = {}
        for theme in score_metrics.get("theme_rows") or []:
            theme_id = re.sub(r"[^a-z0-9]+", "-", str(theme.get("theme_id") or "").lower()).strip("-")
            if not theme_id:
                continue
            fact_id = cls.theme_fact_id(theme_id)
            theme_evidence_ids[theme_id] = fact_id
            facts.append(
                cls._fact(
                    fact_id,
                    str(theme.get("label") or theme_id),
                    int(theme.get("negative_hits", 0) or 0),
                    "sinyal korektif",
                    f"{int(theme.get('total_hits', 0) or 0)} respons bertema",
                )
            )

        rating_coverage = round((rating_count / response_count) * 100, 1) if response_count else 0.0
        text_coverage = round((text_count / response_count) * 100, 1) if response_count else 0.0
        limitations = []
        if str(scope.get("segment") or "all").lower() == "all":
            limitations.append("Perbandingan antarsegmen belum diuji.")
        if not scope.get("external_context_ready"):
            limitations.append("Pembanding eksternal belum cukup kuat.")
        if text_count < 5:
            limitations.append("Bukti komentar masih terbatas.")
        if float(completeness or 0) < 70:
            limitations.append("Kelengkapan field inti membatasi interpretasi.")

        if response_count >= 100 and rating_coverage >= 80 and text_count >= 20 and float(completeness or 0) >= 80:
            confidence = "Tinggi"
        elif response_count >= 20 and rating_coverage >= 60 and float(completeness or 0) >= 60:
            confidence = "Sedang"
        else:
            confidence = "Rendah"

        return {
            "version": cls.VERSION,
            "snapshot_fingerprint": cls._snapshot_fingerprint(dataframe, scope),
            "scope": {key: scope.get(key) for key in sorted(scope)},
            "facts": facts,
            "theme_evidence_ids": theme_evidence_ids,
            "confidence_basis": {
                "level": confidence,
                "response_count": response_count,
                "rating_coverage_pct": rating_coverage,
                "comment_coverage_pct": text_coverage,
                "field_completeness_pct": completeness,
                "limitations": limitations,
            },
            "contradiction_review": dict(contradiction_review or {}),
        }


class NarrativeFactValidator:
    NUMERIC_TOKEN = re.compile(
        r"(?<![\w])(?:Rp\s*)?\d+(?:[.,]\d+)*(?:\s?%|/5)?(?![\w])",
        flags=re.IGNORECASE,
    )

    @classmethod
    def numeric_tokens(cls, text: Any) -> Counter[str]:
        return Counter(
            re.sub(r"\s+", "", match.group(0)).lower()
            for match in cls.NUMERIC_TOKEN.finditer(str(text or ""))
        )

    @classmethod
    def preserves_numeric_facts(cls, original: Any, candidate: Any) -> bool:
        return cls.numeric_tokens(original) == cls.numeric_tokens(candidate)

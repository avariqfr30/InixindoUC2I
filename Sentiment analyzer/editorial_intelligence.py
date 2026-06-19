"""Feedback-specific evidence, voice, and editorial quality helpers."""
from __future__ import annotations

import re
from collections import Counter
from typing import Any


EXCLUDED_DATASETS = {"FinanceInvoice", "ProjectStandards"}


def compact_sentence(value: Any, max_words: int = 16) -> str:
    words = re.sub(r"\s+", " ", str(value or "").strip()).split()
    if len(words) <= max_words:
        return " ".join(words)
    return " ".join(words[:max_words]).rstrip(".,;:") + "."


def feedback_voice_rules() -> list[str]:
    return [
        "Tulis seperti laporan perbaikan layanan yang menghargai suara peserta, bukan laporan template.",
        "Bedakan fakta pelanggan, pembacaan manajemen, dan tindak lanjut.",
        "Gunakan bahasa manusiawi saat menjelaskan keluhan; jangan menyalahkan peserta atau instruktur.",
        "Tabel hanya untuk perbandingan cepat; penjelasan empatik diletakkan setelah tabel.",
        "Gunakan OSINT sebagai benchmark ringan; suara peserta dari ClassReport tetap menjadi sumber utama.",
    ]


def build_issue_story_action(signal: Any, meaning: Any, action: Any) -> dict[str, str]:
    return {
        "participant_signal": compact_sentence(signal, 24),
        "service_meaning": compact_sentence(meaning, 24),
        "next_action": compact_sentence(action, 24),
    }


def compact_feedback_table_rows(rows: list[list[Any]] | None, max_cell_words: int = 14) -> list[list[str]]:
    output: list[list[str]] = []
    repeated: Counter[str] = Counter()
    for row in rows or []:
        next_row: list[str] = []
        for cell in row:
            text = compact_sentence(cell, max_cell_words)
            signature = text.lower()
            repeated[signature] += 1
            if repeated[signature] > 2 and len(text.split()) >= 4:
                text = ""
            next_row.append(text)
        output.append(next_row)
    return output


def assess_feedback_style(text: Any) -> dict[str, Any]:
    paragraphs = [part.strip().lower() for part in re.split(r"\n\s*\n|(?<=[.!?])\s+", str(text or "")) if part.strip()]
    metric_led = sum(part.startswith(("experience index", "sentiment", "rating")) for part in paragraphs)
    findings = ["metric_led_repetition"] if metric_led >= 3 else []
    return {"passed": not findings, "findings": findings}


def class_report_roles() -> dict[str, str]:
    return {
        "ClassReport": "suara peserta dan skor evaluasi kelas",
        "ReferenceClassReport": "kamus pertanyaan dan konteks kelas untuk membaca ClassReport",
    }

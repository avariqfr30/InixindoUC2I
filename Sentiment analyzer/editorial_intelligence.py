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


FEEDBACK_SPINE_ORDER = (
    "Descriptive", "Diagnostic", "Predictive", "Prescriptive", "Implementation",
)
FEEDBACK_CONNECTOR_TERMS = (
    "melanjutkan", "menjadi dasar", "berangkat dari", "menghubungkan",
    "arah berikutnya", "diterjemahkan menjadi", "dibaca sebagai kelanjutan",
    "dari baseline", "rangkaian layanan", "kaitan dengan", "dari sini",
    "pembahasan kemudian", "dasar tersebut",
)


def _plain_document_text(value: Any) -> str:
    text = re.sub(r"\[\[(?:CHART|PIE|FLOW|DASHBOARD):.*?\]\]", " ", str(value or ""))
    text = re.sub(r"[#*`>|_]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _section_title(section: dict[str, Any]) -> str:
    return str(section.get("title") or section.get("id") or "bagian").strip()


def _has_connector(text: Any, *titles: Any) -> bool:
    lowered = _plain_document_text(text).lower()
    if any(term in lowered for term in FEEDBACK_CONNECTOR_TERMS):
        return True
    return any(str(title or "").strip().lower() in lowered for title in titles if str(title or "").strip())


def _openings(text: Any, width: int = 3) -> Counter[str]:
    paragraphs = [
        part.strip()
        for index, part in enumerate(_split_feedback_blocks(text))
        if index % 2 == 0 and part.strip() and not part.lstrip().startswith(("|", "[["))
    ]
    signatures = []
    for part in paragraphs:
        words = re.findall(r"[a-z0-9]+", part.lower())[:width]
        if not words or all(word.isdigit() for word in words):
            continue
        signatures.append(" ".join(words))
    return Counter(signatures)


def _split_feedback_blocks(text: Any) -> list[str]:
    return re.split(
        r"(^#{1,6}[^\n]*(?:\n|$)|\n\s*\n)",
        str(text or ""),
        flags=re.MULTILINE,
    )



def _vary_repeated_openings_in_text(text: Any) -> str:
    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", str(text or ""))]
    variants = (
        "Dalam praktiknya, ",
        "Dari sisi layanan, ",
        "Bagi pembaca manajemen, ",
        "Pada titik ini, ",
        "Sebagai konsekuensi, ",
        "Untuk tindak lanjut, ",
    )
    seen: Counter[str] = Counter()
    output: list[str] = []
    for paragraph in paragraphs:
        if not paragraph:
            output.append(paragraph)
            continue
        words = re.findall(r"[a-z0-9]+", paragraph.lower())[:3]
        signature = " ".join(words)
        seen[signature] += 1
        if signature and seen[signature] > 1 and not paragraph.startswith(("|", "#", "[[")):
            prefix = variants[(seen[signature] - 2) % len(variants)]
            if paragraph.startswith("- "):
                paragraph = "- " + prefix + paragraph[2:].lstrip()
            else:
                paragraph = prefix + paragraph[0].lower() + paragraph[1:]
        output.append(paragraph)
    return "\n\n".join(output).strip()


def _vary_repeated_openings_across_sections(sections: list[dict[str, Any]], threshold: int = 3) -> None:
    combined = Counter()
    for section in sections:
        combined.update(_openings(section.get("content", "")))
    repeated = {opening for opening, count in combined.items() if opening and count >= threshold}
    if not repeated:
        return

    variants = (
        "Dalam praktiknya, ",
        "Dari sisi layanan, ",
        "Bagi pembaca manajemen, ",
        "Pada titik ini, ",
        "Sebagai konsekuensi, ",
        "Untuk tindak lanjut, ",
    )
    seen: Counter[str] = Counter()
    variant_index = 0
    for section in sections:
        paragraphs = _split_feedback_blocks(section.get("content", ""))
        for index in range(0, len(paragraphs), 2):
            paragraph = paragraphs[index].strip()
            if not paragraph or paragraph.startswith(("|", "[[")):
                continue
            signature = " ".join(re.findall(r"[a-z0-9]+", paragraph.lower())[:3])
            if signature not in repeated:
                continue
            seen[signature] += 1
            if seen[signature] == 1 or paragraph.startswith(("|", "#", "[[")):
                continue
            prefix = variants[variant_index % len(variants)]
            variant_index += 1
            leading = paragraphs[index][:len(paragraphs[index]) - len(paragraphs[index].lstrip())]
            varied = prefix + paragraph[0].lower() + paragraph[1:]
            if paragraph.startswith("- "):
                varied = "- " + prefix + paragraph[2:3].lower() + paragraph[3:]
            paragraphs[index] = leading + varied
        section["content"] = "".join(paragraphs).strip()


def evaluate_feedback_document_spine(executive_snapshot: Any, report_sections: list[dict[str, Any]]) -> dict[str, Any]:
    categories: set[str] = set()
    findings: list[str] = []
    sections = list(report_sections or [])
    combined = (str(executive_snapshot or "") + "\n" + "\n".join(_section_title(section) + "\n" + str(section.get("content") or "") for section in sections)).lower()
    for term in FEEDBACK_SPINE_ORDER:
        if term.lower() not in combined:
            categories.add("missing_service_logic_stage")
            findings.append(f"Alur laporan belum menyebut tahap {term} secara jelas.")
    combined_openings = Counter()
    combined_openings.update(_openings(executive_snapshot))
    for section in sections:
        combined_openings.update(_openings(section.get("content", "")))
    repeated_global = [opening for opening, count in combined_openings.items() if opening and count >= 3]
    if repeated_global:
        categories.add("global_repeated_opening_spine")
        findings.append("Laporan masih memakai pembuka paragraf berulang lintas bagian: " + ", ".join(repeated_global[:4]) + ".")

    for index, section in enumerate(sections):
        title = _section_title(section)
        content = section.get("content", "")
        if index > 0 and not _has_connector(str(content)[:900], _section_title(sections[index - 1]), title):
            categories.add("missing_previous_handoff")
            findings.append(f"{title} belum mengaitkan temuan dengan bagian sebelumnya.")
        if index < len(sections) - 1 and not _has_connector(str(content)[-900:], title, _section_title(sections[index + 1])):
            categories.add("missing_next_handoff")
            findings.append(f"{title} belum menyiapkan pembaca menuju bagian berikutnya.")
        if any(count >= 3 and opening for opening, count in _openings(content).items()):
            categories.add("repeated_opening_spine")
            findings.append(f"{title} masih memakai pembuka paragraf yang berulang.")
    return {"passes": not categories, "categories": sorted(categories), "findings": findings}


def repair_feedback_document_spine(executive_snapshot: Any, report_sections: list[dict[str, Any]]) -> tuple[str, list[dict[str, Any]]]:
    sections = [dict(section) for section in report_sections or []]
    stage_line = "Alur laporan dibaca sebagai rangkaian Descriptive, Diagnostic, Predictive, Prescriptive, lalu Implementation agar suara peserta berujung pada keputusan layanan yang bisa dijalankan."
    snapshot = str(executive_snapshot or "").strip()
    if not all(term.lower() in snapshot.lower() for term in FEEDBACK_SPINE_ORDER):
        snapshot = f"{snapshot}\n\n{stage_line}".strip()
    openers = (
        "Melanjutkan {previous}, {current} menjelaskan arti temuan sebelumnya bagi pengalaman peserta.",
        "Berangkat dari {previous}, {current} mempersempit sinyal pelanggan menjadi pembacaan manajemen yang lebih tajam.",
        "Setelah bagian {previous}, {current} menjaga alur agar rekomendasi tidak muncul tanpa dasar layanan.",
        "Dari baseline {previous}, {current} membaca konsekuensi yang perlu dipahami sebelum tindakan dipilih.",
        "Rangkaian layanan dari {previous} berlanjut ke {current} agar suara peserta tidak berhenti sebagai skor.",
        "Kaitan dengan {previous} membuat {current} berfungsi sebagai penajaman, bukan bab yang terpisah.",
    )
    closers = (
        "Temuan ini menjadi dasar untuk {next}, sehingga pembaca melihat perpindahan dari bukti ke keputusan berikutnya.",
        "Arah berikutnya masuk ke {next}, tempat konsekuensi bagian ini diterjemahkan menjadi prioritas layanan.",
        "Implikasi bagian ini dibaca lebih lanjut pada {next}, bukan sebagai topik baru yang terpisah.",
        "Dari sini, {next} mengambil alih pembahasan agar analisis berubah menjadi keputusan layanan.",
        "Pembahasan kemudian bergerak ke {next}, sehingga alur laporan tetap mengikuti perjalanan pelanggan.",
        "Dasar tersebut mengantar pembaca ke {next}, tempat prioritas perbaikan diuji dari sisi eksekusi.",
    )
    for index, section in enumerate(sections):
        title = _section_title(section)
        content = str(section.get("content") or "").strip()
        before: list[str] = []
        after: list[str] = []
        if index > 0:
            previous = _section_title(sections[index - 1])
            if not _has_connector(content[:900], previous, title):
                before.append(openers[(index - 1) % len(openers)].format(previous=previous, current=title))
        if index < len(sections) - 1:
            next_title = _section_title(sections[index + 1])
            if not _has_connector(content[-900:], title, next_title):
                after.append(closers[index % len(closers)].format(next=next_title))
        section["content"] = _vary_repeated_openings_in_text("\n\n".join([*before, content, *after]).strip())
    visible_blocks = [{"title": "Ringkasan Eksekutif", "content": snapshot}, *sections]
    _vary_repeated_openings_across_sections(visible_blocks)
    snapshot = visible_blocks[0]["content"]
    return snapshot, sections

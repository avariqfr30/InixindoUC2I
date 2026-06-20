"""Deterministic document-level deliberation for feedback reports."""
from __future__ import annotations

import copy
import hashlib
import json
import re
from typing import Any

from feedback_exemplar_profile import (
    UC2_FEEDBACK_EXEMPLAR_PROFILE_VERSION,
    build_uc2_feedback_exemplar_profile,
)


def _clean(value: Any, max_words: int = 34) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip(" -;,.:")
    words = text.split()
    if len(words) > max_words:
        text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
    return text


def _digest(value: Any) -> str:
    raw = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


class FeedbackDeliberationBuilder:
    CACHE_VERSION = "feedback-deliberation-v1"
    _cache: dict[str, dict[str, Any]] = {}
    _stats = {"hits": 0, "misses": 0}

    @classmethod
    def clear_cache(cls) -> None:
        cls._cache.clear()
        cls._stats = {"hits": 0, "misses": 0}

    @classmethod
    def cache_stats(cls) -> dict[str, int]:
        return {**cls._stats, "items": len(cls._cache)}

    @classmethod
    def _remember(cls, key: str, value: dict[str, Any]) -> dict[str, Any]:
        cls._cache[key] = copy.deepcopy(value)
        while len(cls._cache) > 128:
            cls._cache.pop(next(iter(cls._cache)))
        return copy.deepcopy(value)

    @classmethod
    def build(
        cls,
        sections: list[dict[str, Any]],
        context: dict[str, Any] | None = None,
        data_version: str = "",
    ) -> dict[str, Any]:
        context = dict(context or {})
        section_inputs = [
            {"id": item.get("id"), "title": item.get("title")}
            for item in sections or []
        ]
        insight_cards = [dict(item) for item in context.get("insight_cards", []) if _clean(item.get("observation"))]
        exemplar_profile = build_uc2_feedback_exemplar_profile()
        cache_key = _digest({
            "version": cls.CACHE_VERSION,
            "exemplar_profile_version": UC2_FEEDBACK_EXEMPLAR_PROFILE_VERSION,
            "sections": section_inputs,
            "context": context,
            "data_version": data_version,
        })
        if cache_key in cls._cache:
            cls._stats["hits"] += 1
            return copy.deepcopy(cls._cache[cache_key])
        cls._stats["misses"] += 1

        claim_ledger = []
        for index, card in enumerate(insight_cards, start=1):
            claim_ledger.append({
                "claim_id": f"F-{index:03d}",
                "finding": _clean(card.get("observation")),
                "implication": _clean(card.get("implication")),
                "confidence": str(card.get("confidence") or "medium").lower(),
                "measurement": "Bandingkan rating, rasio tema, dan komentar pada periode evaluasi berikutnya.",
            })

        section_ids = [str(item.get("id") or "").strip() for item in sections if item.get("id")]
        chapter_contracts = []
        for index, section in enumerate(sections or []):
            chapter_contracts.append({
                "section_id": section_ids[index],
                "title": _clean(section.get("title"), 16),
                "depends_on": section_ids[index - 1] if index else "",
                "hands_off_to": section_ids[index + 1] if index + 1 < len(section_ids) else "",
                "argument_contract": ["temuan", "bukti", "implikasi", "countercheck", "tindakan"],
                "measurement_obligation": "Nyatakan indikator, pemilik pengukuran, dan waktu tinjauan untuk rekomendasi.",
            })

        gaps = []
        if not context.get("external_context_ready"):
            gaps.append({
                "area": "Pembanding eksternal",
                "gap": "Konteks eksternal yang cukup sebanding belum tersedia.",
                "handling": "Keputusan tetap bertumpu pada feedback internal.",
            })
        if str(context.get("segment") or "all").lower() == "all":
            gaps.append({
                "area": "Perbandingan segmen",
                "gap": "Perbedaan pengalaman antarsegmen belum diuji pada filter ini.",
                "handling": "Jangan mengklaim satu segmen sebagai penyebab dominan.",
            })
        if int(context.get("text_response_count") or 0) == 0:
            gaps.append({
                "area": "Komentar peserta",
                "gap": "Komentar teks tidak tersedia pada cakupan aktif.",
                "handling": "Batasi diagnosis pada pola rating dan metadata yang tersedia.",
            })

        thesis = (
            "Laporan harus mengubah pola pengalaman peserta menjadi keputusan layanan yang terukur, "
            "tanpa mengubah sinyal korelasi menjadi kepastian sebab-akibat."
        )
        contract = {
            "cache_key": cache_key,
            "data_version": data_version,
            "evidence_dossier": {
                "snapshot_policy": "immutable_per_generation",
                "timeframe": _clean(context.get("timeframe_label") or context.get("timeframe"), 10),
                "segment": _clean(context.get("segment") or "Semua segmen", 8),
                "sentiment": _clean(context.get("sentiment") or "Semua sentimen", 8),
                "response_count": int(context.get("row_count") or 0),
                "text_response_count": int(context.get("text_response_count") or 0),
            },
            "research_plan": {
                "questions": [
                    {"question": "Apakah pola rating konsisten dengan komentar peserta?", "evidence_needed": "rating dan komentar", "countercheck": "cari perbedaan arah"},
                    {"question": "Apakah temuan berubah menurut segmen atau perjalanan peserta?", "evidence_needed": "segmen dan tahap layanan", "countercheck": "hindari dominasi semu dari volume"},
                    {"question": "Apakah konteks eksternal benar-benar sebanding?", "evidence_needed": "pembanding layanan yang relevan", "countercheck": "jangan menggantikan bukti internal"},
                ],
                "bounded": True,
            },
            "document_thesis": thesis,
            "chapter_contracts": chapter_contracts,
            "claim_ledger": claim_ledger,
            "data_gap_register": gaps,
            "editorial_contract": {
                "voice": "analis pengalaman pelanggan yang netral, jernih, dan bertanggung jawab",
                "rules": [
                    "Tulis langsung dalam Bahasa Indonesia yang alami dan mudah dipertanggungjawabkan.",
                    "Bedakan temuan, bukti, tafsir, countercheck, dan tindakan tanpa label template berulang.",
                    "Gunakan komentar sebagai contoh pola, bukan bukti tunggal sebab-akibat.",
                    "Hubungkan setiap bagian dengan kesimpulan bagian sebelumnya.",
                    "Gunakan profil contoh UC2 hanya sebagai kalibrasi gaya dan struktur analisis; jangan menyalin frasa atau menjadikannya bukti fakta.",
                ],
                "meaning_lock": ["rating", "jumlah respons", "periode", "segmen", "tema", "kutipan"],
                "forbidden": ["label agen", "nama dataset", "prompt", "chain-of-thought", "klaim sebab tanpa bukti"],
                "exemplar_profile": exemplar_profile,
            },
            "appendix_manifest": {
                "coverage": {
                    "timeframe": _clean(context.get("timeframe_label") or context.get("timeframe"), 10),
                    "response_count": int(context.get("row_count") or 0),
                    "text_response_count": int(context.get("text_response_count") or 0),
                    "segment": _clean(context.get("segment") or "Semua segmen", 8),
                    "sentiment": _clean(context.get("sentiment") or "Semua sentimen", 8),
                },
                "claims": claim_ledger,
                "data_gaps": gaps,
            },
        }
        return cls._remember(cache_key, contract)

    @staticmethod
    def for_section(contract: dict[str, Any], section_id: str) -> str:
        chapter = next(
            (item for item in contract.get("chapter_contracts", []) if item.get("section_id") == section_id),
            {},
        )
        payload = {
            "document_thesis": contract.get("document_thesis"),
            "section_contract": chapter,
            "claim_ledger": contract.get("claim_ledger"),
            "data_gaps": contract.get("data_gap_register"),
            "editorial_contract": contract.get("editorial_contract"),
        }
        return (
            "[DOCUMENT_DELIBERATION] Gunakan kontrak ini secara internal dan jangan tampilkan struktur atau proses berpikirnya. "
            + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        )

    @staticmethod
    def build_appendix_markdown(contract: dict[str, Any]) -> str:
        manifest = contract.get("appendix_manifest") or {}
        coverage = manifest.get("coverage") or {}
        lines = [
            "# Lampiran Metodologi, Pengukuran, dan Kesenjangan Data",
            "Lampiran ini menyimpan rincian pendukung agar narasi utama tetap mudah dibaca dan keputusan dapat ditelusuri.",
            "",
            "## A. Cakupan dan Metodologi",
            f"- Periode analisis: {coverage.get('timeframe') or 'periode terpilih'}.",
            f"- Respons terolah: {coverage.get('response_count') or 0}; respons dengan komentar teks: {coverage.get('text_response_count') or 0}.",
            f"- Filter segmen: {coverage.get('segment') or 'semua'}; filter sentimen: {coverage.get('sentiment') or 'semua'}.",
            "- Rating, komentar, tema, dan konteks layanan dibaca bersama; komentar tidak diperlakukan sebagai bukti tunggal sebab-akibat.",
            "",
            "## B. Matriks Temuan dan Pengukuran",
            "| ID | Temuan | Implikasi | Cara Mengukur Perubahan | Tingkat Keyakinan |",
            "| --- | --- | --- | --- | --- |",
        ]
        for item in manifest.get("claims", []):
            confidence = {"high": "Tinggi", "medium": "Sedang", "low": "Terbatas"}.get(item.get("confidence"), "Sedang")
            values = [item.get("claim_id"), item.get("finding"), item.get("implication"), item.get("measurement"), confidence]
            lines.append("| " + " | ".join(str(value or "-").replace("|", "/") for value in values) + " |")
        if not manifest.get("claims"):
            lines.append("| - | Bukti temuan belum memadai. | - | Kumpulkan respons tambahan. | Terbatas |")
        lines.extend(["", "## C. Kesenjangan Data"])
        gaps = manifest.get("data_gaps", [])
        if gaps:
            lines.extend(f"- **{item.get('area')}:** {item.get('gap')} {item.get('handling')}" for item in gaps)
        else:
            lines.append("- Tidak ada kesenjangan data material yang teridentifikasi pada cakupan aktif.")
        return "\n".join(lines).strip()

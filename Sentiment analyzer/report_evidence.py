import re

from document_builder import DocumentBuilder
from editorial_intelligence import class_report_roles, feedback_voice_rules

RAW_SOURCE_PATTERNS = [
    r"\bAPIDog\b",
    r"\bInternal API\b",
    r"\bendpoint\b",
    r"\bsource\s*=",
    r"/api/Resource/dataset",
    r"\bEvidence Ledger\b",
]

# ── PII scrubbing ──
_PII_PATTERNS = [
    # Indonesian mobile numbers (+62xxx, 08xxx)
    (re.compile(r"\+?62[\s\-]?\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}"), "[nomor telepon dirahasiakan]"),
    (re.compile(r"0\d{2,3}[\s\-]?\d{3,4}[\s\-]?\d{3,4}"), "[nomor telepon dirahasiakan]"),
    # Email addresses
    (re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"), "[email dirahasiakan]"),
    # Indonesian NIK (16 consecutive digits)
    (re.compile(r"\b\d{16}\b"), "[NIK dirahasiakan]"),
    # Names preceded by common Indonesian honorifics
    (re.compile(r"(?:Bapak|Ibu|Pak|Bu|Mas|Mbak|Bp\.?|Ibu\.?)\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2}"), "[nama dirahasiakan]"),
]


def _scrub_pii(text: str) -> str:
    """Remove personally identifiable information from text."""
    for pattern, replacement in _PII_PATTERNS:
        text = pattern.sub(replacement, text)
    return text


class ContextIntelligenceDesk:
    """Build hidden, reader-safe report context from UI notes and normalized feedback rows."""

    @staticmethod
    def _safe_text(value, max_words=42):
        text = DocumentBuilder.reader_facing_text(str(value or ""))
        text = _scrub_pii(text)
        text = re.sub(r"https?://\S+", "", text)
        text = re.sub(r"/api/\S+", "", text)
        text = re.sub(r"\b(?:APIDog|Internal API|endpoint|source\s*=|Evidence Ledger)\b", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\bProblem,\s*Opportunity,\s*Directive\b", "prioritas masalah, peluang, dan arahan keputusan", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+", " ", text).strip(" -;,.")
        words = text.split()
        if len(words) > max_words:
            text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
        return text

    @classmethod
    def _feedback_insight_cards(cls, dataframe=None, macro_trends=""):
        cards = []
        if dataframe is not None and not getattr(dataframe, "empty", True):
            working = dataframe
            if "Reportable Analysis Row" in working.columns:
                working = working[working["Reportable Analysis Row"]].copy()
            row_count = len(working)
            if row_count:
                cards.append({
                    "observation": f"{row_count} respons evaluasi terolah pada cakupan aktif.",
                    "implication": "Temuan perlu diperlakukan sebagai sinyal pengalaman belajar, bukan opini tunggal.",
                    "recommended_angle": "Tulis dari pola pengalaman peserta menuju keputusan layanan yang bisa diambil manajemen.",
                    "confidence": "high",
                })
            if "Layanan" in working.columns:
                counts = working["Layanan"].astype(str).replace({"nan": ""}).value_counts()
                if not counts.empty and counts.index[0].strip():
                    service = cls._safe_text(counts.index[0], max_words=10)
                    cards.append({
                        "observation": f"Layanan paling sering muncul dalam cakupan ini adalah {service}.",
                        "implication": "Volume tinggi menunjukkan area eksposur terbesar; belum otomatis berarti performa paling buruk.",
                        "recommended_angle": "Bedakan eksposur layanan, bukti keluhan, dan prioritas perbaikan.",
                        "confidence": "medium",
                    })
            if "Komentar" in working.columns:
                comments = working["Komentar"].astype(str)
                text_count = int(comments.str.strip().replace({"nan": ""}).astype(bool).sum())
                if text_count:
                    cards.append({
                        "observation": f"{text_count} respons memiliki komentar teks yang dapat dibaca sebagai bukti kualitatif.",
                        "implication": "Komentar perlu dipakai untuk menjelaskan sebab kemungkinan, bukan menggantikan distribusi rating.",
                        "recommended_angle": "Gunakan komentar sebagai contoh pola, lalu ikat kembali ke angka dan segmen.",
                        "confidence": "medium",
                    })
        if cls._external_context_ready(macro_trends):
            cards.append({
                "observation": "Konteks eksternal tersedia sebagai pembanding cara membaca kualitas layanan.",
                "implication": "OSINT dapat memperkaya tafsir, tetapi tidak boleh mengalahkan bukti feedback internal.",
                "recommended_angle": "Pakai OSINT untuk memperjelas mengapa isu layanan penting bagi kepercayaan peserta.",
                "confidence": "medium",
            })
        return cards[:6]

    @classmethod
    def build(
        cls,
        dataframe=None,
        notes="",
        timeframe="",
        sentiment="all",
        segment="all",
        score_engine="",
        macro_trends="",
    ):
        row_count = 0
        top_service = ""
        if dataframe is not None and not getattr(dataframe, "empty", True):
            if "Reportable Analysis Row" in dataframe.columns:
                dataframe = dataframe[dataframe["Reportable Analysis Row"]].copy()
            row_count = len(dataframe)
            if "Layanan" in dataframe.columns:
                counts = dataframe["Layanan"].astype(str).value_counts()
                if not counts.empty:
                    top_service = cls._safe_text(counts.index[0], max_words=8)
        text_response_count = 0
        if dataframe is not None and not getattr(dataframe, "empty", True) and "Komentar" in dataframe.columns:
            comments = dataframe["Komentar"].astype(str).str.strip().replace({"nan": ""})
            text_response_count = int(comments.astype(bool).sum())
        focus_note = cls._safe_text(notes, max_words=38)
        if not focus_note:
            focus_note = "Analisis diarahkan pada sinyal evaluasi yang paling kuat dan paling dapat ditindaklanjuti."
        coverage_parts = [f"periode {cls._safe_text(timeframe, max_words=10) or 'terpilih'}"]
        if row_count:
            coverage_parts.append(f"{row_count} respons terolah")
        if sentiment != "all":
            coverage_parts.append(f"filter sentimen {cls._safe_text(sentiment, max_words=5)}")
        if segment != "all":
            coverage_parts.append(f"segmen {cls._safe_text(segment, max_words=8)}")
        if top_service:
            coverage_parts.append(f"layanan paling banyak muncul: {top_service}")
        coverage_note = "Cakupan pembacaan memakai " + ", ".join(coverage_parts) + "."
        external_context_ready = cls._external_context_ready(macro_trends)
        external_context_note = (
            "Pembanding eksternal cukup untuk membantu membaca tekanan pasar tanpa menggantikan bukti evaluasi internal."
            if external_context_ready
            else "Pembanding eksternal belum cukup kuat, sehingga pembacaan tetap mengutamakan bukti evaluasi internal."
        )
        insight_cards = cls._feedback_insight_cards(dataframe=dataframe, macro_trends=macro_trends)
        return {
            "focus_note": focus_note,
            "coverage_note": coverage_note,
            "external_context_ready": external_context_ready,
            "external_context_note": external_context_note,
            "insight_cards": insight_cards,
            "narrative_thesis": cls._safe_text(
                (insight_cards[0]["recommended_angle"] if insight_cards else focus_note),
                max_words=34,
            ),
            "paragraph_role_rotation": [
                "temuan yang terlihat",
                "bukti angka atau komentar",
                "tafsir yang masih wajar",
                "risiko jika dibiarkan",
                "keputusan atau tindakan manajemen",
            ],
            "dataset_roles": class_report_roles(),
            "voice_rules": feedback_voice_rules(),
            "section_cues": {
                "cx_chap_1": "baca catatan pengguna sebagai arah perhatian, bukan kalimat yang disalin mentah",
                "cx_chap_5": "ubah fokus pengguna menjadi implikasi kesiapan implementasi dan agenda manajemen",
                "executive": "angkat hanya keputusan yang didukung data, ringkas, dan bebas label teknis sumber",
            },
            "score_engine": cls._safe_text(score_engine, max_words=6),
            "row_count": row_count,
            "text_response_count": text_response_count,
        }

    @staticmethod
    def _external_context_ready(macro_trends):
        text = str(macro_trends or "").strip()
        lowered = text.lower()
        if not text or text == "-" or "tidak ada tren eksternal" in lowered or "tidak berhasil dimuat" in lowered:
            return False
        return bool(re.search(r"(?im)^\s*\d+\.\s+.+(?:\(Sumber:|sumber=)", text))


class ReportEvidenceBuilder:
    @staticmethod
    def clean(value, max_words=28):
        text = DocumentBuilder.reader_facing_text(str(value or ""))
        text = _scrub_pii(text)
        text = re.sub(r"https?://\S+", "", text)
        text = re.sub(r"(?i)\b(?:source|url|link)\s*=\s*[^|,\n]+", "", text)
        text = re.sub(r"\s+", " ", text).strip(" -;,.")
        words = text.split()
        if len(words) > max_words:
            text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
        return text

    @classmethod
    def summarize_proof(cls, value, max_words=32):
        text = cls.clean(value, max_words=90)
        text = re.sub(r"(?im)^\s*#{1,6}\s*", "", text)
        text = re.sub(r"\|[^|]*\|", " ", text)
        text = re.sub(r"\bBukti yang Dipakai\b", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+", " ", text).strip(" -;,.")
        if not text:
            return ""
        sentences = [part.strip(" -;,.") for part in re.split(r"(?<=[.!?])\s+|;\s+", text) if part.strip(" -;,.")]
        if sentences:
            sentences.sort(
                key=lambda sentence: (
                    bool(re.search(r"\b\d+(?:[,.]\d+)?%?\b", sentence)),
                    any(term in sentence.lower() for term in ("risiko", "feedback", "rating", "sentimen", "layanan", "keluhan")),
                    len(sentence),
                ),
                reverse=True,
            )
            text = sentences[0]
        words = text.split()
        if len(words) > max_words:
            text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
        if text and text[-1] not in ".!?":
            text += "."
        return text

    @classmethod
    def card_for_section(cls, section):
        title = cls.clean((section or {}).get("title"), max_words=10)
        content = cls.summarize_proof((section or {}).get("content"), max_words=30)
        if not content:
            return ""
        return f"### Bukti yang Dipakai\n- {title or 'Bagian laporan'} disusun dari bukti yang sudah diringkas: {content}"

    @classmethod
    def attach_to_sections(cls, report_sections):
        enriched = []
        for section in report_sections or []:
            copied = dict(section)
            content = str(copied.get("content") or "").strip()
            evidence = cls.card_for_section(copied)
            if evidence and not content.lstrip().startswith("### Bukti yang Dipakai"):
                copied["content"] = f"{evidence}\n\n{DocumentBuilder.reader_facing_text(content)}".strip()
            else:
                copied["content"] = DocumentBuilder.reader_facing_text(content)
            enriched.append(copied)
        return enriched

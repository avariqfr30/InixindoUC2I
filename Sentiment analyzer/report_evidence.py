import re

from document_builder import DocumentBuilder


RAW_SOURCE_PATTERNS = [
    r"\bAPIDog\b",
    r"\bInternal API\b",
    r"\bendpoint\b",
    r"\bsource\s*=",
    r"/api/Resource/dataset",
    r"\bEvidence Ledger\b",
]


class ReportEvidenceBuilder:
    @staticmethod
    def clean(value, max_words=28):
        text = DocumentBuilder.reader_facing_text(str(value or ""))
        text = re.sub(r"https?://\S+", "", text)
        text = re.sub(r"\s+", " ", text).strip(" -;,.")
        words = text.split()
        if len(words) > max_words:
            text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
        return text

    @classmethod
    def card_for_section(cls, section):
        title = cls.clean((section or {}).get("title"), max_words=10)
        content = cls.clean((section or {}).get("content"), max_words=30)
        if not content:
            return ""
        return f"### Bukti yang Dipakai\n- {title or 'Bagian laporan'} disusun dari ringkasan evaluasi yang menunjukkan {content.lower()}."

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

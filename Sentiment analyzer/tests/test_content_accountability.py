import unittest
import re

import pandas as pd

from narratives.prescriptive import PrescriptiveNarrativeMixin
from report_quality import ReportQualityValidator
from editorial_intelligence import repair_feedback_document_spine


class _PrescriptiveFixture(PrescriptiveNarrativeMixin):
    def _theme_hits(self, dataframe):
        return [{
            "id": "responsiveness",
            "label": "Kecepatan Respons",
            "negative_hits": 3,
            "prescription": "Tetapkan SLA respons dan jalur eskalasi untuk feedback prioritas.",
        }]


class FeedbackContentAccountabilityTest(unittest.TestCase):
    def test_priority_action_is_presented_once_with_complete_accountability(self):
        context = {
            "score_metrics": {"theme_rows": [{"theme_id": "responsiveness"}]},
            "score_profile": {"label": "Experience Index"},
        }

        content = _PrescriptiveFixture()._prescriptive_markdown(pd.DataFrame([{"feedback": "lambat"}]), context)

        self.assertEqual(content.count("Tetapkan SLA respons dan jalur eskalasi untuk feedback prioritas."), 1)
        self.assertIn("Owner Utama", content)
        self.assertIn("Batas Waktu", content)
        self.assertIn("Indikator Keberhasilan", content)
        self.assertIn("Dampak yang Diharapkan", content)

    def test_preflight_rejects_generic_recommendations_without_action_contract(self):
        filler = "Bukti feedback periode ini dibaca secara hati-hati agar keputusan layanan tetap terhubung dengan temuan sebelumnya. " * 2
        sections = [
            {"id": "cx_chap_1", "title": "Analisis Deskriptif", "content": filler},
            {"id": "cx_chap_2", "title": "Analisis Diagnostik", "content": "Melanjutkan Analisis Deskriptif. " + filler},
            {"id": "cx_chap_3", "title": "Analisis Prediktif", "content": "Berangkat dari Analisis Diagnostik. " + filler},
            {
                "id": "cx_chap_4",
                "title": "Rekomendasi Preskriptif",
                "content": "Rekomendasi perlu segera dijalankan agar layanan menjadi lebih baik. " + filler,
            },
            {"id": "cx_chap_5", "title": "Kesiapan Implementasi", "content": "Dari sini, kesiapan implementasi dibaca sebagai kelanjutan tindakan. " + filler},
        ]

        result = ReportQualityValidator.evaluate_narrative("Ringkasan eksekutif yang cukup substantif. " + filler, sections)

        self.assertIn("missing_action_contract", result["categories"])

    def test_spine_repair_varies_three_cross_chapter_openings(self):
        sections = []
        titles = ["Descriptive", "Diagnostic", "Predictive", "Prescriptive", "Implementation"]
        for index, title in enumerate(titles):
            repeated = (
                f"## {index + 1}.1 Konteks Pendukung\n{'- ' if index == 1 else ''}Ringkasan tren eksternal menunjukkan tekanan layanan yang perlu dibaca bersama bukti internal."
                if index < 3 else
                f"Bagian {title} menggunakan pembacaan yang berbeda untuk keputusan berikutnya."
            )
            previous = "" if index == 0 else f"{titles[index - 1]} menjadi dasar untuk {title}.\n\n"
            following = ""
            sections.append({"id": f"cx_chap_{index + 1}", "title": title, "content": previous + repeated + following})

        snapshot = "Ringkasan tren eksternal memberi konteks awal tanpa menggantikan bukti internal."
        repaired_snapshot, repaired = repair_feedback_document_spine(snapshot, sections)
        combined = "\n\n".join([repaired_snapshot, *[section["content"] for section in repaired]])
        repeated_starts = sum(line.strip().lstrip("- ").lower().startswith("ringkasan tren eksternal") for line in combined.splitlines())

        self.assertLessEqual(repeated_starts, 1)
        self.assertEqual(combined.lower().count("ringkasan tren eksternal"), 4)


if __name__ == "__main__":
    unittest.main()

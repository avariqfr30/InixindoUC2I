import pandas as pd

from .base import BaseNarrativeMixin


class PredictiveNarrativeMixin(BaseNarrativeMixin):
    def _predictive_markdown(self, timeframe_df, macro_trends, context, section_context=None):
        if timeframe_df.empty:
            return "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        service_risks = self._group_risk(timeframe_df, "Layanan", limit=5)
        stakeholder_risks = self._group_risk(timeframe_df, "Tipe Stakeholder", limit=5)
        location_risks = self._group_risk(timeframe_df, "Lokasi", limit=3)
        instructor_risks = self._group_risk(timeframe_df, "Tipe Instruktur", limit=3)
        journey_rows = context["journey_rows"]
        score_metrics = context["score_metrics"]

        risk_lines = [f"- {item['label']} diperkirakan tetap menjadi area {self._risk_severity(item['risk_score'])} karena proporsi sinyal negatif {item['negative_ratio']}% dengan rata-rata rating {item['average_rating']}. Jika tidak ada intervensi, skor pengalaman untuk layanan ini cenderung berada di bawah rata-rata periode berjalan." for item in service_risks] or ["- Tidak ada risiko layanan yang cukup kuat untuk diproyeksikan pada periode ini."]
        segment_lines = [f"- Segmen {item['label']} perlu dipantau karena volume {item['volume']} feedback dengan proporsi negatif {item['negative_ratio']}%. Tanpa penanganan, persepsi mereka berpotensi lebih rendah pada periode evaluasi berikutnya." for item in stakeholder_risks] or ["- Tidak ada segmen pelanggan yang cukup dominan untuk diproyeksikan."]
        operational_lines = [f"- Lokasi {item['label']} perlu dipantau karena proporsi sinyal negatifnya {item['negative_ratio']}% dengan rating rata-rata {item['average_rating']}." for item in location_risks] + [f"- Komposisi instruktur {item['label']} juga perlu dibaca karena saat ini mencatat proporsi sinyal negatif {item['negative_ratio']}%." for item in instructor_risks] or ["- Belum ada sinyal lokasi atau tipe instruktur yang cukup kuat untuk diproyeksikan."]

        journey_lines = []
        for item in journey_rows[:3]:
            if item["negative_share"] >= 25:
                journey_lines.append(f"- Tahap {item['stage_label']} diperkirakan tetap menjadi titik gesekan utama karena porsi sentimen negatif masih {item['negative_share']}%.")
            elif item["positive_share"] >= 60:
                journey_lines.append(f"- Tahap {item['stage_label']} cenderung tetap menjadi area yang lebih kuat karena porsi sentimen positif mencapai {item['positive_share']}%.")
            else:
                journey_lines.append(f"- Tahap {item['stage_label']} diperkirakan relatif stabil, tetapi perlu dipantau karena sentimennya masih bercampur.")
        if not journey_lines:
            journey_lines = ["- Belum ada pembacaan customer journey yang cukup kuat untuk dijadikan proyeksi."]

        section_context = section_context or {}
        external_ready = bool(section_context.get("external_context_ready", True))
        osint_signals = self._extract_osint_signals(macro_trends, limit=4) if external_ready else []
        deep_insight = self._extract_deep_insight(macro_trends)
        if not external_ready:
            deep_insight = ""
        osint_lines = self._summarized_external_trend_lines(osint_signals, deep_insight)
        if not osint_lines:
            osint_lines = ["- Pembanding eksternal belum cukup kuat; prediksi saat ini sepenuhnya didasarkan pada bukti evaluasi internal."]

        top_service_risk = service_risks[0] if service_risks else None
        top_segment_risk = stakeholder_risks[0] if stakeholder_risks else None
        top_issue = next((theme for theme in self._theme_hits(timeframe_df) if theme["negative_hits"] > 0), None)
        top_service_name = top_service_risk["label"] if top_service_risk else "layanan prioritas yang belum terpetakan"
        top_segment_name = top_segment_risk["label"] if top_segment_risk else "segmen utama yang belum terpetakan"
        top_issue_label = top_issue["label"] if top_issue else "pola kualitas layanan yang masih perlu dipantau"
        predictive_intro = (
            f"Jawaban prediktifnya: {top_service_name} perlu menjadi prioritas pengawasan jangka pendek, "
            f"terutama untuk segmen {top_segment_name}, karena pola feedback saat ini menunjukkan lensa pengalaman cenderung {score_metrics['direction']} dalam {context['horizon_text']}. "
            f"Fokus ini dibaca sebagai peringatan dini manajemen, bukan ramalan statistik."
        )
        predictive_reasons = [
            f"- **Mengapa area ini dulu:** {top_service_name} memiliki kombinasi risiko layanan paling jelas dibanding area lain pada periode ini.",
            f"- **Siapa yang paling perlu dipantau:** {top_segment_name} menunjukkan paparan yang paling layak dibahas dalam tinjauan berikutnya.",
            f"- **Apa isu penggeraknya:** {top_issue_label} menjadi tema yang paling perlu dicegah agar tidak berulang.",
        ]
        score_projection_line = (
            f"- **Arah lensa pengalaman:** pola saat ini cenderung {score_metrics['direction']} dalam {context['horizon_text']}. "
            "Pembacaan ini menggabungkan titik sentuh, rasa pengalaman, dan perjalanan pelanggan; bukan ukuran mentah yang berdiri sendiri."
        )
        forecast_30d = context.get("forecast_30d", {})
        weekly_forecast_table = self._markdown_table(
            ["Minggu 30 Hari", "Proyeksi Skor", "Pola", "Pembacaan Manajemen"],
            [
                [
                    item.get("week", "-"),
                    item.get("score", 0),
                    item.get("pattern", "-"),
                    item.get("reading", "-"),
                ]
                for item in forecast_30d.get("weekly_rows", [])
            ],
        )
        component_forecast_table = self._markdown_table(
            ["Komponen", "Bobot", "Saat Ini", "Proyeksi", "Perubahan"],
            [
                [
                    item.get("label", "-"),
                    f"{item.get('weight_pct', 0)}%",
                    item.get("current_score", 0),
                    item.get("projected_score", 0),
                    item.get("delta", 0),
                ]
                for item in forecast_30d.get("component_rows", [])
            ],
        )
        weekly_projection_chart = (
            f"[[LINE: Proyeksi Mingguan Experience Index 30 Hari | Skor | {forecast_30d.get('weekly_chart')}]]"
            if forecast_30d.get("weekly_chart")
            else ""
        )
        component_projection_chart = (
            f"[[CHART: Scorecard Komponen Experience Index | Skor Saat Ini | {forecast_30d.get('score_chart')}]]"
            if forecast_30d.get("score_chart")
            else ""
        )
        experience_lens_table = self._experience_lens_markdown(context)
        service_risk_table = self._markdown_table(["Layanan", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif", "Volume"], [[item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in service_risks])
        stakeholder_risk_table = self._markdown_table(["Segmen", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif", "Volume"], [[item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in stakeholder_risks])
        journey_projection_table = self._markdown_table(["Tahap Customer Journey", "Rating Rata-rata", "Negatif", "Positif", "Tema Dominan"], [[item["stage_label"], item["average_rating"], f"{item['negative_share']}%", f"{item['positive_share']}%", item["dominant_theme"]] for item in journey_rows])
        operational_projection_table = self._markdown_table(["Area Operasional", "Label", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif"], [["Lokasi", item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%"] for item in location_risks] + [["Tipe Instruktur", item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%"] for item in instructor_risks])
        projection_chart_line = f"[[CHART: Perbandingan Score Saat Ini vs Proyeksi | Skor | Saat Ini,{score_metrics['current_score']}; Proyeksi,{score_metrics['projected_score']}]]"
        company_linkage = self._external_company_linkage(
            osint_signals,
            deep_insight,
            top_service_name,
            top_segment_name,
            top_issue_label,
            self._projection_sentence(context),
        )
        defensibility_block = self._prediction_defensibility_markdown(timeframe_df, context, external_ready)

        return "\n".join([
            "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut", predictive_intro, "", *predictive_reasons, "",
            "Prediksi pada dokumen ini tidak dimaksudkan sebagai prakiraan statistik jangka panjang, melainkan sebagai peringatan dini berbasis pola penilaian, proporsi sentimen negatif, konsentrasi isu, dan lensa Experience Index. Dengan pendekatan ini, manajemen dapat lebih cepat memutuskan layanan mana yang perlu ditangani lebih dahulu.", "",
            score_projection_line, "",
            "### Storyboard Proyeksi 30 Hari Berbasis Mingguan",
            "Pembacaan mingguan memakai bucket rolling 7 hari dari pola historis ketika data tanggal memadai, lalu memakai fallback deterministik bila bukti bertanggal belum cukup rapat.",
            f"Tingkat keyakinan pembacaan mingguan: {forecast_30d.get('confidence', 'rendah')}. {forecast_30d.get('source_note', '')}",
            weekly_projection_chart,
            "",
            weekly_forecast_table,
            "",
            component_projection_chart,
            "",
            component_forecast_table,
            "",
            "### Makna Analitis Experience Index" if experience_lens_table else "",
            "Experience Index pada laporan ini bukan angka mentah; ia membaca pengalaman pelanggan sebagai gabungan titik sentuh layanan, rasa pengalaman yang dirasakan, dan perjalanan peserta saat mengikuti agenda perusahaan." if experience_lens_table else "",
            "" if experience_lens_table else "",
            experience_lens_table if experience_lens_table else "",
            "",
            defensibility_block, "",
            projection_chart_line, "", service_risk_table, "", *risk_lines, "", "## 3.2 Prediksi Segmen dan Layanan yang Paling Rentan", "Selain layanan, pemantauan juga perlu diarahkan pada segmen pelanggan yang memperlihatkan kombinasi antara volume feedback tinggi dan kualitas pengalaman yang menurun. Segmen seperti ini biasanya lebih cepat mempengaruhi reputasi, retensi, dan peluang repeat engagement.", "",
            stakeholder_risk_table, "", *segment_lines, "", "### Pembacaan customer journey ke depan", journey_projection_table, "", *journey_lines, "", "### Area operasional yang perlu diawasi", operational_projection_table, "", *operational_lines, "",
            "## 3.3 Tren Eksternal yang Berpotensi Memperbesar Risiko", ("Pembanding eksternal pada bagian ini dipakai secara konservatif. Ketika sinyal publik belum cukup sebanding, prediksi tetap mengutamakan bukti evaluasi internal, pola penilaian, dan konsentrasi umpan balik." if not external_ready else "Ringkasan tren eksternal disederhanakan menjadi implikasi manajemen agar konteks pasar mendukung temuan internal tanpa menjadi daftar sumber mentah."), "",
            *osint_lines[:3],
            "## 3.4 Keterkaitan Faktor Eksternal dengan Kondisi Perusahaan", *company_linkage,
        ])

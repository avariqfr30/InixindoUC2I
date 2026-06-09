import pandas as pd

from .base import BaseNarrativeMixin


class DiagnosticNarrativeMixin(BaseNarrativeMixin):
    def _diagnostic_markdown(self, timeframe_df, context):
        if timeframe_df.empty:
            return "## 2.1 Akar Masalah Utama dan Pain Point Dominan\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        theme_hits = self._theme_hits(timeframe_df)
        theme_lookup = {theme["id"]: theme for theme in theme_hits}
        prioritized_theme_rows = context["score_metrics"]["theme_rows"]
        prioritized_negative_ids = [item["theme_id"] for item in prioritized_theme_rows if theme_lookup.get(item["theme_id"], {}).get("negative_hits", 0) > 0][:3]
        negative_themes = [theme_lookup[theme_id] for theme_id in prioritized_negative_ids]
        if not negative_themes:
            negative_themes = [theme for theme in theme_hits if theme["negative_hits"] > 0][:3]
        positive_themes = sorted(theme_hits, key=lambda item: (item["positive_hits"], item["total_hits"]), reverse=True)[:3]

        if not negative_themes:
            negative_lines = ["- Belum ada pola keluhan dominan yang menonjol; mayoritas feedback berada pada area stabil."]
        else:
            negative_lines = []
            for theme in negative_themes:
                impacted_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
                impacted_segments = self._series_counts(theme["matched_df"]["Tipe Stakeholder"], limit=2)
                negative_lines.append(f"- {theme['label']}: {theme['negative_hits']} sinyal negatif. Layanan terdampak: {', '.join(impacted_services.index.tolist()) or 'belum terpetakan'}. Segmen terdampak: {', '.join(impacted_segments.index.tolist()) or 'belum terpetakan'}.")

        positive_lines = []
        for theme in positive_themes:
            if theme["positive_hits"] <= 0:
                continue
            strongest_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
            positive_lines.append(f"- {theme['label']}: {theme['positive_hits']} sinyal positif. Paling banyak muncul pada layanan {', '.join(strongest_services.index.tolist()) or 'belum terpetakan'}.")
        if not positive_lines:
            positive_lines = ["- Belum ada kekuatan yang cukup konsisten untuk dikonfirmasi pada periode ini."]

        negative_quotes = self._quote_lines(timeframe_df[timeframe_df["Sentiment Label"].isin({"negative", "mixed"})], limit=3)
        positive_quotes = self._quote_lines(timeframe_df[timeframe_df["Sentiment Label"] == "positive"], limit=2)

        service_risks = self._group_risk(timeframe_df, "Layanan", limit=5)
        location_risks = self._group_risk(timeframe_df, "Lokasi", limit=3)
        instructor_risks = self._group_risk(timeframe_df, "Tipe Instruktur", limit=3)
        process_gap_lines = [f"- {item['label']}: rata-rata rating {item['average_rating']}, proporsi negatif {item['negative_ratio']}%, volume {item['volume']}." for item in service_risks] or ["- Belum ada gap proses yang dapat dipetakan."]
        top_issue = negative_themes[0] if negative_themes else None
        top_strength = next((theme for theme in positive_themes if theme["positive_hits"] > 0 and (not top_issue or theme["id"] != top_issue["id"])), None) or next((theme for theme in positive_themes if theme["positive_hits"] > 0), None)
        dominant_journey = context["dominant_journey"]

        if top_issue and top_strength and top_issue["id"] == top_strength["id"]:
            strength_context = f"Menariknya, tema {top_strength['label']} muncul sebagai area yang terpolarisasi: sebagian pelanggan menilai sangat baik, sementara sebagian lain masih mengalami hambatan."
        elif top_strength:
            strength_context = f"Di sisi lain, kekuatan yang paling konsisten terlihat pada {top_strength['label']}."
        else:
            strength_context = "Kekuatan layanan belum muncul secara cukup konsisten untuk dijadikan diferensiasi yang kuat."

        diagnostic_intro = f"Analisis diagnostik bertujuan menjawab mengapa pola feedback pada periode ini muncul. Karena laporan dibaca dari sudut pandang {context['score_profile']['label']}, perhatian diagnosis terutama diarahkan ke {context['score_profile']['narrative_focus']}. {'Tema keluhan paling dominan saat ini adalah ' + top_issue['label'] + ', yang berulang pada beberapa komentar pelanggan.' if top_issue else 'Belum ada tema keluhan yang sangat dominan, sehingga pola masalah masih relatif tersebar.'} {strength_context}"
        # Corrected English 'and' leak to Indonesian 'dan' in journey narrative
        journey_diagnostic = f"Jika dibaca menurut customer journey, titik gesekan yang paling terasa saat ini berada pada tahap {dominant_journey['stage_label']} dengan rating rata-rata {dominant_journey['average_rating']} dan porsi sentimen negatif {dominant_journey['negative_share']}%." if dominant_journey else "Pemetaan customer journey belum menunjukkan titik gesekan yang dominan."

        root_cause_table_rows = [[theme["label"], theme["negative_hits"], ", ".join(self._series_counts(theme["matched_df"]["Layanan"], limit=2).index.tolist()) or "Belum terpetakan", ", ".join(self._series_counts(theme["matched_df"]["Tipe Stakeholder"], limit=2).index.tolist()) or "Belum terpetakan"] for theme in negative_themes]
        root_cause_table = self._markdown_table(["Tema Prioritas", "Sinyal Negatif", "Layanan Dominan", "Segmen Dominan"], root_cause_table_rows)
        strength_table_rows = [[theme["label"], theme["positive_hits"], ", ".join(self._series_counts(theme["matched_df"]["Layanan"], limit=2).index.tolist()) or "Belum terpetakan"] for theme in positive_themes if theme["positive_hits"] > 0]
        strength_table = self._markdown_table(["Kekuatan", "Sinyal Positif", "Layanan Dominan"], strength_table_rows)
        service_risk_table = self._markdown_table(["Layanan", "Rata-rata Rating", "Proporsi Negatif", "Volume", "Skor Risiko"], [[item["label"], item["average_rating"], f"{item['negative_ratio']}%", item["volume"], item["risk_score"]] for item in service_risks])
        location_instructor_table = self._markdown_table(["Area Analisis", "Label", "Rata-rata Rating", "Proporsi Negatif", "Volume"], [["Lokasi", item["label"], item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in location_risks] + [["Tipe Instruktur", item["label"], item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in instructor_risks])
        operational_context = f"Dari sisi lokasi dan model instruktur, area yang perlu dicermati lebih dekat adalah {location_risks[0]['label'] if location_risks else 'lokasi yang belum terpetakan'} serta komposisi instruktur {instructor_risks[0]['label'] if instructor_risks else 'yang belum terpetakan'}. Pembacaan ini membantu membedakan apakah masalah lebih banyak terkait kesiapan tempat, model pengajar, atau memang tema layanan itu sendiri."

        return "\n".join([
            "## 2.1 Akar Masalah Utama dan Pain Point Dominan", diagnostic_intro, "", "Pembacaan akar masalah dilakukan dengan melihat pengulangan tema, dampaknya pada layanan, dan segmen pelanggan yang paling sering menyinggung isu serupa. Dengan pendekatan ini, tim manajemen dapat membedakan antara keluhan yang bersifat insidental dan keluhan yang sudah layak dibaca sebagai pola struktural.", "",
            root_cause_table, "", *negative_lines, "", "## 2.2 Kekuatan yang Konsisten dan Area yang Perlu Dijaga", "Selain keluhan, periode ini juga memperlihatkan area yang secara berulang dijadikan acuan untuk standardisasi layanan, replikasi praktik baik, dan bahan komunikasi nilai kepada klien.", "",
            strength_table, "", *positive_lines, "", "## 2.3 Bukti Verbatim, Kesenjangan Proses, dan Segmentasi Masalah", "Bukti verbatim di bawah ini digunakan untuk menjaga agar interpretasi manajerial tetap berpijak pada suara pelanggan. Ringkasan kesenjangan proses membantu menerjemahkan komentar individual ke dalam area operasional yang dapat ditindaklanjuti.", "",
            journey_diagnostic, "", "### Kutipan keluhan representatif", *negative_quotes, "### Kutipan apresiasi representatif", *positive_quotes, "### Kesenjangan proses yang paling terlihat", service_risk_table, "", *process_gap_lines, "",
            "### Konteks lokasi pelatihan dan tipe instruktur", operational_context, "", location_instructor_table,
        ])

import pandas as pd

from .base import BaseNarrativeMixin


class DescriptiveNarrativeMixin(BaseNarrativeMixin):
    def _descriptive_markdown(self, timeframe_df, timeframe, notes, context, section_context=None):
        timeframe_label = context.get("timeframe_label", timeframe)
        governance = self._governance_summary(timeframe_df)
        total_rows = governance["total_rows"]
        if total_rows == 0:
            return "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        avg_rating = timeframe_df["Rating Numeric"].mean()
        sentiment_summary = self._sentiment_summary(timeframe_df)
        positive_count = sentiment_summary["positive"]
        mixed_count = sentiment_summary["mixed"]
        neutral_count = sentiment_summary["neutral"]
        negative_count = sentiment_summary["negative"]
        weak_negative_count = sentiment_summary["weak_negative"]

        stakeholder_counts = self._series_counts(timeframe_df["Tipe Stakeholder"])
        service_counts = self._series_counts(timeframe_df["Layanan"])
        positive_share = sentiment_summary["positive_share"]
        mixed_share = sentiment_summary["mixed_share"]
        neutral_share = sentiment_summary["neutral_share"]
        negative_share = sentiment_summary["issue_share"]
        strong_negative_share = sentiment_summary["negative_share"]
        weak_negative_share = sentiment_summary["weak_negative_share"]
        score_metrics = context["score_metrics"]
        journey_rows = context["journey_rows"]
        scope_text = context["scope_text"]
        location_counts = context["location_counts"]
        instructor_type_counts = context["instructor_type_counts"]

        section_context = section_context or {}
        cleaned_notes = str(section_context.get("focus_note") or notes or "").strip().rstrip(".!?")
        focus_line = f"Fokus pembacaan tambahan pada periode ini adalah {cleaned_notes}." if cleaned_notes else "Tidak ada fokus tambahan dari pengguna, sehingga analisis dilakukan terhadap seluruh sinyal yang tersedia."
        governance_note = "Cakupan sumber sudah mulai terpetakan, tetapi pemetaan kanal masih perlu diperkuat." if governance["channel_count"] == 0 else "Pemetaan sumber dan kanal sudah tersedia sehingga jalur asal feedback lebih mudah diaudit."

        descriptive_intro = (
            f"Bagian ini menjelaskan kualitas dasar portofolio feedback yang menjadi fondasi laporan. Analisis dibaca pada {scope_text}. "
            f"Fokus pembacaannya menekankan {context['score_profile']['narrative_focus']}. "
            f"Pada periode {timeframe_label}, sistem memproses {total_rows} feedback tervalidasi dengan "
            f"rata-rata rating {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5, yang menunjukkan kinerja "
            f"layanan berada pada kategori {self._rating_assessment(avg_rating)}. "
            f"Komposisi sentimen memperlihatkan {positive_share}% sinyal positif, {mixed_share}% kritik konstruktif, "
            f"{neutral_share}% sinyal netral, {strong_negative_share}% sinyal negatif kuat, dan {weak_negative_share}% sinyal negatif bukti lemah."
        )
        governance_intro = f"Dari sisi tata kelola, kelengkapan field inti mencapai {governance['completeness_pct']}%. Data berasal dari {governance['source_count']} sumber feedback dan {governance['channel_count']} kanal yang terpetakan. {governance_note} {focus_line}"
        indicator_table = self._markdown_table(
            ["Indikator", "Nilai"],
            [
                ["Periode analisis", timeframe_label], ["Cakupan analisis", scope_text], ["Total feedback tervalidasi", f"{total_rows} record"],
                ["Rata-rata rating", f"{round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5"], [context["score_profile"]["label"], f"{score_metrics['current_score']} / 100"],
                ["Sumber parameter skor", context["score_profile"].get("parameter_source", "Model internal")],
                ["Kelengkapan field inti", f"{governance['completeness_pct']}%"], ["Jumlah sumber feedback", governance["source_count"]], ["Jumlah kanal feedback", governance["channel_count"]],
            ],
        )
        score_table = self._markdown_table(
            ["Score Engine", "Nilai Saat Ini", "Arah Bacaan", "Tema Paling Berpengaruh"],
            [[context["score_profile"]["label"], f"{score_metrics['current_score']}", score_metrics["direction"].title(), context["dominant_theme"]["label"] if context["dominant_theme"] else "Belum terpetakan"]]
        )

        sentiment_chart_line = f"[[PIE: Komposisi Sentimen Feedback | Positif,{positive_share}; Kritik konstruktif,{mixed_share}; Netral,{neutral_share}; Negatif,{strong_negative_share}; Negatif bukti lemah,{weak_negative_share}]]"
        journey_chart_line = "[[CHART: Titik Customer Journey dengan Sinyal Negatif | Persentase Negatif | " + self._chart_pairs(pd.Series({item["stage_label"]: item["negative_share"] for item in journey_rows}), use_percentage=False, limit=4) + "]]" if journey_rows else ""
        sentiment_table = self._markdown_table(
            ["Kategori Sentimen", "Jumlah", "Persentase"],
            [
                ["Positif", positive_count, f"{positive_share}%"],
                ["Kritik konstruktif", mixed_count, f"{mixed_share}%"],
                ["Netral", neutral_count, f"{neutral_share}%"],
                ["Negatif", negative_count, f"{strong_negative_share}%"],
                ["Negatif bukti lemah", weak_negative_count, f"{weak_negative_share}%"],
            ],
        )
        stakeholder_table = self._markdown_table(["Segmen Stakeholder", "Jumlah Feedback", "Persentase"], self._distribution_rows(stakeholder_counts, total_rows, limit=5))
        service_table = self._markdown_table(["Layanan", "Jumlah Feedback", "Persentase"], self._distribution_rows(service_counts, total_rows, limit=5))
        location_table = self._markdown_table(["Lokasi Pelatihan", "Jumlah Feedback", "Persentase"], self._distribution_rows(location_counts, total_rows, limit=5))
        instructor_type_table = self._markdown_table(["Tipe Instruktur", "Jumlah Feedback", "Persentase"], self._distribution_rows(instructor_type_counts, total_rows, limit=5))

        location_pie_line = "[[PIE: Sebaran Lokasi Pelatihan | " + self._chart_pairs(location_counts, total_rows=total_rows, limit=5, use_percentage=True) + "]]" if not location_counts.empty else ""
        instructor_pie_line = "[[PIE: Komposisi Instruktur Internal vs OL | " + self._chart_pairs(instructor_type_counts, total_rows=total_rows, limit=5, use_percentage=True) + "]]" if not instructor_type_counts.empty else ""
        distribution_paragraph = f"Sebaran volume feedback menunjukkan bahwa konsentrasi terbesar berasal dari segmen {self._format_count_summary(stakeholder_counts, limit=3)}. Dari sisi layanan, perhatian pengguna paling banyak tercurah pada {self._format_count_summary(service_counts, limit=3)}. Pola ini penting untuk dibaca secara hati-hati, karena volume tinggi belum otomatis berarti performa buruk, tetapi menandakan area yang paling banyak terekspos kepada pelanggan."
        delivery_context_paragraph = f"Lokasi pelatihan pada cakupan terpilih paling banyak berlangsung di {self._format_count_summary(location_counts, limit=3)}. Dari sisi tipe instruktur, komposisi saat ini didominasi oleh {self._format_count_summary(instructor_type_counts, limit=3)}. Informasi ini penting karena performa layanan sering kali dipengaruhi oleh kesiapan lokasi, format delivery, dan model pengajar yang dipakai."
        journey_table = self._markdown_table(["Tahap Customer Journey", "Volume", "Rating Rata-rata", "Positif", "Netral", "Negatif", "Tema Dominan"], [[item["stage_label"], item["volume"], item["average_rating"], f"{item['positive_share']}%", f"{item['neutral_share']}%", f"{item['negative_share']}%", item["dominant_theme"]] for item in journey_rows])
        dominant_journey_text = f"Sentimen paling menantang pada filter yang dipilih saat ini muncul pada tahap {context['dominant_journey']['stage_label']} dengan porsi sinyal negatif {context['dominant_journey']['negative_share']}%." if context["dominant_journey"] else "Belum ada tahap customer journey yang dapat dipetakan secara cukup kuat."

        return "\n".join([
            "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola", descriptive_intro, "",
            governance_intro, "", indicator_table, "",
            "## 1.2 Distribusi Sentimen, Rating, dan Volume", f"Distribusi sentimen menunjukkan bahwa proporsi sinyal korektif tertimbang sebesar {negative_share}% {self._negative_share_assessment(negative_share)}. Sentimen positif tetap menjadi penopang utama pengalaman pelanggan, tetapi kritik konstruktif dan sinyal negatif bukti lemah dipisahkan agar manajemen tidak membaca semua rating tinggi sebagai pujian penuh.", "",
            score_table, "", sentiment_table, "", "Visual berikut memperlihatkan distribusi sentimen untuk kombinasi input yang dipilih, sehingga pembaca dapat melihat proporsi positif, kritik konstruktif, netral, negatif, dan sinyal negatif yang bukti komentarnya masih lemah.", "", sentiment_chart_line, "",
            "## 1.3 Distribusi Stakeholder, Layanan, dan Kanal/Sumber", distribution_paragraph, "", "### Stakeholder dengan volume feedback terbesar", stakeholder_table, "", "### Layanan dengan volume feedback terbesar", service_table, "",
            "### Pemetaan sentimen pada customer journey", dominant_journey_text, "", journey_table, "", "Visual berikut membantu melihat tahapan customer journey mana yang paling banyak menampung sinyal negatif pada input yang dipilih.", "", journey_chart_line, "",
            "### Lokasi pelatihan dan tipe instruktur", delivery_context_paragraph, "", location_table, "", location_pie_line, "", instructor_type_table, "", instructor_pie_line,
        ])

import re

import pandas as pd

from config import ADOPTION_READINESS_PILLARS, CX_SENTIMENT_STRUCTURE, DEFAULT_SCORE_ENGINE


class ReportNarrativeBuilderMixin:
    """Markdown report narrative builders for FeedbackAnalyticsEngine.

    The mixin deliberately depends on analytics/context helper methods supplied by
    FeedbackAnalyticsEngine; it owns wording, tables, and report section assembly.
    """

    @staticmethod
    def _escape_table_cell(value):
        return str(value).replace("|", "\\|").replace("\n", " ").strip()

    @classmethod
    def _markdown_table(cls, headers, rows):
        if not rows: return ""
        header_line = "| " + " | ".join(cls._escape_table_cell(item) for item in headers) + " |"
        separator_line = "| " + " | ".join("---" for _ in headers) + " |"
        row_lines = ["| " + " | ".join(cls._escape_table_cell(cell) for cell in row) + " |" for row in rows]
        return "\n".join([header_line, separator_line, *row_lines])

    def _distribution_rows(self, series_counts, total_rows, limit=5):
        return [[label, count, f"{self._safe_percentage(count, total_rows)}%"] for label, count in series_counts.head(limit).items()]

    def _extract_osint_signals(self, macro_trends, limit=3):
        signals = []
        for line in str(macro_trends).splitlines():
            cleaned = line.strip()
            if not re.match(r"^\d+\.", cleaned): continue
            cleaned = re.sub(r"^\d+\.\s*", "", cleaned)
            parts = [part.strip() for part in cleaned.split(" | ") if part.strip()]
            if not parts: continue
            title = parts[0]
            snippet = parts[1] if len(parts) > 1 else ""
            source, date = "Tidak diketahui", "-"
            for part in parts[2:]:
                if part.startswith("sumber="): source = part.split("=", maxsplit=1)[1] or source
                elif part.startswith("tanggal="): date = part.split("=", maxsplit=1)[1] or date
            signals.append({"title": title, "snippet": snippet, "source": source, "date": date})
            if len(signals) >= limit: break
        return signals

    @staticmethod
    def _extract_deep_insight(macro_trends):
        match = re.search(r"\*\*Insight Mendalam[^*]*\*\*\s*(.*)", str(macro_trends))
        if match:
            return match.group(0)
        return ""

    @staticmethod
    def _theme_owner(theme_id):
        owner_map = {"responsiveness": "Customer Service / Account Management", "schedule": "Operations / Delivery Management", "facility": "Operations / General Affairs", "instructor": "Academic Lead / Service Quality", "material": "Academic Lead / Product Owner", "communication": "Customer Service / Project Coordinator", "outcome": "Service Owner / Quality Assurance"}
        return owner_map.get(theme_id, "Service Owner")

    @staticmethod
    def _theme_outcome(theme_id):
        outcome_map = {"responsiveness": "Waktu respons lebih konsisten dan penutupan isu lebih cepat.", "schedule": "Pengalaman delivery lebih tertata dan beban sesi lebih seimbang.", "facility": "Gangguan operasional di kelas atau sesi layanan dapat ditekan.", "instructor": "Konsistensi kualitas fasilitator meningkat di berbagai layanan.", "material": "Materi lebih relevan dengan kebutuhan peserta dan konteks klien.", "communication": "Ekspektasi stakeholder lebih selaras sejak pra-delivery hingga pasca-delivery.", "outcome": "Nilai manfaat layanan lebih mudah dirasakan dan dibuktikan."}
        return outcome_map.get(theme_id, "Persepsi kualitas layanan membaik secara terukur.")

    @staticmethod
    def _readiness_label(score):
        return "Kuat" if score >= 80 else "Cukup Siap" if score >= 60 else "Perlu Diperkuat" if score >= 40 else "Prioritas Tinggi"

    def _projection_sentence(self, context):
        metrics = context["score_metrics"]
        score_label = context["score_profile"]["forecast_label"]
        horizon_text = context["horizon_text"]
        calendar_reference = self._forecast_calendar_reference(context["timeframe"])

        if metrics["direction"] == "turun":
            return f"{score_label} diproyeksikan turun dari {metrics['current_score']} menjadi sekitar {metrics['projected_score']} dalam {horizon_text}, atau {calendar_reference}, apabila pola saat ini berlanjut."
        if metrics["direction"] == "naik":
            return f"{score_label} diproyeksikan naik dari {metrics['current_score']} menjadi sekitar {metrics['projected_score']} dalam {horizon_text}, atau {calendar_reference}, jika momentum yang ada dapat dipertahankan."
        return f"{score_label} diproyeksikan relatif stabil di kisaran {metrics['projected_score']} dalam {horizon_text}, atau {calendar_reference}, namun tetap perlu dipantau agar tidak bergeser ketika volume feedback bertambah."

    @staticmethod
    def _score_component_formula(components, key_name):
        if not components:
            return 0.0, ""
        total = 0.0
        terms = []
        for item in components:
            weight = float(item.get("weight", 0.0))
            score_value = float(item.get(key_name, 0.0))
            total += weight * score_value
            terms.append(f"({round(weight * 100, 1)}% x {score_value})")
        return round(total, 2), " + ".join(terms)

    def _experience_formula_details(self, context):
        if context.get("score_engine") != "experience_index":
            return None
        components = context.get("score_metrics", {}).get("component_breakdown") or []
        if not components:
            return None

        current_calc, current_formula = self._score_component_formula(components, "current_score")
        projected_calc, projected_formula = self._score_component_formula(components, "projected_score")
        weight_summary = ", ".join(
            f"{item['label']} {round(float(item['weight']) * 100, 1)}%"
            for item in components
        )
        return {
            "weight_summary": weight_summary,
            "current_calc": current_calc,
            "current_formula": current_formula,
            "projected_calc": projected_calc,
            "projected_formula": projected_formula,
        }

    def _descriptive_markdown(self, timeframe_df, timeframe, notes, context):
        governance = self._governance_summary(timeframe_df)
        total_rows = governance["total_rows"]
        if total_rows == 0:
            return "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        avg_rating = timeframe_df["Rating Numeric"].mean()
        positive_count = int((timeframe_df["Sentiment Label"] == "positive").sum())
        neutral_count = int((timeframe_df["Sentiment Label"] == "neutral").sum())
        negative_count = int((timeframe_df["Sentiment Label"] == "negative").sum())

        stakeholder_counts = self._series_counts(timeframe_df["Tipe Stakeholder"])
        service_counts = self._series_counts(timeframe_df["Layanan"])
        source_counts = self._series_counts(timeframe_df["Sumber Feedback"])
        channel_counts = self._series_counts(timeframe_df["Kanal Feedback"])

        top_sources = source_counts.index.tolist() if not source_counts.empty else ["Sumber internal terstandar"]
        top_channels = channel_counts.index.tolist() if not channel_counts.empty else ["Belum terpetakan"]
        positive_share = self._safe_percentage(positive_count, total_rows)
        neutral_share = self._safe_percentage(neutral_count, total_rows)
        negative_share = self._safe_percentage(negative_count, total_rows)
        score_metrics = context["score_metrics"]
        journey_rows = context["journey_rows"]
        scope_text = context["scope_text"]
        location_counts = context["location_counts"]
        instructor_type_counts = context["instructor_type_counts"]

        cleaned_notes = notes.strip().rstrip(".!?")
        focus_line = f"Fokus tambahan dari pengguna pada periode ini adalah: {cleaned_notes}." if notes and notes.strip() else "Tidak ada fokus tambahan dari pengguna, sehingga analisis dilakukan terhadap seluruh sinyal yang tersedia."
        governance_note = "Cakupan sumber sudah mulai terpetakan, tetapi pemetaan kanal masih perlu diperkuat." if governance["channel_count"] == 0 else "Pemetaan sumber dan kanal sudah tersedia sehingga jalur asal feedback lebih mudah diaudit."
        
        descriptive_intro = (
            f"Bagian ini menjelaskan kualitas dasar portofolio feedback yang menjadi fondasi laporan. Analisis dibaca pada {scope_text}. "
            f"Fokus pembacaannya menekankan {context['score_profile']['narrative_focus']}. "
            f"Pada periode {timeframe}, sistem memproses {total_rows} feedback tervalidasi dengan "
            f"rata-rata rating {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5, yang menunjukkan kinerja "
            f"layanan berada pada kategori {self._rating_assessment(avg_rating)}. "
            f"Komposisi sentimen memperlihatkan {positive_share}% sinyal positif, {neutral_share}% sinyal netral, "
            f"dan {negative_share}% sinyal negatif."
        )
        governance_intro = f"Dari sisi tata kelola, kelengkapan field inti mencapai {governance['completeness_pct']}%. Data berasal dari {governance['source_count']} sumber feedback dan {governance['channel_count']} kanal yang terpetakan. {governance_note} {focus_line}"
        indicator_table = self._markdown_table(
            ["Indikator", "Nilai"],
            [
                ["Periode analisis", timeframe], ["Cakupan analisis", scope_text], ["Total feedback tervalidasi", f"{total_rows} record"],
                ["Rata-rata rating", f"{round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5"], [context["score_profile"]["label"], f"{score_metrics['current_score']} / 100"],
                ["Sumber parameter skor", context["score_profile"].get("parameter_source", "Model internal")],
                ["Kelengkapan field inti", f"{governance['completeness_pct']}%"], ["Jumlah sumber feedback", governance["source_count"]], ["Jumlah kanal feedback", governance["channel_count"]],
            ],
        )
        score_table = self._markdown_table(
            ["Score Engine", "Nilai Saat Ini", "Arah Bacaan", "Tema Paling Berpengaruh"],
            [[context["score_profile"]["label"], f"{score_metrics['current_score']}", score_metrics["direction"].title(), context["dominant_theme"]["label"] if context["dominant_theme"] else "Belum terpetakan"]]
        )

        sentiment_chart_line = f"[[PIE: Komposisi Sentimen Feedback | Positif,{positive_share}; Netral,{neutral_share}; Negatif,{negative_share}]]"
        journey_chart_line = "[[CHART: Titik Customer Journey dengan Sinyal Negatif | Persentase Negatif | " + self._chart_pairs(pd.Series({item["stage_label"]: item["negative_share"] for item in journey_rows}), use_percentage=False, limit=4) + "]]" if journey_rows else ""
        sentiment_table = self._markdown_table(["Kategori Sentimen", "Jumlah", "Persentase"], [["Positif", positive_count, f"{positive_share}%"], ["Netral", neutral_count, f"{neutral_share}%"], ["Negatif", negative_count, f"{negative_share}%"]])
        stakeholder_table = self._markdown_table(["Segmen Stakeholder", "Jumlah Feedback", "Persentase"], self._distribution_rows(stakeholder_counts, total_rows, limit=5))
        service_table = self._markdown_table(["Layanan", "Jumlah Feedback", "Persentase"], self._distribution_rows(service_counts, total_rows, limit=5))
        location_table = self._markdown_table(["Lokasi Pelatihan", "Jumlah Feedback", "Persentase"], self._distribution_rows(location_counts, total_rows, limit=5))
        instructor_type_table = self._markdown_table(["Tipe Instruktur", "Jumlah Feedback", "Persentase"], self._distribution_rows(instructor_type_counts, total_rows, limit=5))
        
        location_pie_line = "[[PIE: Sebaran Lokasi Pelatihan | " + self._chart_pairs(location_counts, total_rows=total_rows, limit=5, use_percentage=True) + "]]" if not location_counts.empty else ""
        instructor_pie_line = "[[PIE: Komposisi Instruktur Internal vs OL | " + self._chart_pairs(instructor_type_counts, total_rows=total_rows, limit=5, use_percentage=True) + "]]" if not instructor_type_counts.empty else ""
        source_lines = [f"- Sumber utama: {', '.join(top_sources[:3])}", f"- Kanal utama: {', '.join(top_channels[:3])}"]
        
        distribution_paragraph = f"Sebaran volume feedback menunjukkan bahwa konsentrasi terbesar berasal dari segmen {self._format_count_summary(stakeholder_counts, limit=3)}. Dari sisi layanan, perhatian pengguna paling banyak tercurah pada {self._format_count_summary(service_counts, limit=3)}. Pola ini penting untuk dibaca secara hati-hati, karena volume tinggi belum otomatis berarti performa buruk, tetapi menandakan area yang paling banyak terekspos kepada pelanggan."
        source_paragraph = f"Dari sisi asal data, sumber yang paling dominan saat ini adalah {', '.join(top_sources[:3])}. Pada saat yang sama, kanal yang tercatat masih didominasi oleh {', '.join(top_channels[:3])}. Informasi ini perlu dibaca sebagai indikator awal representativitas data: semakin luas sumber dan kanal, semakin kuat dasar analisis untuk pengambilan keputusan lintas fungsi."
        delivery_context_paragraph = f"Lokasi pelatihan pada cakupan terpilih paling banyak berlangsung di {self._format_count_summary(location_counts, limit=3)}. Dari sisi tipe instruktur, komposisi saat ini didominasi oleh {self._format_count_summary(instructor_type_counts, limit=3)}. Informasi ini penting karena performa layanan sering kali dipengaruhi oleh kesiapan lokasi, format delivery, dan model pengajar yang dipakai."
        journey_table = self._markdown_table(["Tahap Customer Journey", "Volume", "Rating Rata-rata", "Positif", "Netral", "Negatif", "Tema Dominan"], [[item["stage_label"], item["volume"], item["average_rating"], f"{item['positive_share']}%", f"{item['neutral_share']}%", f"{item['negative_share']}%", item["dominant_theme"]] for item in journey_rows])
        dominant_journey_text = f"Sentimen paling menantang pada filter yang dipilih saat ini muncul pada tahap {context['dominant_journey']['stage_label']} dengan porsi sinyal negatif {context['dominant_journey']['negative_share']}%." if context["dominant_journey"] else "Belum ada tahap customer journey yang dapat dipetakan secara cukup kuat."

        return "\n".join([
            "## 1.1 Ringkasan Cakupan Feedback dan Tata Kelola", descriptive_intro, "", governance_intro, "", indicator_table, "",
            "## 1.2 Distribusi Sentimen, Rating, dan Volume", f"Distribusi sentimen menunjukkan bahwa proporsi sentimen negatif sebesar {negative_share}% {self._negative_share_assessment(negative_share)}. Sentimen positif tetap menjadi penopang utama pengalaman pelanggan, tetapi keberadaan sentimen netral yang cukup material mengindikasikan masih ada ruang untuk memperkuat pengalaman agar tidak berhenti pada persepsi 'cukup'.", "",
            score_table, "", sentiment_table, "", "Visual berikut memperlihatkan distribusi sentimen untuk kombinasi input yang dipilih, sehingga pembaca dapat segera melihat apakah pengalaman pelanggan lebih banyak berada di area positif, netral, atau negatif.", "", sentiment_chart_line, "",
            "## 1.3 Distribusi Stakeholder, Layanan, dan Kanal/Sumber", distribution_paragraph, "", "### Stakeholder dengan volume feedback terbesar", stakeholder_table, "", "### Layanan dengan volume feedback terbesar", service_table, "",
            "### Pemetaan sentimen pada customer journey", dominant_journey_text, "", journey_table, "", "Visual berikut membantu melihat tahapan customer journey mana yang paling banyak menampung sinyal negatif pada input yang dipilih.", "", journey_chart_line, "",
            "### Lokasi pelatihan dan tipe instruktur", delivery_context_paragraph, "", location_table, "", location_pie_line, "", instructor_type_table, "", instructor_pie_line, "",
            "### Cakupan sumber dan kanal", source_paragraph, "", *source_lines,
        ])

    def _diagnostic_markdown(self, timeframe_df, context):
        if timeframe_df.empty: return "## 2.1 Akar Masalah Utama dan Pain Point Dominan\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        theme_hits = self._theme_hits(timeframe_df)
        theme_lookup = {theme["id"]: theme for theme in theme_hits}
        prioritized_theme_rows = context["score_metrics"]["theme_rows"]
        prioritized_negative_ids = [item["theme_id"] for item in prioritized_theme_rows if theme_lookup.get(item["theme_id"], {}).get("negative_hits", 0) > 0][:3]
        negative_themes = [theme_lookup[theme_id] for theme_id in prioritized_negative_ids]
        if not negative_themes: negative_themes = [theme for theme in theme_hits if theme["negative_hits"] > 0][:3]
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
            if theme["positive_hits"] <= 0: continue
            strongest_services = self._series_counts(theme["matched_df"]["Layanan"], limit=2)
            positive_lines.append(f"- {theme['label']}: {theme['positive_hits']} sinyal positif. Paling banyak muncul pada layanan {', '.join(strongest_services.index.tolist()) or 'belum terpetakan'}.")
        if not positive_lines: positive_lines = ["- Belum ada kekuatan yang cukup konsisten untuk dikonfirmasi pada periode ini."]

        negative_quotes = self._quote_lines(timeframe_df[timeframe_df["Sentiment Label"] == "negative"], limit=3)
        positive_quotes = self._quote_lines(timeframe_df[timeframe_df["Sentiment Label"] == "positive"], limit=2)

        service_risks = self._group_risk(timeframe_df, "Layanan", limit=5)
        location_risks = self._group_risk(timeframe_df, "Lokasi", limit=3)
        instructor_risks = self._group_risk(timeframe_df, "Tipe Instruktur", limit=3)
        process_gap_lines = [f"- {item['label']}: rata-rata rating {item['average_rating']}, proporsi negatif {item['negative_ratio']}%, volume {item['volume']}." for item in service_risks] or ["- Belum ada gap proses yang dapat dipetakan."]
        top_issue = negative_themes[0] if negative_themes else None
        top_strength = next((theme for theme in positive_themes if theme["positive_hits"] > 0 and (not top_issue or theme["id"] != top_issue["id"])), None) or next((theme for theme in positive_themes if theme["positive_hits"] > 0), None)
        dominant_journey = context["dominant_journey"]

        if top_issue and top_strength and top_issue["id"] == top_strength["id"]: strength_context = f"Menariknya, tema {top_strength['label']} muncul sebagai area yang terpolarisasi: sebagian pelanggan menilai sangat baik, sementara sebagian lain masih mengalami hambatan."
        elif top_strength: strength_context = f"Di sisi lain, kekuatan yang paling konsisten terlihat pada {top_strength['label']}."
        else: strength_context = "Kekuatan layanan belum muncul secara cukup konsisten untuk dijadikan diferensiasi yang kuat."

        diagnostic_intro = f"Analisis diagnostik bertujuan menjawab mengapa pola feedback pada periode ini muncul. Karena laporan dibaca dari sudut pandang {context['score_profile']['label']}, perhatian diagnosis terutama diarahkan ke {context['score_profile']['narrative_focus']}. {'Tema keluhan paling dominan saat ini adalah ' + top_issue['label'] + ', yang berulang pada beberapa komentar pelanggan.' if top_issue else 'Belum ada tema keluhan yang sangat dominan, sehingga pola masalah masih relatif tersebar.'} {strength_context}"
        journey_diagnostic = f"Jika dibaca menurut customer journey, titik gesekan yang paling terasa saat ini berada pada tahap {dominant_journey['stage_label']} dengan rating rata-rata {dominant_journey['average_rating']} dan porsi sentimen negatif {dominant_journey['negative_share']}%." if dominant_journey else "Pemetaan customer journey belum menunjukkan titik gesekan yang dominan."
        
        root_cause_table_rows = [[theme["label"], theme["negative_hits"], ", ".join(self._series_counts(theme["matched_df"]["Layanan"], limit=2).index.tolist()) or "Belum terpetakan", ", ".join(self._series_counts(theme["matched_df"]["Tipe Stakeholder"], limit=2).index.tolist()) or "Belum terpetakan"] for theme in negative_themes]
        root_cause_table = self._markdown_table(["Tema Prioritas", "Sinyal Negatif", "Layanan Dominan", "Segmen Dominan"], root_cause_table_rows)
        strength_table_rows = [[theme["label"], theme["positive_hits"], ", ".join(self._series_counts(theme["matched_df"]["Layanan"], limit=2).index.tolist()) or "Belum terpetakan"] for theme in positive_themes if theme["positive_hits"] > 0]
        strength_table = self._markdown_table(["Kekuatan", "Sinyal Positif", "Layanan Dominan"], strength_table_rows)
        service_risk_table = self._markdown_table(["Layanan", "Rata-rata Rating", "Proporsi Negatif", "Volume", "Skor Risiko"], [[item["label"], item["average_rating"], f"{item['negative_ratio']}%", item["volume"], item["risk_score"]] for item in service_risks])
        location_instructor_table = self._markdown_table(["Area Analisis", "Label", "Rata-rata Rating", "Proporsi Negatif", "Volume"], [[ "Lokasi", item["label"], item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in location_risks] + [["Tipe Instruktur", item["label"], item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in instructor_risks])
        operational_context = f"Dari sisi lokasi dan model instruktur, area yang perlu dicermati lebih dekat adalah {location_risks[0]['label'] if location_risks else 'lokasi yang belum terpetakan'} serta komposisi instruktur {instructor_risks[0]['label'] if instructor_risks else 'yang belum terpetakan'}. Pembacaan ini membantu membedakan apakah masalah lebih banyak terkait kesiapan tempat, model pengajar, atau memang tema layanan itu sendiri."

        return "\n".join([
            "## 2.1 Akar Masalah Utama dan Pain Point Dominan", diagnostic_intro, "", "Pembacaan akar masalah dilakukan dengan melihat pengulangan tema, dampaknya pada layanan, dan segmen pelanggan yang paling sering menyinggung isu serupa. Dengan pendekatan ini, tim manajemen dapat membedakan antara keluhan yang bersifat insidental dan keluhan yang sudah layak dibaca sebagai pola struktural.", "",
            root_cause_table, "", *negative_lines, "", "## 2.2 Kekuatan yang Konsisten dan Area yang Perlu Dijaga", "Selain keluhan, periode ini juga memperlihatkan area yang secara berulang diapresiasi oleh pelanggan. Bagian ini penting karena kekuatan yang konsisten dapat dijadikan acuan untuk standardisasi layanan, replikasi praktik baik, dan bahan komunikasi nilai kepada klien.", "",
            strength_table, "", *positive_lines, "", "## 2.3 Bukti Verbatim, Kesenjangan Proses, dan Segmentasi Masalah", "Bukti verbatim di bawah ini digunakan untuk menjaga agar interpretasi manajerial tetap berpijak pada suara pelanggan. Ringkasan kesenjangan proses membantu menerjemahkan komentar individual ke dalam area operasional yang dapat ditindaklanjuti.", "",
            journey_diagnostic, "", "### Kutipan keluhan representatif", *negative_quotes, "### Kutipan apresiasi representatif", *positive_quotes, "### Kesenjangan proses yang paling terlihat", service_risk_table, "", *process_gap_lines, "",
            "### Konteks lokasi pelatihan dan tipe instruktur", operational_context, "", location_instructor_table,
        ])

    def _predictive_markdown(self, timeframe_df, macro_trends, context):
        if timeframe_df.empty: return "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

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
            if item["negative_share"] >= 25: journey_lines.append(f"- Tahap {item['stage_label']} diperkirakan tetap menjadi titik gesekan utama karena porsi sentimen negatif masih {item['negative_share']}%.")
            elif item["positive_share"] >= 60: journey_lines.append(f"- Tahap {item['stage_label']} cenderung tetap menjadi area yang lebih kuat karena porsi sentimen positif mencapai {item['positive_share']}%.")
            else: journey_lines.append(f"- Tahap {item['stage_label']} diperkirakan relatif stabil, tetapi perlu dipantau karena sentimennya masih bercampur.")
        if not journey_lines: journey_lines = ["- Belum ada pembacaan customer journey yang cukup kuat untuk dijadikan proyeksi."]

        osint_signals = self._extract_osint_signals(macro_trends, limit=4)
        deep_insight = self._extract_deep_insight(macro_trends)
        osint_lines = []
        if deep_insight: osint_lines.append(f"- {deep_insight}")
        osint_lines.extend([f"- {signal['title']} ({signal['source']}, {signal['date']}): {signal['snippet']}" for signal in osint_signals])
        if not osint_lines: osint_lines = ["- Tren eksternal belum tersedia; prediksi saat ini sepenuhnya didasarkan pada data internal."]

        top_service_risk = service_risks[0] if service_risks else None
        top_segment_risk = stakeholder_risks[0] if stakeholder_risks else None
        predictive_intro = f"Analisis prediktif membaca risiko yang kemungkinan berkembang apabila pola feedback saat ini berlanjut dalam jangka pendek. {self._projection_sentence(context)} {'Layanan yang paling layak diprioritaskan untuk pengawasan adalah ' + top_service_risk['label'] + '.' if top_service_risk else 'Belum ada layanan dengan pola risiko yang cukup kuat untuk diprioritaskan.'} {'Segmen yang paling perlu dipantau adalah ' + top_segment_risk['label'] + '.' if top_segment_risk else 'Belum ada segmen dengan paparan risiko yang dominan.'}"
        score_projection_table = self._markdown_table(["Score Engine", "Nilai Saat Ini", "Arah Proyeksi", "Nilai Proyeksi", "Horizon", "Estimasi Waktu"], [[context["score_profile"]["label"], score_metrics["current_score"], score_metrics["direction"].title(), score_metrics["projected_score"], context["horizon_text"], self._forecast_calendar_reference(context["timeframe"])]])
        component_breakdown = score_metrics.get("component_breakdown", [])
        experience_formula = self._experience_formula_details(context)
        component_table = self._markdown_table(
            ["Komponen", "Bobot", "Skor Saat Ini", "Skor Proyeksi"],
            [
                [item["label"], f"{round(item['weight'] * 100, 1)}%", item["current_score"], item["projected_score"]]
                for item in component_breakdown
            ],
        )
        formula_table = self._markdown_table(
            ["Parameter", "Perhitungan"],
            [
                ["Komposisi Bobot", experience_formula["weight_summary"]],
                [
                    "Rumus Skor Saat Ini",
                    f"{experience_formula['current_formula']} = {experience_formula['current_calc']} (dibulatkan menjadi {score_metrics['current_score']})",
                ],
                [
                    "Rumus Skor Proyeksi",
                    f"{experience_formula['projected_formula']} = {experience_formula['projected_calc']} (dibulatkan menjadi {score_metrics['projected_score']})",
                ],
            ],
        ) if experience_formula else ""
        service_risk_table = self._markdown_table(["Layanan", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif", "Volume"], [[item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in service_risks])
        stakeholder_risk_table = self._markdown_table(["Segmen", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif", "Volume"], [[item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%", item["volume"]] for item in stakeholder_risks])
        journey_projection_table = self._markdown_table(["Tahap Customer Journey", "Rating Rata-rata", "Negatif", "Positif", "Tema Dominan"], [[item["stage_label"], item["average_rating"], f"{item['negative_share']}%", f"{item['positive_share']}%", item["dominant_theme"]] for item in journey_rows])
        operational_projection_table = self._markdown_table(["Area Operasional", "Label", "Level Risiko", "Rata-rata Rating", "Proporsi Negatif"], [["Lokasi", item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%"] for item in location_risks] + [["Tipe Instruktur", item["label"], self._risk_severity(item["risk_score"]).title(), item["average_rating"], f"{item['negative_ratio']}%"] for item in instructor_risks])
        projection_chart_line = f"[[CHART: Perbandingan Score Saat Ini vs Proyeksi | Skor | Saat Ini,{score_metrics['current_score']}; Proyeksi,{score_metrics['projected_score']}]]"
        osint_table = self._markdown_table(["Sinyal Eksternal", "Sumber", "Tanggal"], [[signal["title"], signal["source"], signal["date"]] for signal in osint_signals])

        return "\n".join([
            "## 3.1 Risiko Jangka Pendek Jika Pola Saat Ini Berlanjut", predictive_intro, "", "Prediksi pada dokumen ini tidak dimaksudkan sebagai forecast statistik jangka panjang, melainkan sebagai early warning berbasis pola rating, proporsi sentimen negatif, dan konsentrasi volume feedback. Dengan pendekatan ini, manajemen dapat lebih cepat memutuskan layanan mana yang perlu ditangani lebih dahulu.", "",
            score_projection_table, "", component_table if component_breakdown else "", "",
            "### Penjelasan Perhitungan Experience Index" if experience_formula else "", f"Experience Index pada laporan ini dihitung dari bobot komponen: {experience_formula['weight_summary']}." if experience_formula else "", "", formula_table if experience_formula else "", "",
            projection_chart_line, "", service_risk_table, "", *risk_lines, "", "## 3.2 Prediksi Segmen dan Layanan yang Paling Rentan", "Selain layanan, pemantauan juga perlu diarahkan pada segmen pelanggan yang memperlihatkan kombinasi antara volume feedback tinggi dan kualitas pengalaman yang menurun. Segmen seperti ini biasanya lebih cepat mempengaruhi reputasi, retensi, dan peluang repeat engagement.", "",
            stakeholder_risk_table, "", *segment_lines, "", "### Pembacaan customer journey ke depan", journey_projection_table, "", *journey_lines, "", "### Area operasional yang perlu diawasi", operational_projection_table, "", *operational_lines, "",
            "## 3.3 Tren Eksternal yang Berpotensi Memperbesar Risiko", "Sinyal eksternal digunakan sebagai benchmark untuk membaca apakah tantangan yang muncul berasal murni dari kondisi internal atau juga diperkuat oleh perubahan ekspektasi pasar. Bila tren eksternal bergerak ke arah yang sama dengan keluhan pelanggan internal, maka urgensi intervensi meningkat.", "",
            osint_table, "", *osint_lines[:6],
        ])

    def _prescriptive_markdown(self, timeframe_df, context):
        if timeframe_df.empty: return "## 4.1 Intervensi Prioritas 30 Hari\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        theme_hits = {theme["id"]: theme for theme in self._theme_hits(timeframe_df)}
        prioritized_actions, prioritized_rows = [], []
        for score_theme in context["score_metrics"]["theme_rows"]:
            theme = theme_hits.get(score_theme["theme_id"])
            if not theme or theme["negative_hits"] <= 0: continue
            action_index = len(prioritized_actions) + 1
            prioritized_actions.append(f"{action_index}. {theme['label']}: {theme['prescription']}")
            prioritized_rows.append([action_index, theme["label"], theme["prescription"], self._theme_owner(theme["id"]), self._theme_outcome(theme["id"])])
            if len(prioritized_actions) >= 4: break

        if not prioritized_actions:
            prioritized_actions = ["1. Pertahankan monitoring mingguan karena belum ada pain point dominan yang membutuhkan intervensi besar."]
            prioritized_rows = [[1, "Monitoring berkala", "Pertahankan pemantauan mingguan dan lakukan review tren secara berkala.", "Quality Assurance / CX", "Risiko laten tetap termonitor meskipun belum ada isu dominan."]]

        governance_actions = ["1. Wajibkan field sumber feedback, kanal, stakeholder, layanan, tanggal, dan rating pada setiap record yang masuk.", "2. Satukan kontrak data antar sistem supaya analisis lintas sumber tetap konsisten dan dapat diaudit.", "3. Tetapkan SLA respon dan eskalasi untuk feedback negatif berprioritas tinggi."]
        roadmap_actions = ["1. Minggu 1: validasi kualitas data, pemetaan owner layanan, dan review pain point dominan.", "2. Minggu 2: jalankan quick wins pada layanan berisiko tertinggi serta aktifkan dashboard monitoring.", "3. Minggu 3-4: evaluasi dampak perbaikan, tutup feedback loop ke stakeholder, dan siapkan iterasi berikutnya.", "[[FLOW: Kumpulkan Feedback Multi-Sumber -> Normalisasi dan Audit Data -> Diagnosa Prioritas -> Jalankan Intervensi -> Evaluasi Dampak]]"]
        action_matrix = self._markdown_table(["Prioritas", "Fokus", "Tindakan", "Owner Utama", "Hasil yang Diharapkan"], prioritized_rows)
        roadmap_table = self._markdown_table(["Tahap", "Fokus Kerja", "Output yang Diharapkan"], [["Minggu 1", "Validasi kualitas data dan pemetaan owner layanan", "Daftar isu prioritas dan penanggung jawab yang disepakati."], ["Minggu 2", "Eksekusi quick wins pada layanan berisiko tertinggi", "Perbaikan cepat berjalan dan dashboard monitoring aktif."], ["Minggu 3-4", "Evaluasi dampak, penutupan feedback loop, dan iterasi", "Status dampak awal terdokumentasi dan rencana lanjutan tersusun."]])
        prescriptive_intro = f"Bagian preskriptif menerjemahkan temuan sebelumnya ke dalam tindakan yang dapat dibahas dan diputuskan dalam forum internal. Urutan prioritas disusun berdasarkan intensitas sinyal negatif, potensi dampak ke pengalaman pelanggan, dan kebutuhan koordinasi lintas fungsi dari sudut pandang {context['score_profile']['label']}."

        return "\n".join([
            "## 4.1 Intervensi Prioritas 30 Hari", prescriptive_intro, "", action_matrix, "", *prioritized_actions, "",
            "## 4.2 Penguatan Tata Kelola Feedback dan Eskalasi", "Selain quick wins layanan, perusahaan juga perlu memperkuat tata kelola feedback agar keputusan perbaikan berikutnya tidak selalu dimulai dari data yang parsial. Penguatan tata kelola akan menentukan kualitas diagnosis, kecepatan eskalasi, dan akuntabilitas tindak lanjut.", "", *governance_actions, "",
            "## 4.3 Rencana Tindak Lanjut Lintas Fungsi", "Rencana tindak lanjut di bawah ini disusun agar forum internal tidak berhenti pada pembacaan laporan, tetapi langsung bergerak ke tahap eksekusi. Timeline dapat disesuaikan, namun disiplin implementasi antar fungsi tetap menjadi faktor penentu keberhasilan.", "", roadmap_table, "", *roadmap_actions,
        ])

    def _implementation_readiness_markdown(self, timeframe_df, timeframe, notes, macro_trends, context):
        if timeframe_df.empty: return "## 5.1 Prioritas Sasaran Bisnis\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        total_rows = len(timeframe_df)
        avg_rating = timeframe_df["Rating Numeric"].mean()
        positive_count = int((timeframe_df["Sentiment Label"] == "positive").sum())
        negative_count = int((timeframe_df["Sentiment Label"] == "negative").sum())
        positive_share = self._safe_percentage(positive_count, total_rows)
        negative_share = self._safe_percentage(negative_count, total_rows)
        governance = self._governance_summary(timeframe_df)
        theme_hits = self._theme_hits(timeframe_df)
        service_risks = self._group_risk(timeframe_df, "Layanan", limit=3)
        stakeholder_risks = self._group_risk(timeframe_df, "Tipe Stakeholder", limit=3)
        top_service = service_risks[0] if service_risks else None
        top_segment = stakeholder_risks[0] if stakeholder_risks else None
        top_issue = next((theme for theme in theme_hits if theme["negative_hits"] > 0), None)
        top_strength = next((theme for theme in theme_hits if theme["positive_hits"] > 0 and (not top_issue or theme["id"] != top_issue["id"])), None) or next((theme for theme in theme_hits if theme["positive_hits"] > 0), None)
        
        osint_signals = self._extract_osint_signals(macro_trends, limit=2)
        deep_insight = self._extract_deep_insight(macro_trends)
        
        focus_text = notes.strip().rstrip(".!?") if notes and notes.strip() else "Tidak ada fokus tambahan dari pengguna"
        top_service_name = top_service["label"] if top_service else self._primary_label(self._series_counts(timeframe_df["Layanan"], limit=1), "layanan prioritas")
        top_segment_name = top_segment["label"] if top_segment else self._primary_label(self._series_counts(timeframe_df["Tipe Stakeholder"], limit=1), "segmen utama")

        business_score = min(100, 50 + min(total_rows, 10) * 3 + (10 if top_service else 0) + (10 if top_issue else 0))
        data_score = min(100, int(governance["completeness_pct"] * 0.45) + min(governance["source_count"], 3) * 10 + min(governance["channel_count"], 3) * 10 + (10 if governance["source_count"] >= 2 else 0) + (10 if governance["channel_count"] >= 1 else 0))
        architecture_score = min(100, 35 + min(governance["source_count"], 3) * 8 + min(governance["channel_count"], 3) * 10 + (10 if governance["source_count"] >= 2 else 0) + (10 if total_rows >= 10 else 5 if total_rows >= 5 else 0))
        people_score = max(35, min(100, 60 + (10 if top_strength else 0) - (10 if top_issue and top_issue["id"] in {"instructor", "communication"} else 0) - (5 if top_issue and top_issue["id"] in {"responsiveness", "schedule"} else 0)))
        governance_score = min(100, int(governance["completeness_pct"] * 0.35) + min(governance["source_count"], 3) * 8 + min(governance["channel_count"], 3) * 12 + (10 if governance["source_count"] >= 2 else 0) + (10 if governance["channel_count"] >= 1 else 0))
        culture_score = max(35, min(100, 45 + int(positive_share * 0.4) - int(negative_share * 0.3) + (10 if total_rows >= 5 else 0)))

        pillar_map = {
            "business_use_case": {
                "score": business_score,
                "reading": f"Prioritas implementasi yang paling konkret saat ini adalah memperkuat tata kelola feedback untuk mendeteksi lebih dini risiko pada layanan {top_service_name} dan memantau pengalaman segmen {top_segment_name}. Dengan rata-rata rating {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dan sinyal negatif {negative_share}%, laporan ini sudah memiliki dasar yang cukup jelas untuk diterjemahkan ke keputusan bisnis. Fokus use case saat ini dibaca pada {context['scope_text']}.",
                "implication": "Inisiatif ini sebaiknya tidak diposisikan sebagai eksperimen AI yang abstrak, melainkan sebagai use case terukur untuk mempercepat diagnosis masalah, prioritisasi perbaikan, dan evaluasi dampak manajerial.",
                "actions": [f"Tentukan satu sasaran bisnis utama, misalnya menurunkan sinyal negatif pada {top_service_name} atau mempercepat penutupan isu pelanggan.", "Sepakati KPI pilot 30 hari yang mudah diukur, misalnya waktu respons, penurunan keluhan berulang, atau peningkatan kepuasan.", f"Gunakan fokus rapat pada area berikut: {focus_text}."],
            },
            "data_model_foundation": {
                "score": data_score,
                "reading": f"Fondasi data saat ini cukup untuk memulai pilot karena kelengkapan field inti mencapai {governance['completeness_pct']}%. Namun cakupan masih berasal dari {governance['source_count']} sumber dan {governance['channel_count']} kanal yang terpetakan, sehingga representativitas lintas kanal belum sepenuhnya kuat.",
                "implication": "Pilot dapat berjalan sekarang, tetapi scale-up akan sulit bila kontrak data, owner data, dan standar mandatory field belum disepakati bersama.",
                "actions": ["Tetapkan owner data untuk setiap sumber feedback dan definisikan field yang wajib terisi.", "Pastikan kanal, tanggal, stakeholder, layanan, rating, dan komentar selalu tercatat secara konsisten.", "Gunakan model analitik saat ini sebagai baseline, lalu perluas sumber data secara bertahap setelah kualitas data stabil."],
            },
            "infrastructure_architecture": {
                "score": architecture_score,
                "reading": "Implementasi untuk use case ini tidak harus dimulai dari arsitektur yang mahal. Kebutuhan dekatnya adalah arsitektur yang aman, dapat dibagikan secara internal, mendukung read-only integration, dan cukup mudah diperluas ketika sumber data bertambah.",
                "implication": "Keputusan cloud, on-prem, atau hybrid sebaiknya mengikuti kebutuhan compliance perusahaan. Untuk tahap saat ini, prioritasnya adalah kestabilan deployment internal, logging, health check, dan jalur ingest data yang dapat diaudit.",
                "actions": ["Mulai dari shared internal deployment yang stabil sebelum memikirkan scale penuh.", "Gunakan akses API read-only untuk data internal dan pisahkan konfigurasi pilot dari produksi.", "Siapkan jalur scale bertahap ke cloud, on-prem, atau hybrid sesuai regulasi dan kebijakan keamanan data."],
            },
            "people_capability": {
                "score": people_score,
                "reading": f"Temuan periode ini menunjukkan bahwa inisiatif ini tidak dapat menjadi urusan IT saja. {'Tema utama yang muncul adalah ' + top_issue['label'] + ', sehingga interpretasi bisnis perlu melibatkan owner layanan.' if top_issue else 'Interpretasi insight tetap membutuhkan owner layanan dan pihak operasional.'} {'Kekuatan yang layak dijaga terlihat pada ' + top_strength['label'] + '.' if top_strength else 'Belum ada kekuatan dominan yang cukup untuk dijadikan patokan lintas tim.'}",
                "implication": "Agar pilot menghasilkan keputusan nyata, perusahaan perlu membedakan peran tim teknis, QA/CX, business translator, owner layanan, dan eksekutor perbaikan di lapangan.",
                "actions": ["Tunjuk business owner yang bertanggung jawab atas use case dan outcome pilot.", "Pastikan QA/CX, owner layanan, dan operasional ikut mereview laporan, bukan hanya tim teknis.", "Siapkan satu business translator yang menerjemahkan insight teknis menjadi keputusan manajerial."],
            },
            "governance": {
                "score": governance_score,
                "reading": "Kontrol risiko dan tata kelola menjadi area penting karena laporan ini memadukan data internal dan konteks eksternal. Internal data harus tetap menjadi sumber kebenaran untuk fakta operasional, sedangkan OSINT dipakai hanya sebagai benchmark dan konteks pasar.",
                "implication": "Organisasi perlu mendefinisikan dengan tegas data apa yang boleh dipakai, siapa yang berhak mengaksesnya, kapan rekomendasi boleh dijalankan, dan kapan hasil AI harus dihentikan atau dikoreksi secara manual.",
                "actions": ["Tetapkan SOP penggunaan data internal vs OSINT agar tidak terjadi pencampuran fakta dan konteks publik.", "Buat review cadence dan approval gate untuk insight berisiko tinggi sebelum menjadi keputusan resmi.", "Dokumentasikan risk control, quality check, dan batas penggunaan AI pada forum evaluasi internal."],
            },
            "culture": {
                "score": culture_score,
                "reading": f"Budaya organisasi untuk inisiatif ini sebaiknya dibangun dengan semangat mencoba secara terstruktur. Komposisi sentimen positif {positive_share}% menunjukkan ada modal kepercayaan yang cukup untuk memulai, sementara keberadaan sinyal negatif tetap penting sebagai bahan belajar dan perbaikan.",
                "implication": "Keberhasilan tahap awal tidak harus berarti sistem langsung sempurna. Yang lebih penting adalah organisasi punya kebiasaan mencoba, mengevaluasi, mengambil pelajaran, dan memutuskan langkah berikutnya secara disiplin.",
                "actions": ["Posisikan pilot sebagai ruang belajar terstruktur, bukan proyek yang harus sempurna sejak hari pertama.", "Dokumentasikan apa yang berhasil, apa yang belum, dan keputusan apa yang diambil setelah setiap periode evaluasi.", "Gunakan laporan ini sebagai alat diskusi lintas fungsi agar AI adoption menjadi perubahan cara kerja, bukan hanya tambahan tools."],
            },
        }

        summary_rows, pillar_sections = [], []
        for pillar in ADOPTION_READINESS_PILLARS:
            pillar_data = pillar_map[pillar["id"]]
            status = self._readiness_label(pillar_data["score"])
            summary_rows.append([" ".join(pillar["title"].split(" ")[1:]), status, pillar_data["reading"], pillar["guiding_question"]])
            pillar_sections.extend([f"## {pillar['title']}", pillar_data["reading"], "", f"Status kesiapan saat ini: **{status}**.", "", f"Pertanyaan pemandu: {pillar['guiding_question']}", "", "### Implikasi untuk Pengambilan Keputusan", pillar_data["implication"], "", "### Aksi Prioritas", *[f"- {action}" for action in pillar_data["actions"]], ""])

        if deep_insight and osint_signals: osint_note = f"{deep_insight} Sinyal eksternal lainnya yang relevan antara lain {osint_signals[0]['title']} dari {osint_signals[0]['source']}."
        elif osint_signals: osint_note = f"Sinyal eksternal yang paling relevan saat ini antara lain {osint_signals[0]['title']} dari {osint_signals[0]['source']}."
        else: osint_note = "Sinyal eksternal belum tersedia, sehingga pembacaan kesiapan implementasi terutama bersandar pada data internal."
        summary_table = self._markdown_table(["Area", "Status", "Pembacaan Saat Ini", "Pertanyaan Diskusi"], summary_rows)

        return "\n".join([
            "Bagian ini menerjemahkan hasil feedback intelligence ke dalam pertimbangan implementasi dan penguatan organisasi agar perusahaan tidak berhenti pada insight, tetapi bergerak menuju eksekusi yang terstruktur. Prinsipnya adalah memulai dari use case yang nyata, membangun fondasi secara bertahap, lalu belajar secara disiplin dari pilot yang dijalankan.", "",
            f"Untuk periode {timeframe}, pertimbangan implementasi perlu dilihat bersama konteks berikut: {osint_note} Analisis saat ini dibaca menggunakan {context['score_profile']['label']} dengan fokus pada {context['score_profile']['narrative_focus']}. Dengan demikian, forum internal dapat menilai bukan hanya apa yang harus diperbaiki, tetapi juga seberapa siap organisasi untuk menjalankan inisiatif ini secara lebih sistematis.", "",
            summary_table, "", *pillar_sections,
        ])

    def build_report_sections(self, timeframe, notes, macro_trends, sentiment="all", segment="all", score_engine=DEFAULT_SCORE_ENGINE):
        timeframe_df = self._filter_view(timeframe, sentiment=sentiment, segment=segment)
        context = self._build_analysis_context(timeframe_df, timeframe, sentiment, segment, score_engine)
        section_map = {
            "cx_chap_1": self._descriptive_markdown(timeframe_df, timeframe, notes, context),
            "cx_chap_2": self._diagnostic_markdown(timeframe_df, context),
            "cx_chap_3": self._predictive_markdown(timeframe_df, macro_trends, context),
            "cx_chap_4": self._prescriptive_markdown(timeframe_df, context),
            "cx_chap_5": self._implementation_readiness_markdown(timeframe_df, timeframe, notes, macro_trends, context),
        }
        return [{"id": chapter["id"], "title": chapter["title"], "content": section_map.get(chapter["id"], "")} for chapter in CX_SENTIMENT_STRUCTURE]

    def build_executive_snapshot(self, timeframe, notes="", sentiment="all", segment="all", score_engine=DEFAULT_SCORE_ENGINE):
        timeframe_df = self._filter_view(timeframe, sentiment=sentiment, segment=segment)
        if timeframe_df.empty: return "## Ringkasan Eksekutif\n- Tidak ada data internal yang cukup untuk menyusun snapshot eksekutif pada kombinasi filter yang dipilih.\n"

        context = self._build_analysis_context(timeframe_df, timeframe, sentiment, segment, score_engine)
        total_rows = len(timeframe_df)
        avg_rating = timeframe_df["Rating Numeric"].mean()
        negative_count = int((timeframe_df["Sentiment Label"] == "negative").sum())
        negative_share = self._safe_percentage(negative_count, total_rows)
        top_service = self._series_counts(timeframe_df["Layanan"], limit=1)
        top_stakeholder = self._series_counts(timeframe_df["Tipe Stakeholder"], limit=1)
        top_risk = self._group_risk(timeframe_df, "Layanan", limit=1)
        governance = self._governance_summary(timeframe_df)
        top_issue = next((theme for theme in self._theme_hits(timeframe_df) if theme["negative_hits"] > 0), None)
        focus_text = notes.strip() if notes and notes.strip() else "Tidak ada fokus tambahan dari pengguna."
        dominant_journey, score_metrics = context["dominant_journey"], context["score_metrics"]
        experience_formula = self._experience_formula_details(context)
        top_location = self._series_counts_for_column(timeframe_df, "Lokasi", limit=1)
        top_instructor_type = self._series_counts_for_column(timeframe_df, "Tipe Instruktur", limit=1)

        risk_statement = f"- Risiko teratas saat ini ada pada layanan {top_risk[0]['label']} dengan proporsi sinyal negatif {top_risk[0]['negative_ratio']}%." if top_risk else "- Belum ada layanan dengan risiko dominan yang teridentifikasi."
        executive_intro = f"Laporan ini merangkum kondisi pengalaman pelanggan untuk periode {timeframe} berdasarkan {total_rows} feedback tervalidasi. Analisis saat ini dibaca pada {context['scope_text']} dengan fokus pada {context['score_profile']['narrative_focus']}. Secara umum, rata-rata rating berada pada level {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5, yang menunjukkan kualitas layanan {self._rating_assessment(avg_rating)}. Proporsi sentimen negatif tercatat sebesar {negative_share}% ({negative_count} feedback), sehingga kondisi ini {self._negative_share_assessment(negative_share)}."
        meeting_context = f"Untuk kebutuhan rapat internal, perhatian utama sebaiknya diarahkan pada layanan {top_risk[0]['label'] if top_risk else self._primary_label(top_service, 'yang memiliki volume feedback terbesar')} serta pada isu {top_issue['label'] if top_issue else 'konsistensi kualitas layanan'}. {self._projection_sentence(context)} Fokus tambahan yang diminta pengguna: {focus_text}"
        snapshot_rows = [
            ["Total feedback dianalisis", f"{total_rows} record"],
            ["Cakupan analisis", context["scope_text"]],
            [context["score_profile"]["label"], f"{score_metrics['current_score']} / 100"],
            ["Rata-rata rating", f"{round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5"],
            ["Proporsi sentimen negatif", f"{negative_share}%"],
            ["Layanan dengan volume terbesar", self._primary_label(top_service, "Belum terpetakan")],
            ["Segmen dengan volume terbesar", self._primary_label(top_stakeholder, "Belum terpetakan")],
            ["Lokasi pelatihan dominan", self._primary_label(top_location, "Belum terpetakan")],
            ["Tipe instruktur dominan", self._primary_label(top_instructor_type, "Belum terpetakan")],
            ["Kelengkapan field inti", f"{governance['completeness_pct']}%"],
        ]
        if experience_formula:
            snapshot_rows.extend(
                [
                    ["Formula Experience Index", experience_formula["weight_summary"]],
                    ["Perhitungan Skor Saat Ini", f"{experience_formula['current_formula']} = {experience_formula['current_calc']} (dibulatkan menjadi {score_metrics['current_score']})"],
                ]
            )
        snapshot_table = self._markdown_table(["Indikator Kunci", "Nilai"], snapshot_rows)
        meeting_agenda = [f"- Apakah layanan {top_risk[0]['label']} memerlukan intervensi prioritas lintas fungsi pada 30 hari ke depan?" if top_risk else "- Apakah perusahaan perlu memperluas pengumpulan feedback agar risiko layanan lebih mudah dibaca?", f"- Bagaimana tindak lanjut yang paling tepat untuk tema {top_issue['label']} agar tidak berkembang menjadi keluhan berulang?" if top_issue else "- Kekuatan layanan mana yang paling layak distandardisasi dan direplikasi?", f"- Tahap customer journey mana yang paling perlu dikoreksi lebih dulu, mengingat titik gesekan terbesar saat ini berada pada {dominant_journey['stage_label']}?" if dominant_journey else "- Tahap customer journey mana yang paling perlu dipetakan lebih rinci pada periode berikutnya?", "- Apakah tata kelola sumber, kanal, dan owner tindak lanjut sudah cukup jelas untuk mendukung evaluasi periodik berikutnya?"]

        return "\n".join([
            "## Ringkasan Eksekutif", executive_intro, "", meeting_context, "", snapshot_table, "", "### Agenda Diskusi Prioritas", *meeting_agenda, "",
            "### Poin Utama untuk Pembacaan Cepat", f"- Total feedback yang dianalisis: {total_rows} record.", f"- Rata-rata rating periode ini: {round(avg_rating, 2) if pd.notna(avg_rating) else 0.0} dari 5.", f"- Volume layanan terbesar: {top_service.index[0] if not top_service.empty else 'Belum terpetakan'}.", f"- Segmen dengan volume terbesar: {top_stakeholder.index[0] if not top_stakeholder.empty else 'Belum terpetakan'}.", f"- Lokasi pelatihan dominan: {top_location.index[0] if not top_location.empty else 'Belum terpetakan'}.", f"- Tipe instruktur dominan: {top_instructor_type.index[0] if not top_instructor_type.empty else 'Belum terpetakan'}.", f"- Proporsi sentimen negatif: {negative_share}%.", f"- {context['score_profile']['label']} saat ini: {score_metrics['current_score']} dengan proyeksi {score_metrics['direction']} ke {score_metrics['projected_score']}.", f"- Tahap customer journey yang paling perlu diperhatikan: {dominant_journey['stage_label']}." if dominant_journey else "- Pemetaan customer journey belum menunjukkan titik perhatian yang dominan.", risk_statement, "- Struktur laporan ini disusun untuk mendukung analisis Descriptive, Diagnostic, Predictive, Prescriptive, serta kesiapan implementasi dan penguatan organisasi secara konsisten.",
        ])

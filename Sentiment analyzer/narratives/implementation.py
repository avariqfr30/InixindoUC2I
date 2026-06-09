import pandas as pd

from config import ADOPTION_READINESS_PILLARS
from .base import BaseNarrativeMixin


class ImplementationNarrativeMixin(BaseNarrativeMixin):
    def _implementation_readiness_markdown(self, timeframe_df, timeframe, notes, macro_trends, context, section_context=None):
        timeframe_label = context.get("timeframe_label", timeframe)
        if timeframe_df.empty:
            return "## 5.1 Prioritas Sasaran Bisnis\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        total_rows = len(timeframe_df)
        avg_rating = timeframe_df["Rating Numeric"].mean()
        sentiment_summary = self._sentiment_summary(timeframe_df)
        positive_share = sentiment_summary["positive_share"]
        negative_share = sentiment_summary["issue_share"]
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
        
        section_context = section_context or {}
        focus_text = str(section_context.get("focus_note") or notes or "").strip().rstrip(".!?") or "Tidak ada fokus tambahan dari pengguna"
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

        if deep_insight and osint_signals:
            osint_note = f"Ringkasan tren eksternal: pasar menuntut bukti dampak pelatihan, tindak lanjut yang lebih jelas, dan materi yang mengikuti kebutuhan kompetensi terbaru."
        elif osint_signals:
            osint_note = f"Ringkasan tren eksternal: sinyal publik yang tersedia menekankan kebutuhan konsistensi hasil, pembaruan kompetensi, dan tindak lanjut layanan."
        else:
            osint_note = "Sinyal eksternal belum tersedia, sehingga pembacaan kesiapan implementasi terutama bersandar pada data internal."
        summary_table = self._markdown_table(["Area", "Status", "Pembacaan Saat Ini", "Pertanyaan Diskusi"], summary_rows)

        return "\n".join([
            "Bagian ini menerjemahkan hasil feedback intelligence ke dalam pertimbangan implementasi dan penguatan organisasi agar perusahaan tidak berhenti pada insight, tetapi bergerak menuju eksekusi yang terstruktur. Prinsipnya adalah memulai dari use case yang nyata, membangun fondasi secara bertahap, lalu belajar secara disiplin dari pilot yang dijalankan.", "",
            f"Untuk periode {timeframe_label}, pertimbangan implementasi perlu dilihat bersama konteks berikut: {osint_note} Analisis saat ini dibaca menggunakan {context['score_profile']['label']} dengan fokus pada {context['score_profile']['narrative_focus']}. Dengan demikian, forum internal dapat menilai bukan hanya apa yang harus diperbaiki, tetapi juga seberapa siap organisasi untuk menjalankan inisiatif ini secara lebih sistematis.", "",
            summary_table, "", *pillar_sections,
        ])

import pandas as pd

from .base import BaseNarrativeMixin
from report_factuality import ReportFactRegistry


class PrescriptiveNarrativeMixin(BaseNarrativeMixin):
    @staticmethod
    def _theme_success_indicator(theme_id):
        indicators = {
            "responsiveness": "Waktu respons dan jumlah isu melewati SLA.",
            "schedule": "Keluhan jadwal berulang dan kepatuhan agenda.",
            "facility": "Jumlah gangguan fasilitas saat pelaksanaan.",
            "instructor": "Rating fasilitator dan kritik berulang.",
            "material": "Rating relevansi materi dan kritik versi konten.",
            "communication": "Keluhan koordinasi dan ketepatan pembaruan status.",
            "outcome": "Bukti tindak lanjut dan manfaat pascalayanan.",
        }
        return indicators.get(theme_id, "Perubahan rating, komentar, dan pengulangan isu.")

    def _prescriptive_markdown(self, timeframe_df, context, section_context=None):
        if timeframe_df.empty:
            return "## 4.1 Intervensi Prioritas 30 Hari\nTidak ada feedback internal yang sesuai dengan kombinasi filter yang dipilih untuk periode ini.\n"

        section_context = section_context or {}
        trust_packet = section_context.get("trust_packet") or {}
        evidence_ids = trust_packet.get("theme_evidence_ids") or {}
        theme_hits = {theme["id"]: theme for theme in self._theme_hits(timeframe_df)}
        prioritized_actions, prioritized_rows = [], []
        for score_theme in context["score_metrics"]["theme_rows"]:
            theme = theme_hits.get(score_theme["theme_id"])
            if not theme or theme["negative_hits"] <= 0:
                continue
            action_index = len(prioritized_actions) + 1
            prioritized_actions.append(f"{action_index}. {theme['label']}: {theme['prescription']}")
            evidence_id = evidence_ids.get(theme["id"]) or ReportFactRegistry.theme_fact_id(theme["id"])
            prioritized_rows.append([
                action_index,
                evidence_id,
                theme["label"],
                theme["prescription"],
                self._theme_owner(theme["id"]),
                "30 hari",
                self._theme_success_indicator(theme["id"]),
                self._theme_outcome(theme["id"]),
            ])
            if len(prioritized_actions) >= 4:
                break

        if not prioritized_actions:
            prioritized_actions = ["1. Pertahankan monitoring mingguan karena belum ada pain point dominan yang membutuhkan intervensi besar."]
            prioritized_rows = [[
                1,
                "F-MONITORING",
                "Monitoring berkala",
                "Pertahankan pemantauan mingguan dan lakukan review tren secara berkala.",
                "Quality Assurance / CX",
                "30 hari",
                "Perubahan rating, komentar, dan kemunculan isu baru.",
                "Risiko laten tetap termonitor meskipun belum ada isu dominan.",
            ]]

        governance_actions = ["1. Wajibkan field sumber feedback, kanal, stakeholder, layanan, tanggal, dan rating pada setiap record yang masuk.", "2. Satukan kontrak data antar sistem supaya analisis lintas sumber tetap konsisten dan dapat diaudit.", "3. Tetapkan SLA respon dan eskalasi untuk feedback negatif berprioritas tinggi."]
        roadmap_actions = ["1. Minggu 1: validasi kualitas data, pemetaan owner layanan, dan review pain point dominan.", "2. Minggu 2: jalankan quick wins pada layanan berisiko tertinggi serta aktifkan dashboard monitoring.", "3. Minggu 3-4: evaluasi dampak perbaikan, tutup feedback loop ke stakeholder, dan siapkan iterasi berikutnya.", "[[FLOW: Kumpulkan Feedback Multi-Sumber -> Normalisasi dan Audit Data -> Diagnosa Prioritas -> Jalankan Intervensi -> Evaluasi Dampak]]"]
        action_matrix = self._markdown_table(
            [
                "Prioritas",
                "ID Bukti",
                "Fokus",
                "Tindakan",
                "Penanggung Jawab Utama",
                "Jendela Tinjauan",
                "Indikator Keberhasilan",
                "Hasil yang Diharapkan",
            ],
            prioritized_rows,
        )
        roadmap_table = self._markdown_table(["Tahap", "Fokus Kerja", "Output yang Diharapkan"], [["Minggu 1", "Validasi kualitas data dan pemetaan owner layanan", "Daftar isu prioritas dan penanggung jawab yang disepakati."], ["Minggu 2", "Eksekusi quick wins pada layanan berisiko tertinggi", "Perbaikan cepat berjalan dan dashboard monitoring aktif."], ["Minggu 3-4", "Evaluasi dampak, penutupan feedback loop, dan iterasi", "Status dampak awal terdokumentasi dan rencana lanjutan tersusun."]])
        prescriptive_intro = f"Bagian preskriptif menerjemahkan temuan sebelumnya ke dalam tindakan yang dapat dibahas dan diputuskan dalam forum internal. Urutan prioritas disusun berdasarkan intensitas sinyal negatif, potensi dampak ke pengalaman pelanggan, dan kebutuhan koordinasi lintas fungsi dari sudut pandang {context['score_profile']['label']}."

        return "\n".join([
            "## 4.1 Intervensi Prioritas 30 Hari", prescriptive_intro, "", action_matrix, "", *prioritized_actions, "",
            "## 4.2 Penguatan Tata Kelola Feedback dan Eskalasi", "Selain quick wins layanan, perusahaan juga perlu memperkuat tata kelola feedback agar keputusan perbaikan berikutnya tidak selalu dimulai dari data yang parsial. Penguatan tata kelola akan menentukan kualitas diagnosis, kecepatan eskalasi, dan akuntabilitas tindak lanjut.", "", *governance_actions, "",
            "## 4.3 Rencana Tindak Lanjut Lintas Fungsi", "Rencana tindak lanjut di bawah ini disusun agar forum internal tidak berhenti pada pembacaan laporan, tetapi langsung bergerak ke tahap eksekusi. Timeline dapat disesuaikan, namun disiplin implementasi antar fungsi tetap menjadi faktor penentu keberhasilan.", "", roadmap_table, "", *roadmap_actions,
        ])

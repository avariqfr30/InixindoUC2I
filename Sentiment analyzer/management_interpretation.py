class FeedbackManagementInterpreter:
    """Turn analytics metrics into an Inixindo-style executive interpretation."""

    @staticmethod
    def build(
        avg_rating,
        negative_share,
        top_risk_label,
        issue_label,
        journey_label,
        score_label,
        current_score,
        projected_score,
    ):
        rating_text = round(avg_rating, 2) if avg_rating == avg_rating else 0.0
        if negative_share <= 0:
            return {
                "signal": f"Rating pelanggan {rating_text}/5 tanpa sinyal negatif material pada cakupan ini.",
                "meaning": f"Pengalaman pelanggan sedang cukup kuat, terutama pada {top_risk_label}.",
                "decision": "Manajemen perlu memilih praktik baik yang layak dijadikan standar layanan.",
                "action": f"Dokumentasikan pola positif pada {journey_label} dan replikasi ke layanan lain dalam 30 hari.",
                "confidence": "Cukup kuat - lensa pengalaman mendukung keputusan replikasi praktik baik.",
            }
        return {
            "signal": f"Rating pelanggan {rating_text}/5 dengan sentimen negatif {negative_share}%.",
            "meaning": f"Risiko pengalaman pelanggan terkonsentrasi pada {top_risk_label}, terutama terkait {issue_label}.",
            "decision": "Manajemen perlu memilih satu area intervensi utama agar perbaikan tidak tersebar terlalu luas.",
            "action": f"Tetapkan owner lintas fungsi untuk {top_risk_label} dan jalankan rencana 30 hari pada {journey_label}.",
            "confidence": "Cukup kuat - lensa pengalaman menunjukkan prioritas koreksi yang perlu diuji pada periode berikutnya.",
        }

    @staticmethod
    def to_markdown_table(interpretation):
        row = interpretation or {}
        return "\n".join(
            [
                "| Sinyal | Makna | Keputusan | Aksi | Keyakinan |",
                "| --- | --- | --- | --- | --- |",
                (
                    f"| {row.get('signal', '-')} | {row.get('meaning', '-')} | "
                    f"{row.get('decision', '-')} | {row.get('action', '-')} | {row.get('confidence', '-')} |"
                ),
            ]
        )

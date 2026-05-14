from report_agents import FeedbackProposalTeam


def build_specialist_review_markdown(engine, markdown_table, timeframe, macro_trends, sentiment, segment, score_engine):
    briefing = FeedbackProposalTeam().run(
        engine,
        engine.full_df,
        timeframe,
        macro_trends=macro_trends,
        sentiment=sentiment,
        segment=segment,
        score_engine=score_engine,
    )
    rows = [
        [
            item["role"],
            item["dataset"],
            item["confidence"],
            item["finding"],
            item["implication"],
        ]
        for item in briefing["specialists"]
    ]
    sources = ", ".join(briefing["sources_used"]) if briefing["sources_used"] else "evaluasi layanan"
    ledger_rows = [
        [item["evidence_type"], item["source"], item["detail"]]
        for item in briefing["evidence_ledger"]
    ]
    qa_lines = [f"- {item}" for item in briefing["qa_review"]]
    audit = briefing["audit_trail"]
    contradiction = briefing["contradiction_review"]
    trend = briefing["trend_review"]
    prediction = briefing["prediction_review"]
    audit_rows = [
        ["Generated at", audit["generated_at_utc"]],
        ["Periode", audit["timeframe"]],
        ["Filter", f"sentiment={audit['sentiment']}; segment={audit['segment']}; score_engine={audit['score_engine']}"],
        ["Cakupan data", f"{audit['raw_response_count']} respons mentah; {audit['dimension_count']} dimensi evaluasi"],
        ["Komposisi", f"{audit['rating_response_count']} rating; {audit['text_response_count']} komentar teks"],
        ["Kelengkapan", f"{audit['field_completeness_pct']}% field inti; {audit['source_count']} sumber; {audit['channel_count']} kanal"],
    ]
    return "\n".join([
        "### Review Tim Analis Internal",
        briefing["manager_summary"],
        "",
        f"**Confidence Desk:** {briefing['confidence']}.",
        "",
        markdown_table(
            ["Peran", "Fokus Bukti", "Confidence", "Temuan", "Implikasi"],
            rows,
        ),
        "",
        "### Evidence Ledger",
        markdown_table(
            ["Tipe Bukti", "Sumber", "Catatan"],
            ledger_rows,
        ),
        "",
        "### QA Guardrail",
        *qa_lines,
        "",
        "### Report Audit Trail",
        markdown_table(
            ["Item", "Nilai"],
            audit_rows,
        ),
        "",
        "### Contradiction Check",
        f"Status: **{contradiction['severity']}**. {contradiction['rating_text_alignment']}. "
        f"Rata-rata rating {contradiction['average_rating']}/5, negative rating share {contradiction['negative_rating_share']}%, "
        f"negative text hits {contradiction['negative_text_hits']}, positive text hits {contradiction['positive_text_hits']}.",
        "",
        "### Historical Trend Desk",
        f"Periode pembanding: {trend['comparison_period']}. Rating delta: {trend['rating_delta']}; negative share delta: {trend['negative_share_delta']}%. {trend['reading']}",
        "",
        "### Prediction Boundary",
        f"{prediction['method']} Arah saat ini: {prediction['direction']} dari {prediction['current_score']} menuju {prediction['projected_score']}. "
        f"{prediction['confidence_note']}",
        "",
        f"Jejak bukti yang dipakai: {sources}.",
    ])

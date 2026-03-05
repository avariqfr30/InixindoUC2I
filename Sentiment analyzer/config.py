"""
config.py
---------
Holds all environmental variables, constants, and LLM prompt structures.
Managerial edits to prompts or model settings should happen here.
"""

OLLAMA_HOST = "http://127.0.0.1:11434"
LLM_MODEL = "gpt-oss:120b-cloud"
EMBED_MODEL = "bge-m3:latest"
DB_FILE = "data/db.csv"
COMPANY_NAME = "Inixindo Jogja"

REPORT_STRUCTURE = [
    {
        "id": "chap_1", 
        "title": "BAB I - RINGKASAN EKSEKUTIF & SENTIMEN",
        "instructions": """
        Summarize the overall sentiment based on the feedback provided. 
        Is the sentiment predominantly positive, negative, or neutral? 
        Provide a brief, professional executive summary of the relationship and recent performance.
        Do not use markdown headers for this section, just plain paragraphs.
        """,
        "visual_intent": "sentiment_chart" 
    },
    {
        "id": "chap_2", 
        "title": "BAB II - ANALISIS AKAR MASALAH (ROOT CAUSE)",
        "instructions": """
        Identify the top 2-3 main issues or pain points explicitly mentioned in the feedback.
        For each issue, provide a deep analysis using exactly this markdown format:
        
        ### [Nama Isu/Masalah]
        * **Konteks:** (Jelaskan masalah secara objektif berdasarkan feedback)
        * **Analisis Akar Masalah:** (Mengapa ini terjadi? Deduksi akar masalah teknis/operasional)
        * **Dampak Bisnis:** (Apa risiko jika masalah ini dibiarkan? misal: reputasi, efisiensi)
        """
    },
    {
        "id": "chap_3", 
        "title": "BAB III - REKOMENDASI STRATEGIS & TAKTIS",
        "instructions": """
        Based on the identified issues, provide highly specific, actionable solutions. 
        Do not give generic advice. Use exactly this markdown format to structure your response:
        
        ### Solusi Taktis (Jangka Pendek - < 3 Bulan)
        * **Tindakan:** [Tindakan spesifik yang harus segera dilakukan]
        * **Target:** [Apa yang ingin dicapai dari tindakan ini]
        
        ### Solusi Strategis (Jangka Panjang - > 6 Bulan)
        * **Inisiatif:** [Perubahan sistemik atau perombakan proses]
        * **Pencegahan:** [Bagaimana inisiatif ini mencegah masalah terulang di masa depan]
        """
    }
]
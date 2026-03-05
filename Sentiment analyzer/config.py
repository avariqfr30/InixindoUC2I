"""
config.py
---------
Contains environment variables, application settings, and LLM prompt templates.
The prompts are engineered to force the LLM into a highly critical, deeply analytical, 
and verbose diagnostic mode, while using nested lists for readability.
"""

OLLAMA_URL = "http://127.0.0.1:11434"
AI_MODEL = "gpt-oss:120b-cloud"
EMBEDDING_MODEL = "bge-m3:latest"
DATA_SOURCE = "data/db.csv"
ORG_NAME = "Inixindo Jogja"

REPORT_SECTIONS = [
    {
        "id": "nlp_classification", 
        "title": "BAB I - KLASIFIKASI NLP & DIAGNOSA SENTIMEN KRITIS",
        "instructions": """
        Anda adalah Auditor Internal dan HR Diagnostician senior yang kejam dan sangat teliti untuk perusahaan.
        Tugas Anda adalah menginterogasi data yang diberikan tanpa basa-basi atau kalimat bersayap.
        
        1. Lakukan klasifikasi NLP mendalam. Berikan persentase absolut untuk sentimen (Keluhan Kritis, Keluhan Ringan, Netral, Positif).
        2. Tulis evaluasi eksekutif sepanjang 3-4 paragraf yang sangat kritis. Jangan hanya merangkum; carilah kegagalan sistemik dari data tersebut.
        3. Jika ada sentimen positif, catat secara singkat. Jika ada sentimen negatif, bongkar habis-habisan.
        
        ATURAN KETAT:
        - Gunakan sub-bullet point bersarang (nested list) untuk merinci poin-poin yang padat agar mudah dibaca.
        - DILARANG KERAS menggunakan kata-kata yang memperhalus keadaan.
        - DILARANG KERAS menggunakan EMOJI (seperti grafik, roket, tanda seru, dll) atau simbol informal.
        - Jangan gunakan format heading markdown (###) di bagian ini.
        """,
        "include_trend_chart": True 
    },
    {
        "id": "pattern_analysis", 
        "title": "BAB II - ANALISIS POLA & INVESTIGASI AKAR MASALAH (ROOT CAUSE)",
        "instructions": """
        Bedah dataset ini untuk menemukan anomali dan pola kegagalan yang berulang. Anda harus menemukan minimal 3 pola dominan.
        
        ATURAN KETAT: 
        - DILARANG KERAS menggunakan EMOJI atau simbol informal.
        - Gunakan sub-bullet point untuk memecah informasi kompleks.
        
        Untuk SETIAP pola, gunakan framework investigasi mendalam dengan format markdown berikut persis:
        
        ### [Nama Kategori Kegagalan / Pola Kritis]
        * **Bukti Empiris:** * (Kutip sentimen atau keluhan spesifik dari data)
          * (Sebutkan frekuensi atau pola kemunculannya)
        * **Investigasi Akar Masalah:** * (Gunakan prinsip 'Five Whys' menggunakan sub-bullet)
          * (Cari tahu di mana SOP internal atau sistem gagal)
        * **Risiko Fatal (Dampak Bisnis):** * (Jelaskan konsekuensi terburuk jika dibiarkan)
        """
    },
    {
        "id": "hr_recommendations", 
        "title": "BAB III - INTERVENSI MANAJEMEN & REKOMENDASI TAKTIS",
        "instructions": """
        Berdasarkan investigasi di Bab II, susun rencana intervensi untuk HR dan Top Management. 
        DILARANG memberikan saran umum. Anda harus memberikan instruksi operasional yang spesifik, terukur, dan menuntut akuntabilitas.
        
        ATURAN KETAT: 
        - DILARANG KERAS menggunakan EMOJI.
        - Gunakan hierarki sub-bullet untuk merinci langkah-langkah spesifik.
        
        Gunakan format markdown berikut persis:
        
        ### Intervensi Krisis (Tindakan < 30 Hari)
        * **Instruksi Operasional:** * [Langkah spesifik 1]
          * [Langkah spesifik 2]
        * **Akuntabilitas & Tenggat Waktu:** [Siapa divisi yang harus disalahkan jika ini tidak selesai, dan apa matriks keberhasilannya?]
        
        ### Restrukturisasi & Perbaikan Sistemik (Tindakan 3-6 Bulan)
        * **Perubahan SOP / Kurikulum / Infrastruktur:** [Desain ulang sistem apa yang diperlukan?]
        * **KPI Pencegahan:** [Metrik kuantitatif spesifik apa yang harus dipantau?]
        """
    }
]
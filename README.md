# Sentiment and Feedback Analyzer (AI-powered)

Aplikasi ini adalah sistem *enterprise* berbasis web yang dirancang khusus untuk divisi Quality Assurance (QA) dan Manajemen Inixindo Jogja. Sistem ini memanfaatkan kecerdasan buatan (*Large Language Models* via **Ollama**) dan *Vector Database* (**ChromaDB**) untuk membaca, mengklasifikasikan, dan menganalisis ribuan masukan (keluhan dan apresiasi) dari klien maupun siswa secara otomatis.

Alih-alih membaca *feedback* secara manual, sistem ini akan mengekstrak *insight* menggunakan *Natural Language Processing* (NLP) dan merangkumnya menjadi dokumen laporan komprehensif berformat Microsoft Word, lengkap dengan visualisasi data dan rekomendasi mitigasi.

## Fitur Utama

* **Time-Bound RAG Analytics**: Mampu memfilter dan menganalisis sentimen berdasarkan rentang waktu spesifik (Mingguan, Bulanan, Semesteran, Tahunan).
* **Smart Root-Cause Analysis**: Mengelompokkan masukan tidak terstruktur menjadi *Pain Points* yang jelas (misal: isu fasilitas, kompetensi instruktur, atau jaringan infrastruktur).
* **Auto-Generated Mitigation Plan**: AI secara otomatis menyusun kerangka perbaikan (*Start, Stop, Continue*) dan memvisualisasikannya ke dalam *Flowchart* tindakan korektif.
* **Evidence-Based Reporting**: Setiap klaim analisis di dalam laporan akan divalidasi dengan kutipan langsung (*verbatim*) dari teks keluhan/pujian *stakeholder* untuk menjaga objektivitas.
* **Enterprise OSINT**: Menarik tren kepuasan pelanggan secara *real-time* dari internet untuk *benchmarking* standar layanan IT di Indonesia.

## Prasyarat Sistem

* **Python 3.9+** (Untuk pengembangan lokal tanpa Docker).
* **Ollama**: Menjalankan *local daemon* di port `11434` (atau *Ollama Cloud Endpoint*).
* **Serper API**: Membutuhkan `SERPER_API_KEY` untuk mengaktifkan modul OSINT.
* **Docker & Docker Compose** (Untuk *deployment* ke server *cloud* seperti AWS).

## Instalasi Lokal (Development)

### 1. Persiapan Lingkungan Virtual
Sangat disarankan menggunakan *virtual environment* agar dependensi aplikasi terisolasi.

```bash
# Buat virtual environment
python3 -m venv venv

# Aktifkan virtual environment (Mac/Linux)
source venv/bin/activate
# ATAU untuk Windows
# venv\Scripts\activate
```

### 2. Instalasi Dependensi
Instal seluruh *library* yang dibutuhkan dengan perintah berikut:

```bash
pip install flask flask-cors pandas chromadb ollama matplotlib python-docx markdown beautifulsoup4 requests Pillow sqlalchemy gunicorn
```

### 3. Konfigurasi Sistem
Atur *environment variable* sesuai mode yang ingin dijalankan:

* **Demo mode**: `APP_MODE=demo`
* **Hybrid mode**: `APP_MODE=hybrid`
* **Routing AI**: Pastikan `OLLAMA_HOST` mengarah ke endpoint Ollama yang aktif.
* **OSINT**: Isi `SERPER_API_KEY` jika ingin mengaktifkan benchmark eksternal.
* **Internal API**: Untuk `APP_MODE=hybrid`, isi `INTERNAL_API_BASE_URL`, dan jika perlu `INTERNAL_API_KEY` serta `INTERNAL_API_FEEDBACK_ENDPOINT`.

Contoh menjalankan aplikasi dari command line:

```bash
# Demo mode: data internal dari CSV lokal, OSINT tetap aktif jika SERPER_API_KEY diisi
APP_MODE=demo python app.py

# Hybrid mode: data internal dari API perusahaan, benchmark eksternal tetap lewat OSINT
APP_MODE=hybrid \
INTERNAL_API_BASE_URL=https://internal.example.com \
INTERNAL_API_KEY=your_api_key \
python app.py
```

### 4. Menyiapkan Model AI (Ollama)
**Langkah ini krusial.** RAG dan NLP Analyzer tidak akan berjalan tanpa model ini:

```bash
# Model embedding untuk mengubah teks feedback menjadi vektor (Wajib untuk ChromaDB)
ollama pull bge-m3:latest

# Model LLM utama untuk penalaran sentimen (Sesuaikan dengan config.py Anda)
ollama pull gpt-oss:120b-cloud
```

### 5. Sumber Data Internal
Pada `APP_MODE=demo`, letakkan file `db.csv` di folder `Sentiment analyzer/data/`. Sistem menggunakan *mapping* dinamis, namun untuk hasil terbaik, pastikan strukturnya seperti ini:
* `Tipe Stakeholder` (Contoh: Instansi Pemerintah, BUMN, Personal)
* `Layanan` (Contoh: Pelatihan IT Security, Konsultasi Masterplan)
* `Rentang Waktu` (Contoh: 1 Bulan Terakhir (Monthly))
* `Rating` (1-5)
* `Komentar` (Teks masukan bebas)

Pada `APP_MODE=hybrid`, aplikasi akan mengambil data internal dari API perusahaan. Endpoint tersebut minimal perlu mengembalikan data feedback dalam format JSON yang memuat padanan untuk kolom berikut:
* `Tipe Stakeholder`
* `Layanan`
* `Rentang Waktu` atau tanggal feedback
* `Rating`
* `Komentar`

Data internal yang berhasil diambil akan disalin ke `cx_feedback.db` (SQLite) sebagai cache lokal, lalu diproses ke ChromaDB.

### 6. Menjalankan Aplikasi
```bash
python app.py
```
Akses UI *Analyzer* melalui browser di `http://127.0.0.1:5000`.

---

## Deployment ke Production (AWS / Cloud)

Aplikasi ini sepenuhnya *Dockerized* dan siap untuk metode *Lift and Shift* ke server produksi (misal: AWS EC2).

1.  Siapkan VM/Instance di *cloud environment* Anda.
2.  Salin seluruh *source code* ke dalam server.
3.  Jalankan perintah berikut:

```bash
docker-compose up -d --build
```

Arsitektur Docker ini akan memutar *image* web menggunakan **Gunicorn** (4 *workers* untuk menangani banyak permintaan berbarengan), menghubungkannya ke jaringan Ollama internal, dan menjaga data *feedback* tetap aman menggunakan *persistent volumes*.

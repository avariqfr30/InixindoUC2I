# Feedback Intelligence Analyzer

Aplikasi ini adalah platform *feedback intelligence* berbasis web yang dirancang untuk membantu tata kelola feedback yang handal di Inixindo Jogja. Sistem ini mengolah feedback dari berbagai sumber internal, menambahkan benchmark OSINT untuk konteks eksternal, lalu menyusun laporan yang berfokus pada analisis descriptive, diagnostic, predictive, dan prescriptive.

Alih-alih membaca *feedback* secara manual, sistem ini menormalkan data, menghitung indikator risiko dan kekuatan layanan, merangkum bukti verbatim, dan menghasilkan dokumen Word yang siap dipakai untuk pengambilan keputusan.

## Fitur Utama

* **Multi-Source Feedback Governance**: Menormalkan feedback internal dari CSV demo atau API perusahaan ke dalam skema yang konsisten.
* **4-Layer Analytics**: Setiap laporan menyajikan descriptive, diagnostic, predictive, dan prescriptive analytics.
* **Evidence-Based Reporting**: Setiap bagian analisis tetap ditopang kutipan verbatim dan distribusi data nyata.
* **Enterprise OSINT**: Menarik tren pasar dan benchmark publik sebagai konteks eksternal.
* **Fast Report Pipeline**: Laporan disusun terutama dari analytics terstruktur agar lebih stabil untuk eksekusi paralel banyak pengguna.
* **Async Report Jobs for Pilot/VPS**: Jalur generate laporan dapat diproses sebagai background job dengan status polling, sehingga lebih aman untuk dipakai beberapa pengguna sekaligus pada satu server internal.

## Prasyarat Sistem

* **Python 3.9+** (Untuk pengembangan lokal tanpa Docker).
* **Ollama**: Opsional, dibutuhkan bila `ENABLE_VECTOR_INDEX=1` dan Anda ingin membangun embedding index lokal.
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
pip install flask flask-cors pandas chromadb ollama matplotlib python-docx markdown beautifulsoup4 requests Pillow sqlalchemy waitress
```

### 3. Konfigurasi Sistem
Atur *environment variable* sesuai mode yang ingin dijalankan:

* **Demo mode**: `APP_MODE=demo`
* **Hybrid mode**: `APP_MODE=hybrid`
* **Routing AI**: Pastikan `OLLAMA_HOST` mengarah ke endpoint Ollama yang aktif.
* **OSINT**: Isi `SERPER_API_KEY` jika ingin mengaktifkan benchmark eksternal.
* **Internal API**: Untuk `APP_MODE=hybrid`, isi `INTERNAL_API_BASE_URL`, dan jika perlu `INTERNAL_API_KEY` serta `INTERNAL_API_FEEDBACK_ENDPOINT`.
* **Vector index opsional**: Set `ENABLE_VECTOR_INDEX=1` bila ingin membangun embedding index. Secara default dimatikan agar startup dan eksekusi laporan lebih cepat dan lebih ringan.

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

Untuk shared pilot internal, jalankan aplikasi dengan Waitress:

```bash
cd "Sentiment analyzer"
APP_MODE=demo ./run_pilot.sh
```

Atau dengan pengaturan yang lebih eksplisit:

```bash
cd "Sentiment analyzer"
APP_MODE=hybrid \
HOST=0.0.0.0 \
PORT=8000 \
WAITRESS_THREADS=8 \
WAITRESS_CONNECTION_LIMIT=100 \
WAITRESS_CHANNEL_TIMEOUT=240 \
REPORT_JOB_WORKERS=3 \
REPORT_MAX_PENDING_JOBS=24 \
./run_pilot.sh
```

Setelah aktif, karyawan dapat membuka aplikasi dari browser mereka menggunakan:

```text
http://<IP-atau-host-internal>:8000
```

Endpoint health check tersedia di:

```text
http://<IP-atau-host-internal>:8000/health
```

Untuk readiness check yang juga memeriksa kesiapan data dan direktori artefak laporan:

```text
http://<IP-atau-host-internal>:8000/ready
```

### 4. Menyiapkan Model Embedding (Opsional)
Langkah ini hanya diperlukan bila Anda mengaktifkan `ENABLE_VECTOR_INDEX=1`:

```bash
# Model embedding untuk mengubah teks feedback menjadi vektor (Wajib untuk ChromaDB)
ollama pull bge-m3:latest

```

### 5. Sumber Data Internal
Pada `APP_MODE=demo`, letakkan file `db.csv` di folder `Sentiment analyzer/data/`. Sistem menggunakan *mapping* dinamis, namun untuk hasil terbaik, pastikan strukturnya seperti ini:
* `Record ID`
* `Sumber Feedback`
* `Kanal Feedback`
* `Tanggal Feedback`
* `Tipe Stakeholder` (Contoh: Instansi Pemerintah, BUMN, Personal)
* `Layanan` (Contoh: Pelatihan IT Security, Konsultasi Masterplan)
* `Lokasi`
* `Tipe Instruktur` (Contoh: `Internal`, `OL`)
* `Rentang Waktu` (Contoh: 1 Bulan Terakhir (Monthly))
* `Rating` (1-5)
* `Komentar` (Teks masukan bebas)
* `Customer Journey Hint` (Opsional, untuk mengunci area journey tertentu bila sudah tersedia)

Untuk kebutuhan simulasi internal yang lebih stabil, gunakan reseed utility berikut agar dataset demo selalu terisi dengan data sintetis yang lebih lengkap dan konsisten:

```bash
cd "Sentiment analyzer"
python3 seed_demo_data.py
```

Jika ingin otomatis melakukan reseed sebelum internal stress test dijalankan:

```bash
cd "Sentiment analyzer"
APP_MODE=demo RESEED_DEMO_DATA=1 ./run_pilot.sh
```

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
Akses UI *Analyzer* melalui browser di `http://127.0.0.1:8000`.

---

## Deployment ke Production (AWS / Cloud)

Aplikasi ini sepenuhnya *Dockerized* dan siap untuk metode *Lift and Shift* ke server produksi (misal: AWS EC2).

1.  Siapkan VM/Instance di *cloud environment* Anda.
2.  Salin seluruh *source code* ke dalam server.
3.  Jalankan perintah berikut:

```bash
docker-compose up -d --build
```

Arsitektur Docker dapat disesuaikan nanti untuk deployment produksi. Untuk pilot internal yang ringan dan cepat dibagikan, launcher saat ini langsung memakai **Waitress** agar jalur eksekusinya sederhana dan stabil untuk koneksi paralel dari browser karyawan.

## Catatan Operasional untuk VPS Simulation

Pada jalur UI terbaru, permintaan report tidak lagi harus menunggu file Word selesai dibangun di request yang sama. Browser akan:

1. membuat report job ke server,
2. memantau status job secara periodik,
3. mengunduh file Word ketika status sudah `completed`.

Endpoint yang relevan:

* `POST /generate-job` untuk membuat job report.
* `GET /jobs/<job_id>` untuk membaca status, durasi, dan metadata job.
* `GET /download/<job_id>` untuk mengambil dokumen yang sudah selesai.

Environment variable tambahan yang berguna untuk simulasi VPS:

* `REPORT_JOB_WORKERS`: jumlah worker background untuk generate report.
* `REPORT_MAX_PENDING_JOBS`: batas job aktif (`queued` + `running`) agar server tidak overload.
* `REPORT_JOB_RETENTION_SECONDS`: masa simpan metadata dan file report sebelum dibersihkan.
* `REPORT_ARTIFACT_DIR`: direktori penyimpanan file report hasil generate.
* `OSINT_CACHE_TTL_SECONDS`: masa hidup cache OSINT agar query benchmark tidak selalu memukul layanan eksternal.

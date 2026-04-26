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
* **Scoring Parameterized**: Bobot `Learning Score`, `Service Score`, `Facility Score`, dan formula `Experience Index` mengikuti parameter tabel resmi (berbasis `Feedback Score.xlsx`).

## Model Scoring

Model scoring sekarang mengikuti parameter yang Anda berikan:

* `Learning Score`: bobot indikator delivery, engagement, relevance, dan learning outcome.
* `Service Score`: bobot indikator attitude, responsiveness, competence, transport, dan souvenir.
* `Facility Score`: bobot indikator classroom comfort, equipment, supporting facilities, dan accessibility.
* `Experience Index`: dihitung dari gabungan komponen:
  * `Learning Score` = 50%
  * `Service Score` = 30%
  * `Facility Score` = 20%

Implementasi parameter ada di:

* `Sentiment analyzer/config.py` (`SCORE_ENGINE_PARAMETER_TABLES`, `SCORE_ENGINE_PROFILES`)
* `Sentiment analyzer/core.py` (`_score_engine_metrics_single`, `_score_engine_metrics_experience_index`)

## Prasyarat Sistem

* **Python 3.9+**
* **Ollama**: Opsional, hanya dibutuhkan bila `ENABLE_VECTOR_INDEX=1`.
* **Serper API**: Dibutuhkan bila ingin mengaktifkan OSINT eksternal.

## Jalur Operasional yang Disarankan

Untuk mempermudah perpindahan dari development ke production, gunakan dua profil operasional yang terpisah:

* **`demo`**: memakai dataset demo lokal untuk presentasi, UAT ringan, dan stress test.
* **`production`**: memakai connector JSON + endpoint API internal perusahaan.

Operator tidak perlu mengubah kode. Jalur yang disiapkan sekarang adalah:

1. siapkan file profil runtime,
2. siapkan connector runtime bila memakai profile `production`,
3. jalankan validasi satu perintah,
4. start aplikasi.

Semua itu dibungkus lewat script `appctl`.

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
Gunakan script kontrol berikut dari folder aplikasi:

```bash
cd "Sentiment analyzer"
```

Inisialisasi profil runtime:

```bash
./appctl init-profile demo
./appctl init-profile production
```

File yang akan dibuat:

* `profiles/demo.env`
* `profiles/production.env`
* `internal_connector.production.json` untuk profil production

Setelah itu:

* edit `profiles/demo.env` bila perlu mengganti port/thread demo,
* edit `profiles/production.env` untuk secret, auth, dan tuning runtime,
* edit `internal_connector.production.json` untuk endpoint API, request body, record path, dan field map.

Validasi profil sebelum menjalankan aplikasi:

```bash
./appctl validate demo
./appctl validate production
```

Jalankan aplikasi:

```bash
./appctl start demo
./appctl start production
```

Jika Anda masih ingin menjalankan langsung tanpa `appctl`, `run_pilot.sh` sekarang bisa membaca `PROFILE_ENV_FILE`:

```bash
PROFILE_ENV_FILE="$(pwd)/profiles/demo.env" ./run_pilot.sh
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

Pada profil production, aplikasi akan mengambil data internal dari API perusahaan. Endpoint tersebut minimal perlu mengembalikan data feedback dalam format JSON yang memuat padanan untuk kolom berikut:
* `Tipe Stakeholder`
* `Layanan`
* `Rentang Waktu` atau tanggal feedback
* `Rating`
* `Komentar`

Data internal yang berhasil diambil akan disalin ke `cx_feedback.db` (SQLite) sebagai cache lokal, lalu diproses ke ChromaDB.

Untuk operator production, jalur yang sekarang disarankan adalah **connector spec**, bukan menumpuk banyak environment variable. File `internal_connector.production.json` menjadi satu sumber kebenaran untuk:

* endpoint URL,
* method,
* request body,
* record path,
* field mapping,
* required fields.

Workflow paling mudah untuk handover APIDog:

```bash
cd "Sentiment analyzer"

# Live endpoint: fetch, auto-discover JSON records, write connector, then validate.
./appctl connect-api production https://api-company.example/feedback

# Saved APIDog response: inspect JSON shape first, write connector, then validate.
./appctl connect-api production \
  https://api-company.example/feedback \
  --file /tmp/apidog-feedback-response.json

# When using --file, run validation after the live endpoint and credentials are ready.
./appctl validate production
```

Untuk endpoint APIDog/live API yang memakai Bearer token, isi `profiles/production.env` seperti ini:

```env
INTERNAL_API_AUTH_MODE=api_key
INTERNAL_API_AUTH_HEADER=Authorization
INTERNAL_API_AUTH_PREFIX=Bearer
INTERNAL_API_KEY=your_token_here
```

Setelah connector tersimpan, profile `production` akan memakai endpoint tersebut sebagai internal knowledge base. Pada start atau `POST /refresh-knowledge`, aplikasi akan mengambil data API, menormalisasi field, menulis cache ke `cx_feedback.db`, lalu memperbarui index pengetahuan bila `ENABLE_VECTOR_INDEX=1`.

Layer internal API sekarang akan:
* menerima endpoint bernama atau URL penuh,
* mencoba menemukan list record yang paling relevan di dalam JSON secara otomatis,
* me-*flatten* object JSON bertingkat menjadi kolom yang bisa dibaca workflow,
* memetakan field hasil flatten ke kolom kanonik seperti `Rating`, `Komentar`, `Tanggal Feedback`, dan `Tipe Stakeholder`.
* memberi ringkasan field wajib agar operator tahu apakah response JSON sudah siap menjadi knowledge base.

Contoh file referensi tersedia di:

* `Sentiment analyzer/internal_api_endpoints.example.json`
* `Sentiment analyzer/internal_connector.production.example.json`

Untuk melihat endpoint yang sudah terdaftar, mencoba fetch API internal, atau mengecek response APIDog dari file:

```bash
cd "Sentiment analyzer"
python3 inspect_internal_api.py
python3 inspect_internal_api.py feedback
python3 inspect_internal_api.py feedback --fetch
python3 inspect_internal_api.py https://xxx.com/api/tag --fetch
python3 inspect_internal_api.py https://xxx.com/api/tag --method POST --body-mode json --params-json '{}'
python3 inspect_internal_api.py https://xxx.com/api/tag --file /tmp/apidog-feedback-response.json
./appctl inspect-api production https://xxx.com/api/tag --file /tmp/apidog-feedback-response.json --write-connector
./appctl connect-api production https://xxx.com/api/tag
```

### 6. Menjalankan Aplikasi
```bash
python app.py
```
Akses UI *Analyzer* melalui browser di `http://127.0.0.1:8000`.

---

## Deployment ke Production / VPS

Jalur deployment yang aktif saat ini adalah:

* **Waitress**
* **systemd**
* **nginx reverse proxy**

Ini lebih sesuai dengan kebutuhan pilot internal dan VPS tunggal dibanding memaksakan Docker padahal runtime saat ini memang dijalankan langsung dengan Python virtualenv.

Alur aman yang direkomendasikan:

1. simpan runtime config rahasia di file VPS terpisah atau `profiles/*.env` yang tidak ikut ter-*sync*,
2. simpan `internal_connector.production.json` versi runtime di VPS,
3. jalankan `./appctl validate production`,
4. start atau restart service,
5. cek `/health` dan `/ready`.

File runtime berikut sengaja diperlakukan sebagai konfigurasi lokal dan tidak ikut didorong ulang oleh deploy script:

* `Sentiment analyzer/profiles/*.env`
* `Sentiment analyzer/internal_connector.production.json`

## Catatan Operasional untuk VPS Simulation

Pada jalur UI terbaru, permintaan report tidak lagi harus menunggu file Word selesai dibangun di request yang sama. Browser akan:

1. membuat report job ke server,
2. memantau status job secara periodik,
3. mengunduh file Word ketika status sudah `completed`.

Endpoint yang relevan:

* `POST /generate-job` untuk membuat job report.
* `GET /jobs/<job_id>` untuk membaca status, durasi, dan metadata job.
* `GET /download/<job_id>` untuk mengambil dokumen yang sudah selesai.

Metadata job sekarang juga memisahkan:

* `queue_wait_seconds`
* `generation_seconds`
* `total_elapsed_seconds`
* `quality.completeness_score`
* `quality.verified_complete`

Dengan begitu, KPI waktu per tugas dan target dokumen `80% finished` bisa dibaca lebih jelas saat simulasi VPS berlangsung.

Environment variable tambahan yang berguna untuk simulasi VPS:

* `REPORT_JOB_WORKERS`: jumlah worker background untuk generate report.
* `REPORT_MAX_PENDING_JOBS`: batas job aktif (`queued` + `running`) agar server tidak overload.
* `REPORT_JOB_RETENTION_SECONDS`: masa simpan metadata dan file report sebelum dibersihkan.
* `REPORT_ARTIFACT_DIR`: direktori penyimpanan file report hasil generate.
* `OSINT_CACHE_TTL_SECONDS`: masa hidup cache OSINT agar query benchmark tidak selalu memukul layanan eksternal.
* `OSINT_QUERY_WORKERS`: jumlah query OSINT paralel untuk pencarian eksternal.
* `OSINT_DEEP_SCRAPE_MAX_CHARS`: batas teks sumber eksternal yang dibaca untuk insight mendalam.
* `OSINT_TRUSTED_DOMAINS`: daftar domain prioritas untuk menaikkan kualitas ranking sumber.
* `OSINT_BLOCKED_DOMAINS`: daftar domain yang tidak boleh masuk ke sinyal OSINT.
* `SESSION_IDLE_TIMEOUT_SECONDS`: batas idle session sebelum pengguna wajib login ulang.
* `SESSION_MAX_ACTIVE_PER_USER`: batas jumlah session aktif per akun untuk mengurangi penyalahgunaan akun bersama.
* `SESSION_MAX_ACTIVE_TOTAL`: batas total session aktif aplikasi agar beban login tetap terkendali.

File runtime seperti `data/*.db`, `data/report_jobs.json`, dan `data/.osint_cache/` dihasilkan oleh aplikasi dan tidak perlu dikomit. Dataset demo utama tetap berada di `data/db.csv`.

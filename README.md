# Seed Investor Pitch Deck Generator

Aplikasi Streamlit untuk membuat pitch deck investor tahap seed dalam format `.pptx` dan PDF latihan pitching dalam format `.pdf`.

Footer aplikasi, setiap slide PPTX, dan PDF guide:

**Developed by Galuh Adi Insani**

## Pembaruan Versi v8.1 - Perbaikan Metrik Model Bisnis

- Ditambahkan pilihan **durasi pitching**:
  - 3 menit - elevator pitch
  - 5 menit - demo day pitch
  - 8 menit - seed pitch ringkas
  - 10 menit - seed standard
  - 15 menit - detailed seed meeting
  - 20 menit - deep-dive seed meeting
- Skema PPT otomatis menyesuaikan durasi:
  - durasi pendek menggabungkan beberapa topik agar konteks tetap lengkap,
  - durasi menengah membuat deck seed ringkas,
  - durasi panjang memakai full deck dengan Competition dan Milestones yang bisa multi-slide.
- PDF **Pitch Scenario Guide** sekarang menyesuaikan:
  - jumlah slide,
  - timing tiap slide,
  - urutan skenario,
  - ritme pitching,
  - talk track sesuai durasi yang dipilih.
- Ditambahkan pilihan **model bisnis**:
  - SaaS / Subscription
  - Marketplace / Take Rate
  - E-commerce / D2C
  - Transaction Fee / Fintech
  - Usage-Based / API
  - Freemium
  - Enterprise / Licensing
  - Service-Enabled Software
  - Advertising / Media
  - Hybrid / Other
- Setiap model bisnis memiliki:
  - penjelasan sederhana,
  - model yang cocok,
  - cara menjelaskan saat pitching,
  - rumus praktis,
  - metrik utama yang bisa diedit.
- Slide **Business Model**, PDF skenario, dan analisa otomatis sekarang mengikuti model bisnis yang dipilih.

- Perbaikan bug `NameError: model_metric_labels is not defined` pada saat aplikasi dijalankan di Streamlit Cloud.
- Metrik utama sesuai model bisnis sekarang ditampilkan sebagai input yang bisa diedit.
- Nilai metrik tersebut otomatis masuk ke slide Business Model, PDF skenario, dan analisa.
- PDF guide memiliki appendix baru: **Model bisnis startup dan metrik utama**.

## Fitur Utama

- Struktur deck seed-stage startup yang story-led dan data-first.
- Durasi pitching dinamis: deck otomatis dipadatkan atau diperluas.
- Model bisnis dinamis dengan metrik yang dapat disesuaikan.
- Kompetitor dinamis hingga 10 alternatif.
- Milestone execution plan dengan periode, target, success metric, dan owner.
- Desain profesional: clean, minim teks, banyak whitespace, dan adaptif terhadap input panjang.
- Tampilan aplikasi theme-aware untuk light, dark, atau system theme.
- Elemen/emblem bawaan Streamlit disembunyikan.
- Kamus istilah startup dan investor dengan rumus/cara menghitung.
- Analisa otomatis:
  - Deck readiness score
  - Revenue growth Y1->Y2 dan Y2->Y3
  - Burn Year 1
  - Estimasi runway
  - Profit margin Year 3
  - Funding-to-Year-1-cost ratio
  - Strengths
  - Risk / investor questions
  - Rekomendasi pitching
- Output:
  - `.pptx` pitch deck investor
  - `.pdf` Pitch Scenario Guide untuk latihan

## Skema Slide Berdasarkan Durasi

### 3 menit
Deck sangat ringkas. Beberapa konteks digabung:
1. Cover
2. Problem + Solution + Product
3. Market + Business Model
4. Traction + GTM
5. Competition + Milestones
6. Team + Ask
7. Closing

### 5 menit
Deck demo day singkat:
1. Cover
2. Problem + Solution
3. Product
4. Market + Business Model
5. Traction + GTM
6. Competition + Milestones
7. Financials + Ask
8. Team + Closing

### 8-10 menit
Deck seed ringkas:
1. Cover
2. Problem
3. Solution
4. Product
5. Market
6. Business Model
7. Traction
8. GTM + Competition
9. Financials + Milestones
10. Team
11. Fundraising Ask + Closing

### 15-20 menit
Full seed investor deck:
1. Cover
2. Problem
3. Solution
4. Product
5. Market
6. Business Model
7. Traction
8. Go-To-Market
9. Competition, otomatis multi-slide jika data banyak
10. Financials
11. Milestones, otomatis multi-slide jika data banyak
12. Investor Readiness, opsional
13. Team
14. Fundraising Ask
15. Closing

## Cara Menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Struktur File

```text
pitch_deck_seed_generator/
├── .streamlit/
│   └── config.toml
├── docs/
│   └── pitch_scenario_guide_sample.pdf
├── app.py
├── requirements.txt
└── README.md
```

## Cara Menggunakan

1. Pilih warna aksen, mata uang, dan **durasi pitching** di sidebar.
2. Isi identitas startup dan round pendanaan.
3. Isi story: problem, evidence, solution, product flow, dan benefit.
4. Pilih **jenis model bisnis** pada tab Market, Model & Traction.
5. Sesuaikan metrik model bisnis yang muncul otomatis.
6. Isi traction, GTM, financial projection, use of funds, kompetitor, dan milestones.
7. Buka tab **Analisa** untuk membaca kesiapan pitching.
8. Buka tab **Istilah & Rumus** jika ingin mempelajari istilah dan cara menghitung metrik.
9. Klik **Generate Seed Investor Pitch Deck**.
10. Download PPTX dan PDF scenario guide.

## Catatan

CSS di dalam `app.py` sudah menyembunyikan elemen bawaan Streamlit. Namun, tampilan akhir tetap dapat sedikit berbeda tergantung versi Streamlit dan platform deployment.

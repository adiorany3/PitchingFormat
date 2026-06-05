# Seed Investor Pitch Deck Generator

Aplikasi Streamlit untuk membuat pitch deck investor tahap seed dalam format `.pptx` dan PDF latihan pitching dalam format `.pdf`.

Footer aplikasi, setiap slide PPTX, dan PDF guide:

**Developed by Galuh Adi Insani**

## Pembaruan Versi v7 - Istilah & Rumus Founder Baru

- Ditambahkan tab **Istilah & Rumus** di aplikasi.
- Setiap istilah startup/investor dilengkapi:
  - arti sederhana,
  - rumus atau cara menghitung,
  - contoh angka,
  - slide tempat istilah biasanya dipakai.
- Ditambahkan pencarian dan filter kategori istilah agar founder baru lebih mudah belajar.
- PDF **Pitch Scenario Guide** sekarang memiliki appendix **Istilah investor dan cara menghitungnya**.
- Istilah yang dicakup antara lain TAM, SAM, SOM, MRR, ARR, GMV, ARPU, CAC, CAC Payback, Gross Margin, Retention, Churn, Runway, Burn Rate, Valuation, Dilution, Milestone, Moat, GTM, dan lainnya.

## Pembaruan Versi v6 - PDF Pitch Scenario Guide

- Ditambahkan output **Pitch Scenario Guide PDF** setelah generate.
- PDF mengikuti urutan slide PPTX yang dibuat aplikasi.
- Jika slide **Competition** atau **Milestones** otomatis terpecah menjadi beberapa slide, PDF juga mengikuti urutan tersebut.
- PDF berisi panduan latihan pitching untuk setiap slide dan appendix istilah investor:
  - Timing yang disarankan
  - Tujuan slide
  - Data penting yang perlu disebutkan
  - Talk track / narasi bicara
  - Transisi ke slide berikutnya
  - Pertanyaan investor yang perlu disiapkan
- Ditambahkan final rehearsal checklist agar founder dapat mengecek kesiapan sebelum meeting investor.
- Ditambahkan dependency `reportlab` untuk pembuatan PDF.

## Pembaruan Versi v5 - Layout Rapi & Adaptif

- Slide Competition diubah menjadi comparison matrix agar tidak berantakan ketika kompetitor lebih dari satu.
- Kompetitor dibagi maksimal 5 baris per slide dengan tabel adaptif dan pemotongan teks panjang.
- Slide Milestones diubah menjadi timeline yang lebih lapang, maksimal 3 milestone per slide.
- Ukuran font pada headline, metric, card, bullet, tabel, dan takeaway dibuat adaptif terhadap panjang teks.
- Teks panjang otomatis diringkas dengan elipsis agar tidak menumpuk keluar dari card.
- Screenshot/mockup produk dibatasi tinggi dan lebarnya agar tidak keluar area slide.
- Area footer, investor takeaway, dan konten utama diberi jarak aman agar tidak saling bertabrakan.

## Pembaruan Versi v4

- Kompetitor dinamis: pengguna dapat menambah 1 sampai 10 kompetitor/alternatif.
- Kompetitor bisa dikategorikan sebagai direct competitor, indirect competitor, status quo, adjacent tool, atau alternative.
- Slide Competition otomatis memakai comparison matrix dan memecah data menjadi beberapa slide jika jumlah kompetitor banyak.
- Ditambahkan milestone detail: periode, target, success metric, dan owner.
- Slide Milestones otomatis dibuat sebagai timeline adaptif dan dipecah menjadi beberapa slide jika milestone banyak.
- Slide Financials dan Fundraising Ask menampilkan next milestone agar hubungan antara ask, runway, dan eksekusi lebih jelas.
- Analisa investor readiness ikut membaca jumlah kompetitor dan milestone sebagai sinyal kelengkapan deck seed.

## Pembaruan Theme-Aware

- Tampilan aplikasi mengikuti tema Streamlit yang sedang aktif: light, dark, atau system theme.
- Warna teks, background, kartu panduan, panel analisa, tab, input, tombol, upload file, metric, dan footer memakai CSS variable bawaan Streamlit sehingga tetap kontras saat tema berubah.
- Konfigurasi `.streamlit/config.toml` tidak mengunci aplikasi ke light theme.
- Elemen/emblem bawaan Streamlit tetap disembunyikan.

## Fitur

- Struktur deck seed-stage startup.
- Kompetitor dinamis hingga 10 alternatif.
- Milestone execution plan dengan periode, target, success metric, dan owner.
- Desain profesional: clean, data-first, minim teks, banyak whitespace.
- Panduan pengisian pada bagian:
  1. Identitas
  2. Story
  3. Market & Traction
  4. Financial & Funding
  5. Team & Competition
  6. Istilah & Rumus
  7. Analisa
- Kamus istilah startup dan investor dengan rumus/cara menghitung:
  - TAM, SAM, SOM
  - Revenue, MRR, ARR, GMV, Take Rate
  - ARPU, Gross Margin, COGS, CAC, CAC Payback, LTV, LTV/CAC
  - Retention, Churn, MoM Growth, Pipeline, LOI
  - Runway, Burn Rate, Operating Cost, EBITDA/Profit, Profit Margin
  - Valuation, Dilution, Milestone, Success Metric, Moat, Status Quo
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
- Slide otomatis:
  1. Cover
  2. Problem
  3. Solution
  4. Product
  5. Market
  6. Business Model
  7. Traction
  8. Go-To-Market
  9. Competition, comparison matrix otomatis multi-slide
  10. Financials
  11. Milestones, timeline otomatis multi-slide
  12. Investor Readiness, opsional
  13. Team
  14. Fundraising Ask
  15. Closing
- Upload screenshot atau mockup produk.
- Speaker notes otomatis di setiap slide.
- Download hasil dalam format `.pptx`.
- Download **Pitch Scenario Guide** dalam format `.pdf`.
- Contoh PDF scenario guide tersedia di folder `docs/`.

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

## Cara Menggunakan PDF Scenario Guide

1. Isi data pitch seperti biasa.
2. Buka tab **Istilah & Rumus** jika ada istilah yang belum dipahami atau angka yang belum tahu cara dihitung.
3. Klik **Generate Seed Investor Pitch Deck**.
4. Download `.pptx` untuk presentasi investor.
5. Download `.pdf` untuk latihan skenario pitching.
6. Latih pitch mengikuti urutan PDF: tujuan slide, talk track, transisi, pertanyaan investor, appendix istilah, dan checklist.

## Catatan

CSS di dalam `app.py` sudah menyembunyikan elemen bawaan Streamlit. Namun, tampilan akhir tetap dapat sedikit berbeda tergantung versi Streamlit dan platform deployment.

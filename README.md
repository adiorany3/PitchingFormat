# Seed Investor Pitch Deck Generator

Aplikasi Streamlit untuk membuat pitch deck investor tahap seed dalam format `.pptx`.

Footer aplikasi dan setiap slide:

**Developed by Galuh Adi Insani**

## Pembaruan Versi Ini

- Elemen bawaan Streamlit disembunyikan: menu, toolbar, header, footer, deploy button, dan badge/emblem yang umum muncul di UI.
- Ditambahkan `.streamlit/config.toml` untuk membuat tampilan lebih bersih tanpa mengunci theme ke light/dark tertentu.
- Setiap bagian input memiliki keterangan dan help text agar pengguna lebih mudah memahami cara mengisi data.
- Ditambahkan tab **Analisa** untuk membaca input dan memberi insight otomatis tentang kesiapan pitch.
- Ditambahkan opsi **slide analisa otomatis** di PPTX untuk merangkum kekuatan, risiko pertanyaan investor, dan rekomendasi pitching.
- Footer **Developed by Galuh Adi Insani** tetap muncul di UI dan semua slide PPTX.


## Pembaruan Versi Theme-Aware

- Tampilan aplikasi sekarang mengikuti tema Streamlit yang sedang aktif: light, dark, atau system theme.
- Warna teks, background, kartu panduan, panel analisa, tab, input, tombol, upload file, metric, dan footer memakai CSS variable bawaan Streamlit sehingga tetap kontras saat tema berubah.
- Konfigurasi `.streamlit/config.toml` tidak lagi mengunci aplikasi ke light theme.
- Panel insight diperbaiki agar semua isi berada di dalam kartu yang sama dan terbaca jelas pada dark theme.
- Elemen/emblem bawaan Streamlit tetap disembunyikan.


## Pembaruan Versi v4

- Kompetitor sekarang dinamis: pengguna dapat menambah 1 sampai 10 kompetitor/alternatif.
- Kompetitor bisa dikategorikan sebagai direct competitor, indirect competitor, status quo, adjacent tool, atau alternative.
- Slide Competition otomatis menyesuaikan layout dan memecah data menjadi beberapa slide jika jumlah kompetitor lebih dari 4.
- Ditambahkan milestone detail: periode, target, success metric, dan owner.
- Slide Milestones otomatis dibuat dan juga dipecah menjadi beberapa slide jika milestone lebih dari 4.
- Slide Financials dan Fundraising Ask sekarang menampilkan next milestone agar hubungan antara ask, runway, dan eksekusi lebih jelas.
- Analisa investor readiness sekarang ikut membaca jumlah kompetitor dan milestone sebagai sinyal kelengkapan deck seed.

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
  6. Analisa
- Analisa otomatis:
  - Deck readiness score
  - Revenue growth Y1→Y2 dan Y2→Y3
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
  9. Competition, otomatis multi-slide bila kompetitor lebih dari 4
  10. Financials
  11. Milestones, otomatis multi-slide bila milestone lebih dari 4
  12. Investor Readiness, opsional
  13. Team
  14. Fundraising Ask
  15. Closing
- Upload screenshot atau mockup produk.
- Speaker notes otomatis di setiap slide.
- Download hasil dalam format `.pptx`.

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
├── app.py
├── requirements.txt
└── README.md
```

## Catatan

CSS di dalam `app.py` sudah menyembunyikan elemen bawaan Streamlit. Namun, tampilan akhir tetap dapat sedikit berbeda tergantung versi Streamlit dan platform deployment.

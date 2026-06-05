# Seed Investor Pitch Deck Generator

Versi: v11.0 - Indonesia Market Segment Adaptive  
Footer: Developed by Galuh Adi Insani

Aplikasi Streamlit untuk membuat paket pitching startup sekali download. Output ZIP berisi:

1. PPTX investor pitch deck.
2. PDF Pitch Scenario Guide untuk latihan alur pitching.
3. HTML Pitch Prompter untuk simulasi otomatis/offline dengan teks teleprompter custom.
4. JSON project data untuk diedit ulang di kemudian hari.

## Fitur utama

- Durasi pitching adaptif: 3, 5, 8, 10, 15, dan 20 menit.
- Skema slide menyesuaikan durasi tanpa menghilangkan konteks utama.
- Jenis pitch: Investor Seed Round, Demo Day, Pitch Competition, Corporate Partnership, Grant/Hibah, Incubator/Accelerator, dan Internal Business Proposal.
- Model bisnis adaptif: SaaS, Marketplace, E-commerce/D2C, Fintech, Usage-Based/API, Freemium, Enterprise/Licensing, Service-Enabled Software, Advertising/Media, Hybrid/Other.
- Target pasar Indonesia adaptif: B2C, B2B, B2G, B2B2C, B2B2G, dan Hybrid.
- Field target pasar: segment, buyer/persona, decision maker/payer, sales/procurement cycle, channel, procurement notes, dan konteks lokal Indonesia.
- Kalkulator metrik: MRR, ARR, ARPU, CAC, Gross Margin, Burn Rate, Runway, GMV, Take Rate.
- Glossary istilah investor dan cara menghitungnya.
- Validasi input dan warning edukatif.
- Investor readiness score per kategori.
- Q&A investor di UI dan PDF.
- Wizard mode dan Tabs mode.
- Save/load project JSON.
- Brand kit: logo, warna aksen, warna cover, style deck, font style.
- Output bahasa: Bahasa Indonesia, English, atau Bilingual untuk label deck.
- Hasil generate hanya satu ZIP.
- Simulasi presentasi otomatis seperti teleprompter: start/pause/reset, prev/next, fullscreen, speed control, progress bar, auto-scroll scenario, timeline slide, dan teks teleprompter yang bisa diedit manual per slide.

## Cara menjalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Struktur folder

```text
pitch_deck_seed_generator/
├── app.py
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml
└── docs/
    ├── pitch_scenario_guide_sample.pdf
    └── seed_investor_pitch_deck_sample.pptx
```

## Catatan penggunaan

1. Isi bagian Identitas terlebih dahulu.
2. Pilih durasi pitching dan jenis pitch.
3. Pilih target pasar Indonesia dan model bisnis agar slide, metrik, insight, Q&A, dan PDF menyesuaikan konteks startup.
4. Gunakan Kalkulator untuk menghitung metrik dasar.
5. Cek tab Analisa, Simulasi, dan Preview sebelum generate.
6. Gunakan tab Simulasi untuk latihan timing dan edit teks teleprompter sesuai gaya bicara presenter.
7. Klik Generate Pitching Package untuk mengunduh ZIP.



## v9.1 - Theme Contrast Fix

Perbaikan utama:
- Aplikasi tidak lagi memaksa `base = "light"` di konfigurasi Streamlit.
- Teks, panel, card, metric, expander, select dropdown, uploader, tombol, dan footer memakai CSS variable tema Streamlit.
- Semua panel edukasi dan analisa dibuat high-contrast agar tetap terbaca di light theme, dark theme, maupun system theme.
- Warna primary button dibuat memakai pasangan `text-color` dan `background-color` agar tidak hilang ketika accent color terlalu terang atau terlalu gelap.


## v9.2 - Adaptive Font No Ellipsis

Perbaikan utama:
- Teks panjang di PPT tidak lagi dipotong menjadi `...` atau elipsis.
- Font pada text box, bullet, card, metric, takeaway, dan tabel otomatis mengecil sesuai panjang konten.
- Slide Competition dibatasi menjadi maksimal 3 alternatif per slide agar kelemahan dan keunggulan tetap terbaca.
- Slide Milestones dibatasi menjadi maksimal 2 milestone per slide agar target, success metric, dan owner tetap tampil penuh.
- Dataframe/panel UI diberi CSS tambahan agar teks panjang wrap, bukan dipotong.
- Slide Financials memakai card untuk Next Milestone sehingga teks panjang tidak dipotong.


## v10 - Auto Rehearsal Prompter

Fitur baru:
- Tab **Simulasi** untuk latihan pitching otomatis di dalam Streamlit.
- Presentation view menampilkan ringkasan slide yang sedang dipitching.
- Teleprompter menampilkan skenario bicara dan auto-scroll sesuai durasi slide.
- Timer, progress bar, start/pause/reset, prev/next, fullscreen, dan speed control.
- Timing otomatis mengikuti pilihan durasi pitch: 3, 5, 8, 10, 15, atau 20 menit.
- File ZIP output sekarang menyertakan HTML prompter mandiri: `*-pitch-prompter.html`, sehingga latihan bisa dilakukan offline di browser.


## v10.3
- Perbaikan kontras khusus dropdown/selectbox BaseWeb agar menu yang terbuka tetap terbaca di light/dark/system theme.

## v10.5
- Bagian **Urutan slide dan timing** sekarang dirender dalam iframe terisolasi, bukan markdown/table biasa.
- Warna header, baris, border, dan teks tabel memakai nilai fixed high-contrast sehingga tidak bisa tertimpa theme Streamlit/browser.
- Memperbaiki kasus tulisan tidak terlihat karena background dan warna huruf sama.

## v10.6 - Forced Readable Theme

Perbaikan utama:
- Aplikasi sekarang dipaksa memakai light theme melalui `.streamlit/config.toml`.
- Q&A investor dan Istilah/Cara menghitung tidak lagi memakai dataframe/expander bawaan untuk konten utama.
- Q&A investor dan glossary dirender dalam iframe high-contrast terisolasi, sehingga warna tema Streamlit/browser tidak bisa membuat tulisan hilang.
- Code/rummus, dropdown, input, card, tab, expander, metric, dan panel edukasi memakai background terang dan teks gelap yang kontras.


## v11.0 - Indonesia Market Segment Adaptive

Fitur baru:
- Menambahkan pilihan target pasar Indonesia: B2C, B2B, B2G, B2B2C, B2B2G, dan Hybrid.
- Setiap target pasar memiliki panduan siapa user, siapa pembayar, decision maker, sales/procurement cycle, channel masuk pasar, dan konteks lokal Indonesia.
- Slide deck otomatis menambahkan/menyesuaikan slide **Target Customer** untuk pitch 8 menit ke atas.
- PDF Scenario Guide menambahkan appendix **Panduan target pasar Indonesia**.
- Q&A investor menambahkan pertanyaan tentang target pasar dan alasan pemilihan segmen.
- Glossary menambahkan istilah B2C, B2B, B2G, B2B2C, Sales Cycle, dan Decision Maker beserta cara menghitung/memetakannya.
- Investor readiness score menambahkan kategori **Target segment clarity**.
- Memperbaiki potensi DuplicateWidgetID pada field Retention.

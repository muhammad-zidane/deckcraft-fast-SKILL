---
name: deckcraft-fast-skill
description: "High-performance presentation generation skill for DeckCraft. Eliminates sequential MCP tool roundtrips via 1-shot script generation (User Input -> Plan -> Python Script -> Instant Batch CLI/Live TCP -> Native DeckCraft GUI & PPTX)."
---

# DeckCraft Fast Presentation Automation Skill (v1.0)

Anda adalah **Senior Presentation Designer & Automation Engineer** yang merancang slide presentasi di DeckCraft secara instan, estetis, dan profesional tanpa jeda latensi.

Tujuan utama skill ini adalah **mencegah Agent AI terjebak dalam pemanggilan tool MCP berulang-ulang** (puluhan kali bolak-balik turn LLM). Sebagai gantinya, agent menyusun presentasi dalam 1 siklus terstruktur: **User Input → Plan → Skrip Python → Eksekusi 1-Shot → Hasil Muncul di GUI & File PPTX**.

---

## 1. Alur Kerja Wajib (The Fast Protocol)

Ketika pengguna meminta pembuatan atau perubahan slide presentasi, ikuti 5 tahap wajib berikut:

```
[1. User Input]
       │  (Topik, outline, PRD, atau dokumen referensi)
       ▼
[2. Agent Plan]
       │  (Rancang jumlah slide, tema warna, struktur visual, & layout per slide)
       ▼
[★ 3. User Review & Approval Gateway ★]  <── WAJIB TUNGGU PERSETUJUAN USER!
       │  (Agent tampilkan plan ke user. DILARANG lanjut sebelum user menyetujui)
       ▼
[4. Agent Bikin Skrip Python]
       │  (Tulis 1 file skrip utuh mengimpor `DeckCraftClient` & `DeckComponents`)
       ▼
[5. Eksekusi 1-Shot & Muncul di GUI]
       │  (Jalankan `python build_deck.py` -> Selesai dalam < 1 detik!)
       ▼
[Hasil: Jendela DeckCraft GUI terbuka + Tersimpan .pptx & .deckcraft]
```

> [!CAUTION]
> **ATURAN WAJIB GATEWAY PERSETUJUAN (USER APPROVAL GATEWAY)**:
> Setelah membuat **Plan** (Tahap 2), Agent **WAJIB** menyajikan rencana slide tersebut kepada pengguna dan **BERHENTI untuk menunggu persetujuan / konfirmasi pengguna**.
> **DILARANG KERAS** langsung menulis atau mengeksekusi skrip Python sebelum pengguna menyatakan setuju / memberikan revisi ("Lanjut", "Setuju", "Proceed", dsb.).

> [!IMPORTANT]
> **DILARANG KERAS** memanggil tool MCP DeckCraft secara individual satu per satu (`add_slide`, `add_shape`, `set_text`, dll.) untuk membuat presentasi. Selalu tulis 1 skrip Python utuh dan jalankan via terminal setelah plan disetujui.

### Format Standar Presentasi Plan ke Pengguna
Ketika menyajikan Plan kepada user untuk disetujui, gunakan format terstruktur berikut:

```markdown
### 📋 Rencana Presentasi DeckCraft
- **Tema Visual**: `corporate_navy` (atau pilihan tema lain yang sesuai konteks)
- **Total Slide**: N Slide
- **Target Output**: `.pptx` & `.deckcraft`

#### Breakdown Slide:
1. **Slide 1: Cover / Sampul**
   - Komponen: `comp.cover_slide()`
   - Judul: "..."
   - Subtitle: "..."
2. **Slide 2: [Topik Slide 2]**
   - Komponen: `comp.comparison_columns()` / `comp.feature_cards_row()`
   - Poin Utama: ...
3. **Slide 3: [Topik Slide 3 - Data/KPI]**
   - Komponen: `comp.stat_cards_row()` / `client.add_chart()`
   - Data Metrik: ...
...
[Pertanyaan Konfirmasi]: "Apakah susunan rencana slide ini sudah sesuai, atau ada bagian yang ingin disesuaikan sebelum saya membuat dan menjalankan skripnya?"
```

---

## 2. Struktur Modul & Lokasi Impor Resmi

Gunakan modul resmi skill yang sudah teruji. Di setiap skrip builder, impor modul dari:

```python
import sys
from pathlib import Path

# Lokasi resmi modul skrip skill:
SKILL_SCRIPTS = Path.home() / ".gemini" / "config" / "skills" / "deckcraft-fast-skill" / "scripts"
if not SKILL_SCRIPTS.exists():
    # Fallback ke workspace lokal jika sedang pengembangan:
    SKILL_SCRIPTS = Path(r"C:\dev\deckcraft-fast-SKILL - Copy\scripts")
sys.path.insert(0, str(SKILL_SCRIPTS))

from deckcraft_ui import DeckCraftClient
from design_tokens import THEMES, calculate_columns
from deckcraft_components import DeckComponents
```

---

## 3. Template Dasar Skrip Python

Gunakan pola skrip berikut untuk membangun slide secara cepat dan andal:

```python
from pathlib import Path
from deckcraft_ui import DeckCraftClient
from design_tokens import THEMES
from deckcraft_components import DeckComponents

def create_presentation():
    output_dir = Path("./output")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    deck_file = output_dir / "my_presentation.deckcraft"
    pptx_file = output_dir / "my_presentation.pptx"

    # 1. Pilih tema visual (corporate_navy | modern_dark | clean_slate | vibrant_purple | emerald_growth | warm_editorial)
    theme = THEMES["corporate_navy"]

    # 2. Inisialisasi client (auto_gui=True akan otomatis membuka DeckCraft GUI)
    with DeckCraftClient(auto_gui=True) as client:
        client.set_outputs(
            save_deckcraft=str(deck_file),
            export_pptx=str(pptx_file)
        )
        comp = DeckComponents(client, theme=theme)

        # SLIDE 1: Cover Hero
        comp.cover_slide(
            title="Judul Presentasi Utama",
            subtitle="Penjelasan ringkas maksud dan tujuan presentasi",
            author="Tim Antigravity",
            tag="EXECUTIVE SUMMARY"
        )

        # SLIDE 2: Masalah vs Solusi (2 Kolom)
        client.new_slide(layout="blank")
        client.set_background(theme.bg_canvas)
        comp.header(title="Tantangan & Peluang", category="ANALISIS")
        comp.comparison_columns(
            left_col={"tag": "MASALAH", "title": "Kondisi Lama", "points": ["Poin 1", "Poin 2"]},
            right_col={"tag": "SOLUSI", "title": "Solusi Baru", "points": ["Solusi 1", "Solusi 2"], "highlight": True}
        )

        # SLIDE 3: KPI Metrics
        client.new_slide(layout="blank")
        client.set_background(theme.bg_canvas)
        comp.header(title="Kinerja Kunci", subtitle="Indikator utama keberhasilan proyek", category="METRICS")
        comp.stat_cards_row([
            {"value": "98%", "label": "Efisiensi Waktu", "subtext": "Penghematan signifikan", "icon": "zap"},
            {"value": "10x", "label": "Kecepatan Render", "subtext": "Sub-detik", "icon": "trending-up"}
        ])

        # SLIDE 4: Call-to-Action / Penutup
        comp.closing_slide(
            title="Terima Kasih",
            subtitle="Mari mulai kolaborasi bersama kami.",
            cta_text="Hubungi Kami"
        )

    print(f"Selesai! File tersimpan: {pptx_file}")

if __name__ == "__main__":
    create_presentation()
```

---

## 4. Standar Kanvas & Layout (16:9 Standard)

DeckCraft menggunakan koordinat slide berbasis poin dengan orientasi 16:9:
* **Dimensi Kanvas**: Lebar `960 pt`, Tinggi `540 pt`.
* **Safe Margin**: Kiri `60 pt`, Kanan `60 pt`, Atas `48 pt`, Bawah `48 pt`.
* **Area Konten Efektif**: Lebar `840 pt`, Tinggi `444 pt`.
* **Y Positions Standar**:
  - Header: `y = 48`
  - Subtitle: `y = 86 - 96`
  - Konten Utama / Kartu: `y = 150`
  - Tinggi Kartu Standar: `height = 320 - 330`

### Tipografi Skala:
* Cover Display Title: `42 pt` (Bold)
* Slide Header: `28 pt` (Bold)
* Subtitle: `16 pt` (Regular)
* Card Title: `18 pt` (Bold)
* Body / Poin: `14 pt` (Regular)
* Caption / Keterangan: `10 - 11 pt`
* KPI Big Number: `36 - 44 pt` (Bold)

---

## 5. Pilihan Tema Warna Terpasang (`THEMES`)

Gunakan tema yang selaras dengan pesan presentasi:
1. `"corporate_navy"`: Biru dongker, slate, latar putih/abu bersih. Cocok untuk presentasi bisnis, investor, dan laporan resmi.
2. `"modern_dark"`: Latar hitam elegan (`#0B0F19`), kartu gelap, aksen cyan/biru neon. Cocok untuk produk teknologi, AI, dan developer.
3. `"clean_slate"`: Latar putih bersih, aksen indigo (`#4F46E5`). Cocok untuk pitch deck startup modern.
4. `"emerald_growth"`: Nuansa hijau segar, ideal untuk presentasi finansial, sustainability, atau penjualan.
5. `"vibrant_purple"`: Latar ungu gelap dengan aksen violet dan coral, cocok untuk tema kreatif dan inovasi.
6. `"warm_editorial"`: Tipografi Playfair Display dengan palet warna hangat, cocok untuk studi kasus, brand fashion, atau edukasi.

---

## 6. Katalog Komponen Siap Pakai (`DeckComponents`)

* `comp.header(title, subtitle=None, category=None)`: Header slide rapi dengan badge kategori opsional.
* `comp.cover_slide(title, subtitle, author=None, tag="PRESENTATION", date_str=None)`: Slide sampul visual kuat.
* `comp.stat_cards_row(stats=[{"value": "...", "label": "...", "subtext": "...", "icon": "..."}], y=160, height=170)`: Deretan kartu metrik angka besar.
* `comp.feature_cards_row(cards=[{"title": "...", "body": "...", "badge": "...", "icon": "..."}], y=150, height=320)`: 2, 3, atau 4 kartu kolom otomatis terhitung jaraknya.
* `comp.comparison_columns(left_col, right_col, y=150, height=320)`: Dua kolom komparasi (Before vs After, Old vs New).
* `comp.timeline_steps(steps=[{"phase": "Q1", "title": "...", "desc": "..."}, ...], y=200)`: Roadmap tonggak sejarah / rencana masa depan.
* `comp.closing_slide(title, subtitle, cta_text, contact_info)`: Slide penutup dengan tombol aksi utama.

---

## 7. Komponen Lanjutan (`client`)

* **Chart**:
  ```python
  client.add_chart(
      chart_type="column",  # column, bar, line, pie, doughnut, area
      rect=[60, 160, 450, 280],
      categories=["Q1", "Q2", "Q3", "Q4"],
      series=[{"name": "Penjualan", "values": [120, 190, 240, 310]}]
  )
  ```
* **Table**:
  ```python
  client.add_table(
      rect=[60, 160, 840, 260],
      rows=4,
      cols=3,
      headers=["Fitur", "Paket Dasar", "Paket Enterprise"],
      data=[
          ["Kecepatan", "Standar", "Instan 1-Shot"],
          ["Kapasitas", "10 Slide", "Unlimited"],
          ["Format", "PPTX", "PPTX + Native"]
      ],
      style="medium2-accent1"
  )
  ```
* **Icon Bawaan**:
  Ikon vektor bawaan: `"zap"`, `"star"`, `"check"`, `"trending-up"`, `"target"`, `"lightbulb"`, `"chart-bar"`, `"users"`, `"file-text"`, `"globe"`, `"shield"`, dll.
  ```python
  client.add_icon(name="trending-up", rect=[cx, cy, 28, 28], fill="#2563EB")
  ```

---

## 8. Aturan Anti-Glitch & Kualitas Visual

1. **Jaga Batas Teks dalam Kartu**: Jangan biarkan teks lebih panjang dari tinggi kartu. Sesuaikan ukuran font atau ringkas kalimat jika teks meluap.
2. **Patuhi Safe Margin**: Hindari menempatkan elemen penting di luar koordinat `x < 60` atau `x > 900`, `y < 48` atau `y > 492`.
3. **Kombinasi Warna Kontras Tinggi**: Jangan gunakan teks abu-abu gelap di atas background gelap, atau teks putih di atas latar terang. Selalu gunakan `theme.text_primary` dan `theme.text_secondary`.
4. **Wajib Format PPTX**: Sesuai standar industri presentasi, setiap eksekusi wajib mengikutsertakan parameter `export_pptx` agar pengguna dapat langsung membuka dan membagikan file ke rekan kerja.

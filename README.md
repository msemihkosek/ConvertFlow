<div align="center">
  <img src="assets/logo.svg" alt="ConvertFlow Logo" width="120" height="120" />
  <h1>ConvertFlow</h1>
  <p>
    <strong>Windows, macOS ve Linux Sağ Tık Menüsüne Entegre Evrensel Dosya Dönüştürücü</strong>
  </p>
  <p>
    Tarayıcı veya arayüz açmadan, doğrudan dosya gezgininden tek tıkla <strong>47 dosya uzantısı</strong> arasında <strong>369+ dönüşüm</strong>.
  </p>

  <p>
    <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-0078D6?style=flat-square" alt="Cross Platform" />
    <img src="https://img.shields.io/badge/Node.js-18%2B-339933?style=flat-square&logo=node.js&logoColor=white" alt="Node.js" />
    <img src="https://img.shields.io/badge/Python-PyMuPDF%20%2B%20pdf2docx-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python" />
    <img src="https://img.shields.io/badge/Media-FFmpeg-007808?style=flat-square&logo=ffmpeg&logoColor=white" alt="FFmpeg" />
    <img src="https://img.shields.io/badge/Image-Sharp%20(libvips)-99CC00?style=flat-square" alt="Sharp" />
  </p>
</div>

---

## ✨ Özellikler

- **🖱️ Çoklu Platform Sağ Tık Entegrasyonu:**
  - **Windows 10/11:** Her dosya uzantısına özel kademeli alt menü (`ConvertFlow ile Dönüştür ▸`), sıfır yönetici izni.
  - **macOS (Finder):** Sağ Tık → **Hızlı İşlemler (Quick Actions)** → `ConvertFlow ile Dönüştür` (dosyaya özel yerel format seçici).
  - **Linux (GNOME Files / Nemo / Caja / KDE Dolphin):** Dosya yöneticisi sağ tık menüsü & `.desktop` servis menüsü.
- **🤫 Sessiz Arka Plan Çalışması & Bildirim:** Pencere açmadan arka planda çalışır; işlem başlangıç ve bitişinde yerel işletim sistemi bildirimi (`Windows Toast`, `macOS Notification`, `Linux notify-send`) gösterir.
- **🧠 Hibrit PDF Motoru (`PyMuPDF` + `pdf2docx`):**
  - **Metin/Tablo/Vektör PDF'ler:** Tablo, resim, sütun ve yazı düzenini koruyarak gerçek düzenlenebilir `.docx` üretir.
  - **Taranmış (Scan) PDF'ler:** Metin katmanı olmayan kitap/belgeleri otomatik algılar ve tüm sayfaları uyarlanabilir DPI kalitesiyle eksiksiz Word belgesine aktarır.
- **🏢 Yerel Microsoft Office COM (Windows):** Office kuruluysa Word, Excel ve PowerPoint belgelerini piksel kaybı olmadan doğrudan PDF veya görsel slayt olarak dışa aktarır.
- **🛡️ Çakışmasız Çıktı:** Dönüştürülen dosya aynı klasöre kaydedilir; aynı isimde dosya varsa üzerine yazmaz (`dosya (1).docx` şeklinde numaralandırır).

---

## 📂 Desteklenen Formatlar

| Kategori | Kaynak Formatlar | Hedef Formatlar |
| :--- | :--- | :--- |
| **Belgeler** | `pdf`, `docx`, `doc`, `rtf`, `pptx`, `ppt`, `txt`, `md`, `html` | `pdf`, `docx`, `pptx`, `html`, `txt`, `md`, `png`, `jpg`, `webp`, `zip` |
| **Görseller** | `png`, `jpg`, `webp`, `avif`, `heic`, `tiff`, `bmp`, `ico`, `svg`, `gif` | `png`, `jpg`, `webp`, `avif`, `ico` (çok katmanlı), `tiff`, `gif`, `pdf`, `mp4` (gif için), `zip` |
| **Video** | `mp4`, `mkv`, `avi`, `mov`, `webm`, `flv`, `wmv` | `mp4`, `mkv`, `webm`, `avi`, `mov`, `gif`, `mp3`, `wav`, `aac`, `flac`, `m4a`, `zip` |
| **Ses** | `mp3`, `wav`, `ogg`, `aac`, `flac`, `m4a`, `wma`, `opus` | `mp3`, `wav`, `aac`, `flac`, `ogg`, `m4a`, `opus`, `zip` |
| **Veri & Tablo** | `xlsx`, `xls`, `csv`, `tsv`, `json`, `xml`, `yaml`, `yml` | `xlsx`, `csv`, `json`, `xml`, `yaml`, `sql`, `html`, `md`, `pdf` (Excel için), `zip` |
| **Arşiv** | `zip` (veya tüm dosyalar) | `Klasöre Çıkart` / `ZIP (Arşivle)` |

---

## 🚀 Tek Tıkla / Tek Komutla Kurulum

Projeyi indirdikten sonra işletim sisteminize göre tek dosyayı çalıştırmanız yeterlidir — eksik bağımlılıkları, izole Python `.venv` ortamını ve sağ tık menüsünü otomatik kurar:

```bash
git clone https://github.com/msemihkosek/ConvertFlow.git
cd ConvertFlow
```

### 🪟 Windows
1. **`Windows_Kurulum.bat`** dosyasına çift tıklayın *(Node.js veya Python yoksa `winget` ile otomatik kurar, ardından sağ tık menüsünü ekler)*.
2. Kaldırmak için **`Windows_Kaldir.bat`** dosyasına çift tıklayın.

### 🍎 macOS & 🐧 Linux
```bash
chmod +x macOS_Linux_Kurulum.sh macOS_Linux_Kaldir.sh
./macOS_Linux_Kurulum.sh
```
*(Kaldırmak için `./macOS_Linux_Kaldir.sh`)*

> **Not:** Linux GNOME/Nemo kullanıyorsanız görsel format seçici için `zenity` ve bildirimler için `libnotify-bin` önerilir (`sudo apt install -y zenity libnotify-bin`).

---

## 💡 Kullanım

- **Windows 11:** Dosyaya **`Shift + Sağ Tık`** yapın *(veya Sağ Tık → **Diğer seçenekleri göster**)* → **`ConvertFlow ile Dönüştür`**.
  - *İsteğe bağlı:* Shift'e basmadan ilk sağ tıkta doğrudan görmek için `npm run menu:win11-classic` (`npm run menu:win11-modern` ile geri dönebilirsiniz).
- **Windows 10:** Normal sağ tık menüsünde doğrudan görünür.
- **macOS (Finder):** Sağ Tık → **Hızlı İşlemler (Quick Actions)** → **`ConvertFlow ile Dönüştür`**.
- **Linux:** GNOME Files / Nemo'da **Sağ Tık → Scripts → `ConvertFlow ile Dönüştür`**; KDE Dolphin'de **Sağ Tık → `ConvertFlow ile Dönüştür`**.

### Komut Satırından (CLI)
```bash
npm run convert -- "/yol/belge.pdf" docx
npm test   # 31 uçtan uca dönüşüm testini çalıştırır
```

---

## 🏗️ Proje Yapısı

```text
ConvertFlow/
├── Windows_Kurulum.bat          # Windows tek tıkla kurulum
├── Windows_Kaldir.bat           # Windows kaldırma
├── macOS_Linux_Kurulum.sh       # macOS & Linux tek komutla kurulum
├── macOS_Linux_Kaldir.sh        # macOS & Linux kaldırma
├── assets/
│   └── logo.svg                 # Proje logosu
├── context-menu/
│   ├── setup.js                 # Evrensel .venv + sağ tık kurulum sihirbazı
│   ├── formats.js               # Uzantı bazlı hedef format haritası
│   ├── engine.js                # Çapraz platform dönüştürme orkestratörü
│   ├── convert-cli.js           # Sağ tık & CLI giriş noktası + bildirimler
│   ├── install-menu.js          # Windows / macOS / Linux menü kurucusu
│   ├── uninstall-menu.js        # Windows / macOS / Linux menü kaldırıcı
│   ├── run-hidden.vbs           # Windows penceresiz sessiz çalıştırıcı
│   ├── notify.ps1               # Windows Toast bildirim betiği
│   └── win11-classic-menu.js    # Windows 11 menü görünüm değiştirici
├── server/
│   ├── converters/              # FFmpeg, Sharp, ExcelJS, Mammoth, Office COM modülleri
│   └── scripts/
│       ├── pdfEngine.py         # PyMuPDF + pdf2docx hibrit PDF motoru
│       ├── officeConvert.ps1    # Office COM PowerShell köprüsü (Windows)
│       └── pdfRender.ps1        # Windows WinRT PDF yedek işleyici
└── scratch/
    └── test_context_menu_engine.js
```

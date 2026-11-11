<div align="center">
  <img src="assets/logo.svg" alt="ConvertFlow Logo" width="120" height="120" />
  <h1>ConvertFlow</h1>
  <p>
    <strong>Windows Sağ Tık Menüsüne Entegre Evrensel Dosya Dönüştürücü</strong>
  </p>
  <p>
    Tarayıcı veya arayüz açmadan, doğrudan dosya gezgininden tek tıkla <strong>47 dosya uzantısı</strong> arasında <strong>369+ dönüşüm</strong>.
  </p>

  <p>
    <img src="https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011-0078D6?style=flat-square&logo=windows" alt="Windows" />
    <img src="https://img.shields.io/badge/Node.js-20%2B-339933?style=flat-square&logo=node.js&logoColor=white" alt="Node.js" />
    <img src="https://img.shields.io/badge/Python-PyMuPDF%20%2B%20pdf2docx-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python" />
    <img src="https://img.shields.io/badge/Media-FFmpeg-007808?style=flat-square&logo=ffmpeg&logoColor=white" alt="FFmpeg" />
    <img src="https://img.shields.io/badge/Image-Sharp%20(libvips)-99CC00?style=flat-square" alt="Sharp" />
  </p>
</div>

---

## ✨ Özellikler

- **🖱️ Doğrudan Sağ Tık Entegrasyonu:** Her dosya türüne özel kademeli alt menü (`ConvertFlow ile Dönüştür ▸`) — PDF'e sağ tıkladığınızda sadece PDF hedefleri, videoya tıkladığınızda sadece medya hedefleri görünür.
- **🤫 Sıfır Pencere / Sessiz Çalışma:** `wscript.exe` tabanlı görünmez başlatıcı sayesinde arkada siyah komut satırı penceresi açılmaz; büyük dosyalarda başlangıç ve bitişte yerel **Windows Toast Bildirimi** gösterir.
- **🧠 Hibrit PDF Motoru (`PyMuPDF` + `pdf2docx`):**
  - **Metin/Tablo/Vektör PDF'ler:** Tablo, resim, sütun ve yazı düzenini koruyarak gerçek düzenlenebilir `.docx` üretir.
  - **Taranmış (Scan) PDF'ler:** Metin katmanı olmayan kitap/belgeleri otomatik algılar ve tüm sayfaları uyarlanabilir DPI kalitesiyle eksiksiz Word belgesine aktarır.
- **🏢 Yerel Microsoft Office COM Desteği:** Bilgisayarda Office kuruluysa Word, Excel ve PowerPoint belgelerini piksel kaybı olmadan doğrudan PDF veya görsel slayt olarak dışa aktarır.
- **🛡️ Çakışmasız Çıktı:** Dönüştürülen dosya aynı klasöre kaydedilir; aynı isimde dosya varsa üzerine yazmaz (`dosya (1).docx` şeklinde numaralandırır).
- **🔓 Yönetici İzni Gerektirmez:** Tüm kayıt defteri işlemleri `HKCU\Software\Classes` altında kullanıcı bazlı yapılır.

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

## 🚀 Kurulum

### Gereksinimler
- **Windows 10 / 11**
- **Node.js** (v18 veya üzeri)
- **Python 3.10+** (Hibrit PDF motoru için)
- *(İsteğe bağlı)* **Microsoft Office** (Word / Excel / PowerPoint yerel COM dönüşümleri için)

### Adımlar

1. **Projeyi klonlayın ve klasöre girin:**
   ```bash
   git clone https://github.com/<kullanici-adin>/convertto.git
   cd convertto
   ```

2. **Node.js bağımlılıklarını yükleyin:**
   ```bash
   npm install
   ```

3. **Python PDF motoru bağımlılıklarını yükleyin:**
   ```bash
   pip install pymupdf pdf2docx python-docx
   ```

4. **Windows sağ tık menüsünü kurun:**
   ```bash
   npm run menu:install
   ```

---

## 💡 Kullanım

### Windows Sağ Tık ile
- **Windows 11:** Herhangi bir dosyanın üzerine **`Shift + Sağ Tık`** yapın *(veya Sağ Tık → **Diğer seçenekleri göster**)* → **`ConvertFlow ile Dönüştür`** üzerinden hedef formatı seçin.
- **Windows 10:** Normal sağ tık menüsünde doğrudan görünür.

> **İpucu (Windows 11):** Shift tuşuna basmadan ilk sağ tıkta doğrudan görmek isterseniz Windows 11 tam sağ tık menüsünü açabilirsiniz:
> ```bash
> npm run menu:win11-classic   # İlk sağ tıkta tam menü
> npm run menu:win11-modern    # Varsayılan Windows 11 menüsüne dönüş
> ```

### Komut Satırından (CLI)
```bash
npm run convert -- "C:\yol\belge.pdf" docx
```

### Diğer Komutlar
```bash
npm test                 # Uçtan uca dönüşüm testlerini çalıştır (31 test)
npm run menu:uninstall   # Sağ tık menüsünü kayıt defterinden tamamen kaldır
```

---

## 🏗️ Proje Yapısı

```text
convertto/
├── assets/
│   └── logo.svg                 # Proje logosu
├── context-menu/
│   ├── formats.js               # Uzantı bazlı hedef format haritası
│   ├── engine.js                # Merkezi dönüştürme orkestratörü
│   ├── convert-cli.js           # Sağ tık & CLI giriş noktası + log/bildirim
│   ├── run-hidden.vbs           # Penceresiz sessiz çalıştırıcı
│   ├── notify.ps1               # Windows Toast bildirim betiği
│   ├── install-menu.js          # HKCU kayıt defteri menü kurucusu
│   ├── uninstall-menu.js        # Menü kaldırıcı
│   └── win11-classic-menu.js    # Windows 11 menü görünüm değiştirici
├── server/
│   ├── converters/              # FFmpeg, Sharp, ExcelJS, Mammoth, Office COM modülleri
│   └── scripts/
│       ├── pdfEngine.py         # PyMuPDF + pdf2docx hibrit PDF motoru
│       ├── officeConvert.ps1    # Office COM PowerShell köprüsü
│       └── pdfRender.ps1        # Windows WinRT PDF yedek işleyici
└── scratch/
    └── test_context_menu_engine.js
```


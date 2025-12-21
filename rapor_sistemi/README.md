# Elektrik Tesisatı Rapor Sistemi v2.1

## 📋 Genel Bakış

Bu uygulama, elektrik tesisatı periyodik kontrol raporları oluşturmak için geliştirilmiş çoklu pano destekli bir GUI uygulamasıdır.

**ÖNEMLİ:** DOCX şablon dosyası (`Elektrik Tesisatı Gözle Kontrol ve Fonksyion Testleri Periyodik Kontrol Raporu.docx`) EXE içine gömülüdür. EXE dosyasını herhangi bir dizine kopyalayıp çalıştırabilirsiniz - ek dosya gerekmez!

## 🚀 Özellikler

### Ana Özellikler
- **Çoklu Pano Desteği**: Aynı firma için birden fazla pano raporu oluşturabilirsiniz
- **Ortak Bilgi Girişi**: Firma bilgileri bir kez girilir, tüm panolara uygulanır
- **Otomatik Rapor Numaralama**: TPK2025-XXXX-Y formatında otomatik numaralama
- **PDF Sözleşme Okuyucu**: İSG-KATİP hizmet sözleşmesi PDF'inden otomatik veri çekme
- **Termal Görüntü Entegrasyonu**: Fluke DOCX dosyalarından termal görsel çıkarma

### Fonksiyon Testleri
- Linye bazlı veri girişi
- Otomatik Ib hesaplaması (In × 0.7)
- Otomatik Iz hesaplaması (kablo kesitine göre)
- KAKR (Kaçak Akım Koruma Rölesi) desteği
- Toplu linye ekleme (Linye Grubu özelliği)
- Standart eğri tipleri (B, C, D, K, Z, AAA)

### Gözle Kontrol
- 29 adet standart kontrol maddesi
- Uygun / Uygun Değil / Uygulanamaz seçenekleri
- Otomatik kusur raporu oluşturma

### Desteklenen Formatlar
- **Giriş**: Excel (.xlsx), Fluke DOCX, İSG-KATİP PDF
- **Çıkış**: Word DOCX raporu

## 📁 Dosya Yapısı

```
rapor_sistemi/
├── multi_pano_gui.py       # Ana GUI uygulaması
├── report_generator.py     # Rapor üretici modülü
├── excel_reader.py         # Excel okuyucu
├── docx_writer.py          # DOCX yazıcı
├── fluke_extractor.py      # Fluke görsel çıkarıcı
├── sozlesme_parser.py      # PDF sözleşme parser
├── RaporSistemi.spec       # PyInstaller spec dosyası
├── DERLE.bat               # Derleme scripti
└── dist/
    └── ElektrikRaporSistemi.exe  # Derlenmiş uygulama
```

## 💻 Kullanım

### Derlenmiş Sürüm (EXE)
1. `dist/ElektrikRaporSistemi.exe` dosyasını çalıştırın
2. DOCX şablon dosyasının `config/system_config.json` içinde tanımlı olduğundan emin olun

### Python ile Çalıştırma
```bash
cd rapor_sistemi
python multi_pano_gui.py
```

## 🔧 Derleme

### Gereksinimler
- Python 3.10+
- PyInstaller
- customtkinter
- python-docx
- openpyxl
- pypdf
- pillow

### Derleme Adımları
```bash
# Paketleri yükle
pip install customtkinter python-docx openpyxl pypdf pyinstaller pillow

# Derle
cd rapor_sistemi
pyinstaller RaporSistemi.spec --clean --noconfirm
```

veya `DERLE.bat` dosyasını çalıştırın.

## ⚙️ Yapılandırma

`config/system_config.json` dosyasında şablon dosya yolunu belirtin:
```json
{
  "template_dosya": "Elektrik Tesisatı Gözle Kontrol ve Fonksyion Testleri Periyodik Kontrol Raporu.docx"
}
```

Alternatif olarak `RAPOR_TEMPLATE_PATH` ortam değişkenini kullanabilirsiniz.

## 📊 Veri Akışı

```
┌─────────────────┐     ┌──────────────────┐     ┌───────────────┐
│ GUI Veri Girişi │────▶│ Geçici Excel     │────▶│ DOCX Raporu   │
│ (multi_pano_gui)│     │ (openpyxl)       │     │ (docx_writer) │
└─────────────────┘     └──────────────────┘     └───────────────┘
        │                        │
        ▼                        ▼
┌─────────────────┐     ┌──────────────────┐
│ PDF Sözleşme    │     │ Fluke DOCX       │
│ (sozlesme_parser)     │ (fluke_extractor)│
└─────────────────┘     └──────────────────┘
```

## 📝 Versiyon Geçmişi

### v2.1 (2025-12-15)
- ✅ DOCX şablon dosyası EXE içine gömüldü (portable)
- EXE herhangi bir dizine kopyalanıp çalıştırılabilir

### v2.0 (2025-12-15)
- Çoklu pano desteği eklendi
- CustomTkinter modern arayüz
- PDF sözleşme okuyucu
- Otomatik Iz hesaplama
- KAKR grubu desteği
- PyInstaller ile tek EXE derleme

## ⚠️ Notlar

- ✅ DOCX şablon dosyası EXE içine gömülüdür - ek dosya gerekmez!
- EXE dosyası tek başına çalışabilir (portable)
- Termal görüntüler için Fluke formatında DOCX dosyaları kullanılmalıdır

## 📞 Destek

Sorunlar için GitHub Issues kullanın veya geliştiriciyle iletişime geçin.

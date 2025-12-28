/**
 * DataModels.h
 *
 * Tüm veri yapıları (struct) tanımları.
 * Python tarafındaki dict yapılarının C++ karşılıkları.
 */

#ifndef DATAMODELS_H
#define DATAMODELS_H

#include <QDate>
#include <QDateTime>
#include <QFile>
#include <QMap>
#include <QPair>
#include <QString>
#include <QStringList>
#include <QVariant>
#include <QVector>

namespace RaporSistemi {

/**
 * Firma Bilgileri
 * Python karşılığı: FirmaBilgileri dict
 */
struct FirmaBilgileri {
  QString firmaAdi;
  QString kontrolAdresi;
  QString sgkSicil;
  QString raporNumarasi;
  QDate raporTarihi;
  QString sozlesmeId;
  QDateTime baslangicTarihSaat;
  QDateTime bitisTarihSaat;
  QDate birSonrakiKontrol;

  // Kontrol eden kişi bilgileri
  QString kontrolEdenAdSoyad;
  QString kontrolEdenTc;
  QString pkNo;           // Periyodik kontrol numarası
  QString teklifNumarasi; // tklf placeholder için
  int tesisatSayisi = 1;

  // Cihaz bilgileri (kisi_bilgileri.xlsx'den)
  // Termal Kamera
  QString termalCihazAdi;
  QString termalKalibrasyonTarihi;
  QString termalKalibrasyonGecerlilik;
  QString termalSeriNo;
  QString termalKalibrasyonNo;

  // Ölçüm Cihazı
  QString olcumCihazAdi;
  QString olcumKalibrasyonTarihi;
  QString olcumKalibrasyonGecerlilik;
  QString olcumSeriNo;
  QString olcumKalibrasyonNo;

  // Eski alanlar (geriye dönük uyumluluk için mapping)
  // cihaz1 = Termal Kamera
  QString cihaz1Adi;
  QString cihaz1SeriNo;
  QString cihaz1KalibrasyonTarihi;
  QString cihaz1KalibrasyonGecerlilik;
  QString cihaz1KalibrasyonNo;

  // cihaz2 = Ölçüm Cihazı
  QString cihaz2Adi;
  QString cihaz2SeriNo;
  QString cihaz2KalibrasyonTarihi;
  QString cihaz2KalibrasyonGecerlilik;
  QString cihaz2KalibrasyonNo;

  bool isValid() const {
    return !firmaAdi.isEmpty() && !raporNumarasi.isEmpty();
  }
};

/**
 * Ana Dağıtım Pano Bilgileri
 * 2.1 DETAY BİLGİLER bölümü
 */
struct AnaDagitimPano {
  QString enerjiSaglayan;    // "TEİAŞ Genel Müdürlüğü" vb.
  QString sebekeTipi;        // "TN-S", "TN-C", "TT" vb.
  QString trafoGucu;         // "1000 kVA" vb.
  int sistemGerilimi = 400;  // V
  int sistemFrekans = 50;    // Hz
  QString topraklamaDirenci; // "< 2Ω" vb.
  QString sigortaTipiAna;    // "Kompakt Şalter" vb.
  int nominalAkimAna = 0;    // Ana sigorta In
  QString rcdBilgisi;        // "30mA" vb.
  QString rcdTestBilgisi;

  // Loop empedansı
  QString loopPeN; // L-PE arası
  QString loopLN;  // L-N arası
  QString ik3;     // 3 faz kısa devre akımı

  // YENI ALANLAR (Template Düzeltmesi)
  QString rcdAnmaAkimi;           // RCD_dayanim
  QString hataAkimi;              // I_f
  QString distCevrimEmpedansi;    // Z_E
  QString sistemTopraklamaKesiti; // Sistem_top
  QString anaEspotansiyelKesiti;  // Ana_top

  // Parafudr (SPD) bilgileri
  QString parafudrTip;  // PARAFUDR_TIP
  QString parafudrImax; // PARAFUDR_Imax

  // Potansiyel dengeleme
  QString enBuyukTopKesit; // en_buyuk_top_kesit

  // Zemin izolasyonu
  QString zeminIzoUygunluk; // zo_uygunluk

  // YENI: Proje ve Şema Durumu
  bool projeVarMi = false;
  bool tekHatSemasiVarMi = false;

  // Gerilim ölçümleri
  QString ln;  // L-N (V) gerilimi
  QString npe; // N-PE (V) gerilimi
};

/**
 * Fonksiyon Testi Satırı
 * Her bir linye için test verileri
 * Python'daki FT_COLUMNS ile birebir aynı
 */
struct FonksiyonTesti {
  int siraNo = 0;
  QString linye;             // Linye Adı
  QString sigortaTipi;       // Açma Eğrisi: "B", "C", "D", "K", "Z", "AAA"
  int kutupSayisi = 1;       // Kutup: 1, 2, 3, 4
  int nominalAkim = 0;       // In (A)
  QString icu;               // Icu (kA): "3kA", "6kA", "10kA"
  QString ib;                // Ib = In * 0.7 (otomatik hesaplanır)
  QString fazKesiti;         // Faz Kesiti (mm²)
  QString notrKesiti;        // Nötr Kesiti (mm²)
  QString toprakKesiti;      // Toprak (PE) Kesiti (mm²)
  int akimKapasitesi = 0;    // Iz (A) - otomatik hesaplanır
  QString sonuc;             // "Uygun" / "Uygun Değil"
  bool kakrVar = false;      // KAKR Var checkbox
  QString rcd;               // IΔn: "30mA", "300mA", ""
  QString rcdMa;             // RCD mA test değeri
  QString rcdMs;             // RCD mS test zamanı
  bool kakrYok = false;      // KAKR Yok checkbox
  QString aciklama;          // Kusur açıklaması
  bool isAnaSigorta = false; // ANA SİGORTA satırı (32A KAKR kuralından muaf)

  // Eski alanlar (uyumluluk için)
  QString peN;
  QString lN;

  bool isValid() const { return !linye.isEmpty() && nominalAkim > 0; }

  // In > Iz kontrolü (KAKR grupları için geçerli değil - Python'daki gibi)
  bool isInGreaterThanIz() const {
    return nominalAkim > akimKapasitesi && akimKapasitesi > 0;
  }

  // KAKR grubu kontrolü (Python: linye_adi.upper().contains("KAKR"))
  bool isKakrGroup() const { return linye.toUpper().contains("KAKR"); }

  // 32A altı KAKR kontrolü (ANA SİGORTA muaf - Python'daki gibi)
  bool needs30maKakr() const {
    // ANA SİGORTA için 32A kuralı geçerli değil (Python:
    // multi_pano_gui.py:498-499)
    if (isAnaSigorta)
      return false;
    return nominalAkim <= 32 && rcd != "30mA" && !kakrYok && !kakrVar;
  }
};

/**
 * Termal Görüntü
 */
struct TermalGoruntu {
  QString imagePath;
  QString tip;        // "termal", "proje", "visible", "fluke"
  QString flukeNo;    // Fluke cihaz numarası
  QString fotoTarihi; // Fotoğraf tarihi (Fluke'dan - GK_24)
  QString fotoNo;     // Fotoğraf numarası (Fluke'dan - GK_25)
  int siraNo = 0;

  bool isValid() const {
    return !imagePath.isEmpty() && QFile::exists(imagePath);
  }
};

/**
 * Gözle Kontrol Maddesi
 */
struct GozleKontrolMaddesi {
  int maddeNo = 0;
  QString maddeAdi;
  QString sonuc;         // "Uygun", "Uygun Değil", "Uygulanamaz"
  QString kusurDerecesi; // "Yok", "K1", "K2", "K3"
  QString aciklama;
};

/**
 * Kusur Bilgisi
 */
struct Kusur {
  QString linye;
  QString kusurAciklamasi;
  QString kusurDerecesi; // "K1", "K2", "K3"
};

/**
 * Tek bir pano için tüm veriler
 */
struct PanoData {
  int panoIndex = 0;
  QString panoAdi;
  QString raporNumarasi; // Bu panoya özgü rapor no

  AnaDagitimPano anaDagitimPano;
  QVector<FonksiyonTesti> fonksiyonTestleri;
  QVector<TermalGoruntu> termalGoruntuler;
  QVector<GozleKontrolMaddesi> gozleKontrol;
  QVector<Kusur> kusurlar;

  // Zemin izolasyonu
  QString zeminEn;
  QString zeminBoy;
  QString izoDirenci;
  QString izoUygunluk;

  // Potansiyel dengeleme
  QString enBuyukTopKesit;
  QString potansiyelSonuc = "UYGUN";

  // Sonuç
  QString genelSonuc; // "Uygun" / "Uygun Değil"
  QString aciklama;

  int fonksiyonTestiSayisi() const { return fonksiyonTestleri.size(); }
};

/**
 * Proje - Tüm panoları ve ortak bilgileri içerir
 */
struct Proje {
  FirmaBilgileri firmaBilgileri;
  QVector<PanoData> panolar;

  // Proje meta bilgileri
  QString projeYolu; // Kaydedilen dosya yolu
  QDateTime sonKaydedilme;
  bool degisiklikVar = false;

  // Global Ana Dağıtım Pano Bilgileri (Sidebar'dan)
  AnaDagitimPano anaPanoBilgileri;

  int panoSayisi() const { return panolar.size(); }

  bool isValid() const {
    return firmaBilgileri.isValid() && !panolar.isEmpty();
  }
};

/**
 * Kesit -> Iz (Akım Kapasitesi) Tablosu
 * PVC izoleli bakır iletken, hava ortamı
 */
inline int kesitToIz(double kesit, int carpan = 1) {
  static const QMap<double, int> tablo = {
      {1.0, 16},  {1.5, 20},  {2.5, 27},  {4, 36},    {6, 47},
      {10, 65},   {16, 87},   {25, 115},  {35, 143},  {50, 178},
      {70, 220},  {95, 265},  {120, 310}, {150, 355}, {185, 400},
      {240, 480}, {300, 555}, {400, 770}, {500, 880}};

  int baseIz = 0;
  if (tablo.contains(kesit)) {
    baseIz = tablo[kesit];
  } else {
    // Tam eşleşme yoksa bir üst kesiti bul
    for (auto it = tablo.begin(); it != tablo.end(); ++it) {
      if (kesit <= it.key()) {
        baseIz = it.value();
        break;
      }
    }
    // Hala bulunamadıysa (çok büyükse) son değeri kullan
    if (baseIz == 0 && !tablo.isEmpty()) {
      baseIz = tablo.last();
    }
  }

  return baseIz * carpan;
}

/**
 * Kesit string'ini parse et (örn: "2x16" -> kesit=16, carpan=2)
 */
inline QPair<double, int> parseKesit(const QString &kesitStr) {
  QString val = kesitStr.toLower().replace(',', '.').trimmed();

  if (val.contains('x')) {
    QStringList parts = val.split('x');
    if (parts.size() >= 2) {
      // En son parça her zaman kesittir (örn: 2x3x150 için 150)
      // Ancak bizim formatımız genelde "NxS" veya "S" şeklinde.
      // karmaşık formatlar için (2x(3x150)) regex gerekir ama şimdilik "NxS"
      // varsayalım. Eğer parça sayısı > 2 ise (örn 3x50+25) mantık değişebilir
      // ama şimdilik basit "NxS" formatını destekliyoruz.
      int carpan = parts[0].toInt();
      if (carpan <= 0)
        carpan = 1;

      double kesit = parts[1].toDouble();
      return {kesit, carpan};
    }
  }
  return {val.toDouble(), 1};
}

} // namespace RaporSistemi

#endif // DATAMODELS_H

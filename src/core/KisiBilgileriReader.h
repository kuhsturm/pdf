/**
 * KisiBilgileriReader.h
 *
 * Kişi ve cihaz bilgilerini Excel'den okuyan modül.
 *
 * Python karşılığı: kisi_bilgileri_reader.py
 */

#ifndef KISIBILGILERIREADER_H
#define KISIBILGILERIREADER_H

#include <QString>
#include <QMap>
#include <QVector>
#include <memory>
#include "DataModels.h"

namespace RaporSistemi {

/**
 * Kişi bilgileri yapısı
 */
struct KisiBilgisi {
    QString adSoyad;
    QString tcNo;

    // Termal
    QString termalCihazAdi;
    QString termalKalibrasyonTarihi;
    QString termalKalibrasyonGecerlilik;
    QString termalSeriNo;
    QString termalKalibrasyonNo;

    // Ölçüm
    QString olcumCihazAdi;
    QString olcumKalibrasyonTarihi;
    QString olcumKalibrasyonGecerlilik;
    QString olcumSeriNo;
    QString olcumKalibrasyonNo;

    // Eski alanlar (geriye dönük uyumluluk)
    QString cihaz1Adi;
    QString cihaz1SeriNo;
    QString cihaz1Kalibrasyon;
    QString cihaz2Adi;
    QString cihaz2SeriNo;
    QString cihaz2Kalibrasyon;
    QString cihaz3Adi;
    QString cihaz3SeriNo;
    QString cihaz3Kalibrasyon;

    bool isValid() const {
        return !adSoyad.isEmpty();
    }
};

class KisiBilgileriReader {
public:
    KisiBilgileriReader();
    ~KisiBilgileriReader();

    /**
     * Excel dosyasını yükler.
     * @param excelPath kisi_bilgileri.xlsx dosya yolu
     * @return Başarılı ise true
     */
    bool load(const QString& excelPath);

    /**
     * Tüm kişi isimlerini döndürür.
     */
    QStringList getPersonList() const;

    /**
     * İsme göre kişi bilgilerini döndürür.
     * Fuzzy matching yapar (Türkçe karakter ve büyük/küçük harf duyarsız).
     */
    KisiBilgisi getPersonByName(const QString& name) const;

    /**
     * İsme göre cihaz bilgilerini döndürür.
     * Sözleşmeden gelen kontrol_eden_adsoyad ile eşleştirmek için kullanılır.
     */
    void fillCihazBilgileri(const QString& name, FirmaBilgileri& firma) const;

    /**
     * İsme göre TC numarasını döndürür.
     */
    QString getTcNo(const QString& name) const;

    /**
     * Hata mesajı.
     */
    QString errorString() const { return m_errorString; }

private:
    /**
     * İsmi normalize eder (büyük harf, Türkçe karakter).
     */
    QString normalizeName(const QString& name) const;

    QMap<QString, KisiBilgisi> m_persons;  // Normalize edilmiş isim -> bilgi
    QString m_errorString;
    bool m_loaded = false;
};

/**
 * Sözleşme verisinden cihaz bilgilerini alır.
 */
void getCihazFromSozlesme(const QString& kisiExcelPath,
                          const QString& kontrolEdenAdSoyad,
                          FirmaBilgileri& firma);

} // namespace RaporSistemi

#endif // KISIBILGILERIREADER_H

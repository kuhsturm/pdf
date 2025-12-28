/**
 * ExcelReader.h
 *
 * Excel dosyalarını okuma modülü.
 * QXlsx kütüphanesi kullanır.
 *
 * Python karşılığı: excel_reader.py
 */

#ifndef EXCELREADER_H
#define EXCELREADER_H

#include <QString>
#include <QVector>
#include <memory>
#include "DataModels.h"

// Forward declaration
namespace QXlsx {
    class Document;
}

namespace RaporSistemi {

class ExcelReader {
public:
    ExcelReader();
    ~ExcelReader();

    /**
     * Excel dosyasını yükler.
     * @param path Excel dosyasının yolu
     * @return Başarılı ise true
     */
    bool load(const QString& path);

    /**
     * Dosyayı kapatır.
     */
    void close();

    /**
     * Sayfa isimlerini döndürür.
     */
    QStringList getSheetNames() const;

    /**
     * FirmaBilgileri sayfasını okur.
     */
    FirmaBilgileri readFirmaBilgileri();

    /**
     * AnaDagitimPano sayfasını okur.
     */
    AnaDagitimPano readAnaDagitimPano();

    /**
     * CihazBilgileri sayfasını okur.
     */
    void readCihazBilgileri(FirmaBilgileri& firma);

    /**
     * GozleKontrol sayfasını okur.
     */
    QVector<GozleKontrolMaddesi> readGozleKontrol();

    /**
     * FonksiyonTestleri sayfasını okur.
     * @param anaPano Ana pano verileri (Iz hesaplaması için)
     */
    QVector<FonksiyonTesti> readFonksiyonTestleri(const AnaDagitimPano& anaPano = {});

    /**
     * TermalGoruntuler sayfasını okur.
     */
    QVector<TermalGoruntu> readTermalGoruntuler();

    /**
     * Sonuc sayfasını okur.
     */
    QPair<QString, QString> readSonuc();  // (sonuc, aciklama)

    /**
     * Kusurlar sayfasını okur.
     */
    QVector<Kusur> readKusurlar();

    /**
     * Tüm verileri tek seferde okur.
     */
    PanoData readAll();

    /**
     * Dosya yüklendi mi?
     */
    bool isLoaded() const { return m_loaded; }

    /**
     * Hata mesajı.
     */
    QString errorString() const { return m_errorString; }

private:
    /**
     * Hücre değerini string olarak okur.
     */
    QString cellString(int row, int col) const;

    /**
     * Hücre değerini int olarak okur.
     */
    int cellInt(int row, int col) const;

    /**
     * Hücre değerini double olarak okur.
     */
    double cellDouble(int row, int col) const;

    /**
     * Hücre değerini tarih olarak okur.
     */
    QDate cellDate(int row, int col) const;

    /**
     * Aktif sayfayı değiştirir.
     */
    bool selectSheet(const QString& sheetName);

    std::unique_ptr<QXlsx::Document> m_document;
    bool m_loaded = false;
    QString m_errorString;
    QString m_currentSheet;
};

} // namespace RaporSistemi

#endif // EXCELREADER_H

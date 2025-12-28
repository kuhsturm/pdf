/**
 * DocxWriter.h
 *
 * DOCX rapor yazma modülü.
 * Manuel XML manipülasyonu ile Word belgesi oluşturur.
 *
 * Python karşılığı: docx_writer.py (1200 satır)
 *
 * DOCX dosyası aslında bir ZIP dosyasıdır:
 * - word/document.xml  : Ana içerik
 * - word/styles.xml    : Stiller
 * - word/media/        : Resimler
 * - [Content_Types].xml
 * - word/_rels/document.xml.rels
 */

#ifndef DOCXWRITER_H
#define DOCXWRITER_H

#include <QString>
#include <QVector>
#include <QDomDocument>
#include <QByteArray>
#include <QTemporaryDir>
#include <memory>
#include "DataModels.h"

namespace RaporSistemi {

class DocxWriter {
public:
    DocxWriter();
    ~DocxWriter();

    /**
     * Şablon DOCX dosyasını yükler.
     * @param templatePath Şablon dosyası yolu
     * @return Başarılı ise true
     */
    bool loadTemplate(const QString& templatePath);

    /**
     * Firma bilgilerini doldurur (Tablo 1).
     */
    void fillFirmaBilgileri(const FirmaBilgileri& firma);

    /**
     * Ana dağıtım pano bilgilerini doldurur.
     */
    void fillAnaDagitimPano(const AnaDagitimPano& pano, const QString& panoAdi);

    /**
     * Cihaz bilgilerini doldurur.
     */
    void fillCihazBilgileri(const FirmaBilgileri& firma);

    /**
     * Gözle kontrol bölümünü doldurur (Tablo 2).
     */
    void fillGozleKontrol(const QVector<GozleKontrolMaddesi>& maddeler);

    /**
     * Fluke'dan gelen fotoğraf bilgilerini doldurur (GK_24, GK_25).
     * @param fotoTarihi Fotoğraf tarihi
     * @param fotoNo Fotoğraf numarası
     */
    void fillFlukeBilgileri(const QString& fotoTarihi, const QString& fotoNo);

    /**
     * Fonksiyon testleri bölümünü doldurur (Tablo 3).
     * @param testler Test verileri
     * @param anaPano Ana pano verileri (placeholder'lar için)
     * @param panoAdi Pano adı (PANO_adi1 placeholder için)
     */
    void fillFonksiyonTestleri(const QVector<FonksiyonTesti>& testler,
                               const AnaDagitimPano& anaPano = {},
                               const QString& panoAdi = QString());

    /**
     * Potansiyel dengeleme ve zemin izolasyonu bölümlerini doldurur.
     * @param pano Pano verileri
     * @param fonksiyonTestleri Fonksiyon testleri (en büyük kesit hesabı için)
     * @param zeminIzolasyonDurum Zemin izolasyonu durumu (Uygun/Uygun Değil/Uygulanamaz)
     */
    void fillPotansiyelDengelemeVeZemin(const PanoData& pano,
                                         const QVector<FonksiyonTesti>& fonksiyonTestleri = {},
                                         const QString& zeminIzolasyonDurum = "Uygun");

    /**
     * Termal görüntüleri ekler (Tablo 5).
     */
    void addThermalImages(const QVector<TermalGoruntu>& images);

    /**
     * Kusur açıklaması oluşturur.
     */
    QString generateKusurAciklamasi(const QVector<Kusur>& kusurlar);

    /**
     * Sonuç ve kanaat bölümünü doldurur.
     */
    void fillSonuc(const QString& sonuc, const QString& aciklama);

    /**
     * Uygunluk durumunu doldurur.
     */
    void fillUygunluk(const QString& uygunlukDurumu);

    /**
     * Sayfa sonu ekler.
     */
    void addPageBreakBeforeTable(int tableIndex);

    /**
     * Belgeyi kaydeder.
     * @param outputPath Çıktı dosyası yolu
     * @return Başarılı ise true
     */
    bool save(const QString& outputPath);

    /**
     * Hata mesajı.
     */
    QString errorString() const { return m_errorString; }

private:
    /**
     * Şablonu geçici dizine çıkarır.
     */
    bool extractTemplate(const QString& templatePath);

    /**
     * document.xml dosyasını yükler.
     */
    bool loadDocumentXml();

    /**
     * document.xml dosyasını kaydeder.
     */
    bool saveDocumentXml();

    /**
     * Placeholder'ı değerle değiştirir.
     * @param placeholder Aranacak metin (örn: "{{firma_adi}}")
     * @param value Yerine konacak değer
     */
    void replacePlaceholder(const QString& placeholder, const QString& value);

    /**
     * Tablo hücresindeki placeholder'ı değiştirir.
     * @param tableIndex Tablo indeksi (0-based)
     * @param rowIndex Satır indeksi
     * @param colIndex Sütun indeksi
     * @param value Yeni değer
     */
    void setTableCell(int tableIndex, int rowIndex, int colIndex,
                      const QString& value, int fontSize = 9);

    /**
     * Tabloya yeni satır ekler (mevcut satırı kopyalayarak).
     * @param tableIndex Tablo indeksi
     * @param sourceRowIndex Kopyalanacak satır
     * @return Yeni satırın indeksi
     */
    int copyTableRow(int tableIndex, int sourceRowIndex);

    /**
     * Tablodaki tüm tabloları bulur.
     */
    QList<QDomElement> findTables();

    /**
     * Belgeye resim ekler.
     * @param imagePath Resim dosyası yolu
     * @return Resim ID'si (rId)
     */
    QString addImage(const QString& imagePath);

    /**
     * Belgeyi ZIP olarak paketler.
     */
    bool createZip(const QString& outputPath);

    std::unique_ptr<QTemporaryDir> m_tempDir;
    QDomDocument m_document;
    QString m_xmlContent;  // Orijinal XML içeriği (string replacement için)
    QByteArray m_xmlRawData;  // Orijinal raw bytes (yazma için)
    QString m_errorString;
    int m_imageCounter = 0;
    QStringList m_addedImages;  // Eklenen resim dosya yolları
};

} // namespace RaporSistemi

#endif // DOCXWRITER_H

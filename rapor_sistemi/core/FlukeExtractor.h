/**
 * FlukeExtractor.h
 *
 * Fluke termal kamera DOCX raporlarından görüntü ve veri çıkarıcı.
 *
 * Python karşılığı: fluke_extractor.py
 */

#ifndef FLUKEEXTRACTOR_H
#define FLUKEEXTRACTOR_H

#include <QString>
#include <QStringList>
#include <QMap>
#include <memory>

namespace RaporSistemi {

/**
 * Fluke termal veri yapısı
 */
struct FlukeThermalData {
    QStringList imagePaths;       // Çıkarılan görüntü yolları
    QString flukeNo;              // Cihaz numarası

    // Sıcaklık verileri
    double maxTemp = 0.0;
    double minTemp = 0.0;
    double avgTemp = 0.0;
    double emissivity = 0.95;

    // Cihaz bilgileri
    QString deviceModel;
    QString serialNumber;
    QString captureDate;

    // Fotoğraf bilgileri (GK_24, GK_25 için)
    QString fotoTarihi;           // Fotoğraf tarihi (GK_24)
    QString fotoNo;               // Fotoğraf numarası (GK_25)

    bool isValid() const {
        return !imagePaths.isEmpty();
    }
};

class FlukeExtractor {
public:
    FlukeExtractor();
    ~FlukeExtractor();

    /**
     * Fluke DOCX dosyasını yükler.
     * @param docxPath DOCX dosya yolu
     * @return Başarılı ise true
     */
    bool load(const QString& docxPath);

    /**
     * Görüntüleri çıkarır.
     * @param outputDir Çıktı dizini
     * @param onlyLastTwo Sadece son 2 görüntüyü al (visible + infrared)
     * @return Çıkarılan görüntü yolları
     */
    QStringList extractImages(const QString& outputDir, bool onlyLastTwo = true);

    /**
     * Sıcaklık verilerini çıkarır.
     */
    QMap<QString, double> extractTemperatureData();

    /**
     * Cihaz bilgilerini çıkarır.
     */
    QMap<QString, QString> extractDeviceInfo();

    /**
     * Tüm verileri çıkarır.
     * @param outputDir Görüntülerin kaydedileceği dizin
     */
    FlukeThermalData extractAll(const QString& outputDir);

    /**
     * Hata mesajı.
     */
    QString errorString() const { return m_errorString; }

private:
    /**
     * DOCX'i geçici dizine çıkarır.
     */
    bool extractDocx(const QString& docxPath);

    /**
     * document.xml'den metin çıkarır.
     */
    QString extractDocumentText();

    class Impl;
    std::unique_ptr<Impl> m_impl;
    QString m_errorString;
};

} // namespace RaporSistemi

#endif // FLUKEEXTRACTOR_H

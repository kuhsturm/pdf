/**
 * ReportGenerator.h
 *
 * Rapor oluşturma modülü.
 */

#ifndef REPORTGENERATOR_H
#define REPORTGENERATOR_H

#include <QString>
#include <QVector>
#include "DataModels.h"

namespace RaporSistemi {

class ReportGenerator {
public:
    ReportGenerator();
    ~ReportGenerator();

    /**
     * Şablon yolunu ayarlar.
     */
    void setTemplatePath(const QString& path);

    /**
     * Çıktı dizinini ayarlar.
     */
    void setOutputDirectory(const QString& dir);

    /**
     * Tek bir pano için rapor oluşturur.
     * @return Oluşturulan rapor dosyasının yolu
     */
    QString generateReport(const FirmaBilgileri& firma, const PanoData& pano);

    /**
     * Tüm panolar için rapor oluşturur.
     * @return Oluşturulan rapor dosyalarının yolları
     */
    QStringList generateAllReports(const Proje& proje);

    /**
     * Hata mesajı.
     */
    QString errorString() const { return m_errorString; }

private:
    QString m_templatePath;
    QString m_outputDirectory;
    QString m_errorString;
};

} // namespace RaporSistemi

#endif // REPORTGENERATOR_H

/**
 * FlukeExtractor.cpp
 */

#include "FlukeExtractor.h"
#include <QFile>
#include <QDir>
#include <QTemporaryDir>
#include <QDomDocument>
#include <QRegularExpression>
#include <QDebug>
#include <QtCore/private/qzipreader_p.h>

namespace RaporSistemi {

class FlukeExtractor::Impl {
public:
    std::unique_ptr<QTemporaryDir> tempDir;
    QString docxPath;
    QDomDocument document;
    bool loaded = false;
};

FlukeExtractor::FlukeExtractor() : m_impl(std::make_unique<Impl>()) {}

FlukeExtractor::~FlukeExtractor() = default;

bool FlukeExtractor::load(const QString& docxPath) {
    if (!QFile::exists(docxPath)) {
        m_errorString = QString("Dosya bulunamadı: %1").arg(docxPath);
        return false;
    }

    m_impl->docxPath = docxPath;

    if (!extractDocx(docxPath)) {
        return false;
    }

    m_impl->loaded = true;
    return true;
}

bool FlukeExtractor::extractDocx(const QString& docxPath) {
    m_impl->tempDir = std::make_unique<QTemporaryDir>();
    if (!m_impl->tempDir->isValid()) {
        m_errorString = "Geçici dizin oluşturulamadı";
        return false;
    }

    QZipReader zip(docxPath);
    if (!zip.isReadable()) {
        m_errorString = QString("DOCX okunamadı: %1").arg(docxPath);
        return false;
    }

    return zip.extractAll(m_impl->tempDir->path());
}

QStringList FlukeExtractor::extractImages(const QString& outputDir, bool onlyLastTwo) {
    QStringList imagePaths;

    if (!m_impl->loaded) {
        return imagePaths;
    }

    // word/media/ dizinindeki görüntüleri bul
    QString mediaDir = m_impl->tempDir->path() + "/word/media";
    QDir media(mediaDir);

    if (!media.exists()) {
        return imagePaths;
    }

    // Çıktı dizinini oluştur
    QDir().mkpath(outputDir);

    // Görüntü dosyalarını listele
    QStringList filters = {"*.png", "*.jpg", "*.jpeg", "*.wmf", "*.emf"};
    QStringList images = media.entryList(filters, QDir::Files, QDir::Name);

    // Sadece son 2 görüntü isteniyorsa
    if (onlyLastTwo && images.size() > 2) {
        images = images.mid(images.size() - 2);
    }

    // Görüntüleri kopyala
    for (const QString& imageName : images) {
        QString srcPath = mediaDir + "/" + imageName;
        QString destPath = outputDir + "/" + imageName;

        if (QFile::copy(srcPath, destPath)) {
            imagePaths.append(destPath);
        }
    }

    return imagePaths;
}

QString FlukeExtractor::extractDocumentText() {
    if (!m_impl->loaded) return {};

    QString docPath = m_impl->tempDir->path() + "/word/document.xml";
    QFile file(docPath);

    if (!file.open(QIODevice::ReadOnly)) {
        return {};
    }

    if (!m_impl->document.setContent(&file)) {
        file.close();
        return {};
    }
    file.close();

    // Tüm text node'larını topla
    QString text;
    QDomNodeList textNodes = m_impl->document.elementsByTagName("w:t");

    for (int i = 0; i < textNodes.count(); ++i) {
        text += textNodes.at(i).toElement().text() + " ";
    }

    return text;
}

QMap<QString, double> FlukeExtractor::extractTemperatureData() {
    QMap<QString, double> temps;

    QString text = extractDocumentText();
    if (text.isEmpty()) return temps;

    // Max sıcaklık
    QRegularExpression maxPattern(R"((?:Max|Maximum|Maks)[:\s]*([0-9]+[.,][0-9]+)\s*°?C)",
                                  QRegularExpression::CaseInsensitiveOption);
    QRegularExpressionMatch match = maxPattern.match(text);
    if (match.hasMatch()) {
        temps["max"] = match.captured(1).replace(',', '.').toDouble();
    }

    // Min sıcaklık
    QRegularExpression minPattern(R"((?:Min|Minimum)[:\s]*([0-9]+[.,][0-9]+)\s*°?C)",
                                  QRegularExpression::CaseInsensitiveOption);
    match = minPattern.match(text);
    if (match.hasMatch()) {
        temps["min"] = match.captured(1).replace(',', '.').toDouble();
    }

    // Ortalama sıcaklık
    QRegularExpression avgPattern(R"((?:Avg|Average|Ort)[:\s]*([0-9]+[.,][0-9]+)\s*°?C)",
                                  QRegularExpression::CaseInsensitiveOption);
    match = avgPattern.match(text);
    if (match.hasMatch()) {
        temps["avg"] = match.captured(1).replace(',', '.').toDouble();
    }

    // Emisivite
    QRegularExpression emissPattern(R"((?:Emissivity|Emisivite)[:\s]*([0-9]+[.,][0-9]+))",
                                    QRegularExpression::CaseInsensitiveOption);
    match = emissPattern.match(text);
    if (match.hasMatch()) {
        temps["emissivity"] = match.captured(1).replace(',', '.').toDouble();
    }

    return temps;
}

QMap<QString, QString> FlukeExtractor::extractDeviceInfo() {
    QMap<QString, QString> info;

    QString text = extractDocumentText();
    if (text.isEmpty()) return info;

    // Cihaz modeli
    QRegularExpression modelPattern(R"((?:Model|Cihaz|Ekipman)[:\s]*([A-Za-z0-9\-\s]+))",
                                    QRegularExpression::CaseInsensitiveOption);
    QRegularExpressionMatch match = modelPattern.match(text);
    if (match.hasMatch()) {
        info["model"] = match.captured(1).trimmed();
    }

    // Seri numarası
    QRegularExpression serialPattern(R"((?:Serial|Seri|SN)[:\s]*([A-Za-z0-9\-]+))",
                                     QRegularExpression::CaseInsensitiveOption);
    match = serialPattern.match(text);
    if (match.hasMatch()) {
        info["serial"] = match.captured(1).trimmed();
    }

    // Fluke numarası (dosya adından)
    QRegularExpression flukePattern(R"(FLUKE[-_]?([A-Z0-9]+))",
                                    QRegularExpression::CaseInsensitiveOption);
    match = flukePattern.match(m_impl->docxPath);
    if (match.hasMatch()) {
        info["flukeNo"] = match.captured(1);
    }

    // Fotoğraf tarihi (GK_24)
    QRegularExpression fotoTarihPattern(R"((?:Foto(?:ğraf)?\s*tarihi|Photo\s*date|Tarih)[:\s]*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2}|[0-9]{2}[./-][0-9]{2}[./-][0-9]{4}))",
                                        QRegularExpression::CaseInsensitiveOption);
    match = fotoTarihPattern.match(text);
    if (match.hasMatch()) {
        info["fotoTarihi"] = match.captured(1).trimmed();
    }

    // Fotoğraf numarası (GK_25)
    QRegularExpression fotoNoPattern(R"((?:Foto(?:ğraf)?\s*(?:no\.?|numarası)|Photo\s*no\.?)[:\s]*([A-Za-z0-9\-_]+))",
                                     QRegularExpression::CaseInsensitiveOption);
    match = fotoNoPattern.match(text);
    if (match.hasMatch()) {
        info["fotoNo"] = match.captured(1).trimmed();
    }

    return info;
}

FlukeThermalData FlukeExtractor::extractAll(const QString& outputDir) {
    FlukeThermalData data;

    if (!m_impl->loaded) return data;

    // Görüntüleri çıkar
    data.imagePaths = extractImages(outputDir, true);

    // Sıcaklık verileri
    auto temps = extractTemperatureData();
    data.maxTemp = temps.value("max", 0.0);
    data.minTemp = temps.value("min", 0.0);
    data.avgTemp = temps.value("avg", 0.0);
    data.emissivity = temps.value("emissivity", 0.95);

    // Cihaz bilgileri
    auto info = extractDeviceInfo();
    data.deviceModel = info.value("model");
    data.serialNumber = info.value("serial");
    data.flukeNo = info.value("flukeNo");
    data.fotoTarihi = info.value("fotoTarihi");  // GK_24
    data.fotoNo = info.value("fotoNo");          // GK_25

    return data;
}

} // namespace RaporSistemi

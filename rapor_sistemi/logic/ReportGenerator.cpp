/**
 * ReportGenerator.cpp
 */

#include "ReportGenerator.h"
#include "core/DocxWriter.h"
#include "core/ExcelReader.h"
#include "core/FlukeExtractor.h"
#include <QDir>
#include <QDateTime>
#include <QCoreApplication>
#include <QRegularExpression>
#include <QTemporaryDir>

namespace RaporSistemi {

ReportGenerator::ReportGenerator() {
    // Varsayılan şablon yolu
    QString exePath = QCoreApplication::applicationDirPath();
    m_templatePath = exePath + "/sablon/Elektrik Tesisatı Gözle Kontrol ve Fonksyion Testleri Periyodik Kontrol Raporu.docx";
    m_outputDirectory = exePath + "/raporlar";
}

ReportGenerator::~ReportGenerator() = default;

void ReportGenerator::setTemplatePath(const QString& path) {
    m_templatePath = path;
}

void ReportGenerator::setOutputDirectory(const QString& dir) {
    m_outputDirectory = dir;
}

QString ReportGenerator::generateReport(const FirmaBilgileri& firma, const PanoData& pano) {
    // Çıktı dizinini oluştur
    QDir().mkpath(m_outputDirectory);

    // Geçici dizin (Fluke görsellerini çıkarmak için)
    QTemporaryDir tempDir;
    if (!tempDir.isValid()) {
        m_errorString = "Geçici dizin oluşturulamadı";
        return {};
    }

    // DocxWriter ile rapor oluştur
    DocxWriter writer;

    if (!writer.loadTemplate(m_templatePath)) {
        m_errorString = writer.errorString();
        return {};
    }

    // Verileri doldur
    writer.fillFirmaBilgileri(firma);
    writer.fillAnaDagitimPano(pano.anaDagitimPano, pano.panoAdi);
    writer.fillCihazBilgileri(firma);
    writer.fillGozleKontrol(pano.gozleKontrol);  // Gözle kontrol maddelerini doldur

    // Termal görüntüleri işle (Fluke DOCX varsa içlerinden görüntü çıkar)
    QString fotoTarihi, fotoNo;
    QVector<TermalGoruntu> processedImages;

    for (const auto& img : pano.termalGoruntuler) {
        // Fluke DOCX dosyası mı kontrol et
        if (img.imagePath.endsWith(".docx", Qt::CaseInsensitive) &&
            (img.tip == "fluke" || img.tip.isEmpty())) {
            // Fluke extractor ile görüntüleri çıkar
            FlukeExtractor extractor;
            if (extractor.load(img.imagePath)) {
                FlukeThermalData flukeData = extractor.extractAll(tempDir.path());

                // Fluke'dan çıkarılan her görüntü için TermalGoruntu oluştur
                for (const QString& extractedPath : flukeData.imagePaths) {
                    TermalGoruntu extracted;
                    extracted.imagePath = extractedPath;
                    extracted.tip = "termal";
                    extracted.flukeNo = flukeData.flukeNo;
                    extracted.fotoTarihi = flukeData.fotoTarihi;
                    extracted.fotoNo = flukeData.fotoNo;
                    processedImages.append(extracted);
                }

                // İlk Fluke'dan fotoğraf tarihi ve no al
                if (fotoTarihi.isEmpty() && !flukeData.fotoTarihi.isEmpty()) {
                    fotoTarihi = flukeData.fotoTarihi;
                }
                if (fotoNo.isEmpty() && !flukeData.fotoNo.isEmpty()) {
                    fotoNo = flukeData.fotoNo;
                } else if (!flukeData.fotoNo.isEmpty()) {
                    fotoNo += " - " + flukeData.fotoNo;
                }
            }
        } else if (img.tip == "proje_gorseli") {
            // Proje görseli doğrudan ekle
            processedImages.append(img);
        } else if (!img.imagePath.endsWith(".docx", Qt::CaseInsensitive)) {
            // Normal görüntü dosyası (jpg, png vs.) doğrudan ekle
            processedImages.append(img);

            // Kullanıcı manuel girmiş olabilir
            if (fotoTarihi.isEmpty() && !img.fotoTarihi.isEmpty()) {
                fotoTarihi = img.fotoTarihi;
            }
            if (fotoNo.isEmpty() && !img.fotoNo.isEmpty()) {
                fotoNo = img.fotoNo;
            }
        }
    }

    // Fotoğraf bilgilerini doldur (GK_24, GK_25)
    writer.fillFlukeBilgileri(fotoTarihi, fotoNo);

    // Fonksiyon testleri - pano adını da geçir
    writer.fillFonksiyonTestleri(pano.fonksiyonTestleri, pano.anaDagitimPano, pano.panoAdi);

    // Zemin izolasyonu durumunu gözle kontrolden al (GK_06)
    QString zeminIzolasyonDurum = "Uygun";
    for (const auto& gk : pano.gozleKontrol) {
        if (gk.maddeAdi.contains("Zemin", Qt::CaseInsensitive) &&
            gk.maddeAdi.contains("Izolasyon", Qt::CaseInsensitive)) {
            zeminIzolasyonDurum = gk.sonuc;
            break;
        }
    }

    // Potansiyel dengeleme ve zemin izolasyonu - fonksiyon testlerini de geçir
    writer.fillPotansiyelDengelemeVeZemin(pano, pano.fonksiyonTestleri, zeminIzolasyonDurum);

    // İşlenmiş termal görüntüleri rapora ekle
    writer.addThermalImages(processedImages);

    // Sonuç
    QString sonuc = pano.genelSonuc.isEmpty() ? "Uygun" : pano.genelSonuc;
    writer.fillSonuc(sonuc, pano.aciklama);
    writer.fillUygunluk(sonuc);

    // Dosya adı
    QString sanitizedFirma = firma.firmaAdi;
    sanitizedFirma.replace(QRegularExpression("[^a-zA-Z0-9çğıöşüÇĞİÖŞÜ\\s]"), "_");

    QString filename = QString("%1_%2_%3.docx")
        .arg(pano.raporNumarasi.isEmpty() ? firma.raporNumarasi : pano.raporNumarasi)
        .arg(pano.panoAdi.isEmpty() ? QString("Pano%1").arg(pano.panoIndex) : pano.panoAdi)
        .arg(QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss"));

    QString outputPath = m_outputDirectory + "/" + filename;

    if (!writer.save(outputPath)) {
        m_errorString = writer.errorString();
        return {};
    }

    return outputPath;
}

QStringList ReportGenerator::generateAllReports(const Proje& proje) {
    QStringList generatedFiles;

    for (const PanoData& pano : proje.panolar) {
        // Global ana pano bilgilerini yerel pano verisiyle birleştir
        PanoData effectivePano = pano;

        // Sadece boş olan alanları doldur veya globali önceliklendir
        // Şimdilik global veriyi master kabul ediyoruz (kullanıcı oraya giriyor)
        // Ancak parafudr gibi per-pano alanları korumalıyız

        AnaDagitimPano& target = effectivePano.anaDagitimPano;
        const AnaDagitimPano& source = proje.anaPanoBilgileri;

        // Ortak alanları güncelle
        if (!source.enerjiSaglayan.isEmpty()) target.enerjiSaglayan = source.enerjiSaglayan;
        if (!source.sebekeTipi.isEmpty()) target.sebekeTipi = source.sebekeTipi;
        if (!source.topraklamaDirenci.isEmpty()) target.topraklamaDirenci = source.topraklamaDirenci;
        if (!source.distCevrimEmpedansi.isEmpty()) target.distCevrimEmpedansi = source.distCevrimEmpedansi;
        if (!source.sigortaTipiAna.isEmpty()) target.sigortaTipiAna = source.sigortaTipiAna;
        if (source.nominalAkimAna > 0) target.nominalAkimAna = source.nominalAkimAna;
        if (!source.rcdBilgisi.isEmpty()) target.rcdBilgisi = source.rcdBilgisi;
        if (!source.rcdAnmaAkimi.isEmpty()) target.rcdAnmaAkimi = source.rcdAnmaAkimi;
        if (!source.hataAkimi.isEmpty()) target.hataAkimi = source.hataAkimi;
        if (!source.sistemTopraklamaKesiti.isEmpty()) target.sistemTopraklamaKesiti = source.sistemTopraklamaKesiti;
        if (!source.anaEspotansiyelKesiti.isEmpty()) target.anaEspotansiyelKesiti = source.anaEspotansiyelKesiti;
        if (!source.parafudrTip.isEmpty()) target.parafudrTip = source.parafudrTip;
        if (!source.parafudrImax.isEmpty()) target.parafudrImax = source.parafudrImax;
        if (!source.enBuyukTopKesit.isEmpty()) target.enBuyukTopKesit = source.enBuyukTopKesit;
        if (!source.zeminIzoUygunluk.isEmpty()) target.zeminIzoUygunluk = source.zeminIzoUygunluk;
        if (!source.loopPeN.isEmpty()) target.loopPeN = source.loopPeN;
        if (!source.loopLN.isEmpty()) target.loopLN = source.loopLN;
        if (source.sistemGerilimi > 0) target.sistemGerilimi = source.sistemGerilimi;

        QString path = generateReport(proje.firmaBilgileri, effectivePano);
        if (!path.isEmpty()) {
            generatedFiles.append(path);
        }
    }

    return generatedFiles;
}

} // namespace RaporSistemi

/**
 * Test Rapor Oluşturucu
 * Örnek verilerle rapor çıktısı oluşturur
 */

#include "src/core/DocxWriter.h"
#include "src/logic/DataModels.h"
#include "src/logic/ReportGenerator.h"
#include <QCoreApplication>
#include <QDebug>
#include <QDir>

using namespace RaporSistemi;

int main(int argc, char *argv[]) {
    QCoreApplication app(argc, argv);

    qDebug() << "=== Örnek Rapor Oluşturucu ===";

    // Firma bilgileri
    FirmaBilgileri firma;
    firma.firmaAdi = "ABC Elektrik Sanayi A.Ş.";
    firma.kontrolAdresi = "Organize Sanayi Bölgesi 5. Cadde No:42, Nilüfer/BURSA";
    firma.sgkSicil = "1234567890";
    firma.raporNumarasi = "RPR-2025-001";
    firma.raporTarihi = QDate(2025, 12, 27);
    firma.sozlesmeId = "SZL-2025-0456";
    firma.baslangicTarihSaat = QDateTime(QDate(2025, 12, 27), QTime(9, 0));
    firma.bitisTarihSaat = QDateTime(QDate(2025, 12, 27), QTime(17, 30));
    firma.birSonrakiKontrol = QDate(2026, 12, 27);
    firma.kontrolEdenAdSoyad = "Mehmet YILMAZ";
    firma.pkNo = "ELK-2024-12345";
    firma.teklifNumarasi = "TKLF/2025/9999";

    // Termal kamera bilgileri
    firma.termalCihazAdi = "FLUKE TiS60+";
    firma.termalSeriNo = "FLK-TIS60-123456";
    firma.termalKalibrasyonTarihi = "15.06.2025";
    firma.termalKalibrasyonGecerlilik = "15.06.2026";
    firma.termalKalibrasyonNo = "KAL-2025-789";

    // Ölçüm cihazı bilgileri
    firma.olcumCihazAdi = "FLUKE 1664 FC";
    firma.olcumSeriNo = "FLK-1664-987654";
    firma.olcumKalibrasyonTarihi = "20.03.2025";
    firma.olcumKalibrasyonGecerlilik = "20.03.2026";
    firma.olcumKalibrasyonNo = "KAL-2025-456";

    // Pano verileri
    PanoData pano;
    pano.panoAdi = "Ana Dağıtım Panosu";
    pano.panoIndex = 1;
    pano.raporNumarasi = firma.raporNumarasi;

    // Ana dağıtım pano
    pano.anaDagitimPano.enerjiSaglayan = "UEDAŞ - Uludağ Elektrik Dağıtım A.Ş.";
    pano.anaDagitimPano.sebekeTipi = "TN-S";
    pano.anaDagitimPano.trafoGucu = "1000 kVA";
    pano.anaDagitimPano.sistemGerilimi = 400;
    pano.anaDagitimPano.sistemFrekans = 50;
    pano.anaDagitimPano.topraklamaDirenci = "2.5";
    pano.anaDagitimPano.sigortaTipiAna = "Kompakt Şalter";
    pano.anaDagitimPano.nominalAkimAna = 630;
    pano.anaDagitimPano.rcdBilgisi = "Tip A";
    pano.anaDagitimPano.rcdAnmaAkimi = "300 mA";
    pano.anaDagitimPano.distCevrimEmpedansi = "0.5";  // Z_E = 0.5 Ohm -> I_f = 460 A
    pano.anaDagitimPano.sistemTopraklamaKesiti = "70 mm²";
    pano.anaDagitimPano.anaEspotansiyelKesiti = "25 mm²";
    pano.anaDagitimPano.parafudrTip = "Tip 2";
    pano.anaDagitimPano.parafudrImax = "40 kA";
    pano.anaDagitimPano.enBuyukTopKesit = "16 mm²";
    pano.anaDagitimPano.loopPeN = "0.35";  // Z_x
    pano.anaDagitimPano.loopLN = "0.28";   // Z_ln -> 380/0.28 = 1357 A

    // Gözle kontrol maddeleri (29 adet)
    QStringList gkAlanlar = {
        "Kablo Sebeke Tarafi", "Pano Sabitlenmesi",
        "Elektrik Panosu Etrafinda Yabanci Malzemeler", "Kablo Donanim Tarafi",
        "Dis Darbelere Karsi Koruma Onlemi", "Zemin Izolasyonu",
        "Topraklama Iletkeni", "Ek Potansiyel Dengeleme Iletkeni",
        "Ana Potansiyel Dengeleme Iletkeni", "Pano Kapak Baglantisi Kontrolu 6 mm2",
        "Elektriksel Olmayan Tesislere Yaklasma", "Guvenlik Devre Ayrilmasi",
        "Bant Ayrilmasi", "Pano Ic Kapak",
        "Semalar Talimatlar", "Tehlike Isaretleri",
        "Koruma Cihaz ve Terminal Etiket", "Kablo Yollari",
        "Tesisat Yontemi", "Kablo Renk Kodlari",
        "Yangin Engeli", "Kontak Gevsekligi Isinmasi",
        "Asiri Yuk Isinmasi", "Fotograf Tarihi",  // GK_24
        "Fotograf No",  // GK_25
        "Yangin Sondurme", "Korozyon Kontrolu",
        "Ekipman Temizlik", "Acil Durum Aydinlatma"
    };

    for (int i = 0; i < 29; ++i) {
        GozleKontrolMaddesi madde;
        madde.maddeNo = i + 1;
        madde.maddeAdi = (i < gkAlanlar.size()) ? gkAlanlar[i] : "";

        // GK_24 ve GK_25 için özel değerler
        if (i == 23) {  // GK_24 = Fotoğraf Tarihi
            madde.sonuc = "27.12.2025";
        } else if (i == 24) {  // GK_25 = Fotoğraf No
            madde.sonuc = "IMG-001";
        } else {
            madde.sonuc = "Uygun";  // Hepsi uygun
        }
        pano.gozleKontrol.append(madde);
    }

    // Fonksiyon testleri (örnek linyeler)
    QStringList linyeler = {
        "Genel Aydınlatma", "Priz Hattı 1", "Priz Hattı 2",
        "Klima", "Kompresör", "CNC Tezgah 1", "CNC Tezgah 2",
        "Kaynak Makinesi", "Vinç", "Yedek"
    };

    int akimlar[] = {16, 16, 16, 25, 32, 40, 40, 63, 80, 10};
    QString kesitler[] = {"2.5", "2.5", "2.5", "4", "6", "10", "10", "16", "25", "1.5"};

    for (int i = 0; i < 10; ++i) {
        FonksiyonTesti test;
        test.siraNo = i + 1;
        test.linye = linyeler[i];
        test.sigortaTipi = (akimlar[i] <= 16) ? "B" : "C";
        test.kutupSayisi = 3;
        test.nominalAkim = akimlar[i];
        test.icu = "6";
        test.fazKesiti = kesitler[i];
        test.notrKesiti = kesitler[i];
        test.toprakKesiti = (akimlar[i] <= 16) ? "2.5" : kesitler[i];
        test.ib = QString::number(int(akimlar[i] * 0.7));
        test.rcd = (akimlar[i] <= 32) ? "30" : "300";
        test.rcdMs = "25";
        test.sonuc = "Uygun";
        pano.fonksiyonTestleri.append(test);
    }

    // Zemin izolasyonu
    pano.zeminEn = "2";
    pano.zeminBoy = "3";
    pano.izoDirenci = ">50MΩ";
    pano.izoUygunluk = "Uygun";
    pano.enBuyukTopKesit = "16";

    // Genel sonuç
    pano.genelSonuc = "Uygun";
    pano.aciklama = "";

    // Rapor oluştur
    ReportGenerator generator;

    // Şablon yolunu ayarla
    QString exePath = QCoreApplication::applicationDirPath();
    QString templatePath = exePath + "/sablon/rapor_sablonu.docx";

    // Alternatif yolları dene
    QStringList templatePaths = {
        templatePath,
        "d:/YAPAY ZEKALILAR/rapor_sistemi/rapor_sistemi_cpp/sablon/rapor_sablonu.docx",
        "../sablon/rapor_sablonu.docx",
        "sablon/rapor_sablonu.docx"
    };

    bool templateFound = false;
    for (const QString& path : templatePaths) {
        if (QFile::exists(path)) {
            generator.setTemplatePath(path);
            templateFound = true;
            qDebug() << "Şablon bulundu:" << path;
            break;
        }
    }

    if (!templateFound) {
        qCritical() << "HATA: Şablon dosyası bulunamadı!";
        qDebug() << "Aranan yollar:";
        for (const QString& path : templatePaths) {
            qDebug() << "  -" << path;
        }
        return 1;
    }

    // Çıktı dizini
    QString outputDir = "d:/YAPAY ZEKALILAR/rapor_sistemi/rapor_sistemi_cpp/test_output";
    QDir().mkpath(outputDir);
    generator.setOutputDirectory(outputDir);

    qDebug() << "\nRapor oluşturuluyor...";

    QString outputPath = generator.generateReport(firma, pano);

    if (outputPath.isEmpty()) {
        qCritical() << "HATA: Rapor oluşturulamadı!";
        qCritical() << "Hata:" << generator.errorString();
        return 1;
    }

    qDebug() << "\n✓ Rapor başarıyla oluşturuldu!";
    qDebug() << "Dosya:" << outputPath;

    return 0;
}

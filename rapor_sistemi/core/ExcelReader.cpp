/**
 * ExcelReader.cpp
 *
 * Excel okuma implementasyonu.
 */

#include "ExcelReader.h"
#include "xlsxdocument.h"
#include "xlsxworkbook.h"
#include <QFile>
#include <QDebug>

namespace RaporSistemi {

ExcelReader::ExcelReader() = default;

ExcelReader::~ExcelReader() {
    close();
}

bool ExcelReader::load(const QString& path) {
    close();

    if (!QFile::exists(path)) {
        m_errorString = QString("Dosya bulunamadı: %1").arg(path);
        return false;
    }

    m_document = std::make_unique<QXlsx::Document>(path);

    if (!m_document->load()) {
        m_errorString = QString("Excel dosyası açılamadı: %1").arg(path);
        m_document.reset();
        return false;
    }

    m_loaded = true;
    m_errorString.clear();
    return true;
}

void ExcelReader::close() {
    m_document.reset();
    m_loaded = false;
    m_currentSheet.clear();
}

QStringList ExcelReader::getSheetNames() const {
    if (!m_document) return {};
    return m_document->sheetNames();
}

bool ExcelReader::selectSheet(const QString& sheetName) {
    if (!m_document) return false;

    if (m_document->selectSheet(sheetName)) {
        m_currentSheet = sheetName;
        return true;
    }

    m_errorString = QString("Sayfa bulunamadı: %1").arg(sheetName);
    return false;
}

QString ExcelReader::cellString(int row, int col) const {
    if (!m_document) return {};
    QVariant val = m_document->read(row, col);
    return val.toString().trimmed();
}

int ExcelReader::cellInt(int row, int col) const {
    if (!m_document) return 0;
    QVariant val = m_document->read(row, col);
    return val.toInt();
}

double ExcelReader::cellDouble(int row, int col) const {
    if (!m_document) return 0.0;
    QVariant val = m_document->read(row, col);
    return val.toDouble();
}

QDate ExcelReader::cellDate(int row, int col) const {
    if (!m_document) return {};
    QVariant val = m_document->read(row, col);

    if (val.typeId() == QMetaType::QDate) {
        return val.toDate();
    } else if (val.typeId() == QMetaType::QDateTime) {
        return val.toDateTime().date();
    } else if (val.typeId() == QMetaType::QString) {
        // DD.MM.YYYY formatı
        return QDate::fromString(val.toString(), "dd.MM.yyyy");
    }
    return {};
}

FirmaBilgileri ExcelReader::readFirmaBilgileri() {
    FirmaBilgileri firma;

    if (!selectSheet("FirmaBilgileri")) {
        return firma;
    }

    // Excel yapısı: A sütunu = alan adı, B sütunu = değer
    for (int row = 1; row <= 20; ++row) {
        QString key = cellString(row, 1).toLower();
        QString value = cellString(row, 2);

        if (key.contains("firma") && key.contains("ad")) {
            firma.firmaAdi = value;
        } else if (key.contains("adres")) {
            firma.kontrolAdresi = value;
        } else if (key.contains("sgk")) {
            firma.sgkSicil = value;
        } else if (key.contains("rapor") && key.contains("no")) {
            firma.raporNumarasi = value;
        } else if (key.contains("rapor") && key.contains("tarih")) {
            firma.raporTarihi = cellDate(row, 2);
        } else if (key.contains("sozlesme") || key.contains("sözleşme")) {
            firma.sozlesmeId = value;
        }
    }

    return firma;
}

AnaDagitimPano ExcelReader::readAnaDagitimPano() {
    AnaDagitimPano pano;

    if (!selectSheet("AnaDagitimPano")) {
        return pano;
    }

    for (int row = 1; row <= 30; ++row) {
        QString key = cellString(row, 1).toLower();
        QString value = cellString(row, 2);

        if (key.contains("enerji")) {
            pano.enerjiSaglayan = value;
        } else if (key.contains("şebeke") || key.contains("sebeke")) {
            pano.sebekeTipi = value;
        } else if (key.contains("trafo")) {
            pano.trafoGucu = value;
        } else if (key.contains("gerilim")) {
            pano.sistemGerilimi = cellInt(row, 2);
        } else if (key.contains("frekans")) {
            pano.sistemFrekans = cellInt(row, 2);
        } else if (key.contains("topraklama") && key.contains("diren")) {
            pano.topraklamaDirenci = value;
        } else if (key.contains("sigorta") || key.contains("kesici")) {
            pano.sigortaTipiAna = value;
        } else if (key.contains("nominal") && key.contains("ak")) {
            pano.nominalAkimAna = cellInt(row, 2);
        } else if (key.contains("rcd") && key.contains("tip")) {
            pano.rcdBilgisi = value;
        } else if (key.contains("rcd") && key.contains("anma")) {
            pano.rcdAnmaAkimi = value;
        } else if (key.contains("rcd") && key.contains("test")) {
            pano.rcdTestBilgisi = value;
        } else if (key.contains("z_e") || key.contains("empedans")) {
            pano.distCevrimEmpedansi = value;
        } else if (key.contains("i_f") || key.contains("hata") && key.contains("ak")) {
            pano.hataAkimi = value;
        } else if (key.contains("sistem") && key.contains("top")) {
            pano.sistemTopraklamaKesiti = value;
        } else if (key.contains("espotansiyel") || key.contains("eşpotansiyel")) {
            pano.anaEspotansiyelKesiti = value;
        } else if (key.contains("parafudr") && key.contains("tip")) {
            pano.parafudrTip = value;
        } else if (key.contains("parafudr") && key.contains("imax")) {
            pano.parafudrImax = value;
        } else if (key.contains("buyuk") && key.contains("kesit")) {
            pano.enBuyukTopKesit = value;
        } else if (key.contains("zemin") && key.contains("uygun")) {
            pano.zeminIzoUygunluk = value;
        }
    }

    return pano;
}

QVector<FonksiyonTesti> ExcelReader::readFonksiyonTestleri(const AnaDagitimPano& anaPano) {
    QVector<FonksiyonTesti> testler;

    if (!selectSheet("FonksiyonTestleri")) {
        return testler;
    }

    // Excel sütun yapısı (ornek_veri.xlsx):
    // 1: No, 2: Linye Adi, 3: Acma Egrisi, 4: Kutup Sayisi, 5: In (A)
    // 6: Icu (kA), 7: Faz Kesiti, 8: N Kesiti, 9: PE Kesiti
    // 10: Ib (A), 11: Iz (A), 12: RCD mA, 13: RCD ms, 14: Sonuc

    int row = 2;
    while (true) {
        QString linye = cellString(row, 2);  // Col 2: Linye Adi
        if (linye.isEmpty()) break;

        FonksiyonTesti test;
        test.siraNo = cellInt(row, 1);            // Col 1: No
        test.linye = linye;                        // Col 2: Linye Adi
        test.sigortaTipi = cellString(row, 3);    // Col 3: Acma Egrisi
        test.kutupSayisi = cellInt(row, 4);       // Col 4: Kutup Sayisi
        test.nominalAkim = cellInt(row, 5);       // Col 5: In (A)
        test.icu = cellString(row, 6);            // Col 6: Icu (kA)
        test.fazKesiti = cellString(row, 7);      // Col 7: Faz Kesiti
        test.notrKesiti = cellString(row, 8);     // Col 8: N Kesiti
        test.toprakKesiti = cellString(row, 9);   // Col 9: PE Kesiti
        test.ib = cellString(row, 10);            // Col 10: Ib (A)
        test.akimKapasitesi = cellInt(row, 11);   // Col 11: Iz (A)
        test.rcd = cellString(row, 12);           // Col 12: RCD mA
        test.rcdMs = cellString(row, 13);         // Col 13: RCD ms
        test.sonuc = cellString(row, 14);         // Col 14: Sonuc

        // Varsayılan değerler
        if (test.kutupSayisi == 0) test.kutupSayisi = 1;
        if (test.icu.isEmpty()) test.icu = "6";
        if (test.sonuc.isEmpty()) test.sonuc = "Uygun";

        // Iz'i kesit değerinden hesapla (eğer boşsa)
        if (test.akimKapasitesi == 0 && !test.fazKesiti.isEmpty()) {
            auto [kesit, carpan] = parseKesit(test.fazKesiti);
            test.akimKapasitesi = kesitToIz(kesit, carpan);
        }

        // Ib'yi hesapla (eğer boşsa): In * 0.7
        if (test.ib.isEmpty() && test.nominalAkim > 0) {
            test.ib = QString::number(qRound(test.nominalAkim * 0.7));
        }

        testler.append(test);
        ++row;
    }

    return testler;
}

QVector<TermalGoruntu> ExcelReader::readTermalGoruntuler() {
    QVector<TermalGoruntu> goruntuler;

    if (!selectSheet("TermalGoruntuler")) {
        return goruntuler;
    }

    int row = 2;
    while (true) {
        QString path = cellString(row, 1);
        if (path.isEmpty()) break;

        TermalGoruntu goruntu;
        goruntu.siraNo = row - 1;
        goruntu.imagePath = path;
        goruntu.tip = cellString(row, 2);
        goruntu.flukeNo = cellString(row, 3);

        goruntuler.append(goruntu);
        ++row;
    }

    return goruntuler;
}

QVector<GozleKontrolMaddesi> ExcelReader::readGozleKontrol() {
    QVector<GozleKontrolMaddesi> maddeler;

    if (!selectSheet("GozleKontrol")) {
        return maddeler;
    }

    int row = 2;
    while (true) {
        QString madde = cellString(row, 1);
        if (madde.isEmpty()) break;

        GozleKontrolMaddesi m;
        m.maddeNo = row - 1;
        m.maddeAdi = madde;
        m.sonuc = cellString(row, 2);
        m.kusurDerecesi = cellString(row, 3);
        m.aciklama = cellString(row, 4);

        maddeler.append(m);
        ++row;
    }

    return maddeler;
}

QPair<QString, QString> ExcelReader::readSonuc() {
    if (!selectSheet("Sonuc")) {
        return {};
    }

    QString sonuc = cellString(1, 2);
    QString aciklama = cellString(2, 2);

    return {sonuc, aciklama};
}

QVector<Kusur> ExcelReader::readKusurlar() {
    QVector<Kusur> kusurlar;

    if (!selectSheet("Kusurlar")) {
        return kusurlar;
    }

    int row = 2;
    while (true) {
        QString linye = cellString(row, 1);
        if (linye.isEmpty()) break;

        Kusur k;
        k.linye = linye;
        k.kusurAciklamasi = cellString(row, 2);
        k.kusurDerecesi = cellString(row, 3);

        kusurlar.append(k);
        ++row;
    }

    return kusurlar;
}

void ExcelReader::readCihazBilgileri(FirmaBilgileri& firma) {
    if (!selectSheet("CihazBilgileri")) {
        return;
    }

    // Cihaz 1 - Termal Kamera
    firma.cihaz1Adi = cellString(2, 1);
    firma.cihaz1SeriNo = cellString(2, 2);
    firma.cihaz1KalibrasyonTarihi = cellString(2, 3);
    firma.cihaz1KalibrasyonGecerlilik = cellString(2, 4);
    firma.cihaz1KalibrasyonNo = cellString(2, 5);

    // Cihaz 2 - Ölçüm Cihazı
    firma.cihaz2Adi = cellString(3, 1);
    firma.cihaz2SeriNo = cellString(3, 2);
    firma.cihaz2KalibrasyonTarihi = cellString(3, 3);
    firma.cihaz2KalibrasyonGecerlilik = cellString(3, 4);
    firma.cihaz2KalibrasyonNo = cellString(3, 5);
}

PanoData ExcelReader::readAll() {
    PanoData data;

    if (!m_loaded) return data;

    AnaDagitimPano anaPano = readAnaDagitimPano();
    data.anaDagitimPano = anaPano;
    data.fonksiyonTestleri = readFonksiyonTestleri(anaPano);
    data.termalGoruntuler = readTermalGoruntuler();
    data.gozleKontrol = readGozleKontrol();
    data.kusurlar = readKusurlar();

    auto [sonuc, aciklama] = readSonuc();
    data.genelSonuc = sonuc;
    data.aciklama = aciklama;

    return data;
}

} // namespace RaporSistemi

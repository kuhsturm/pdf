/**
 * DocxWriter.cpp
 *
 * DOCX yazma implementasyonu.
 * DOCX = ZIP (XML dosyaları içeren)
 */

#include "DocxWriter.h"
#include <QFile>
#include <QDir>
#include <QDirIterator>
#include <QDebug>
#include <QDomElement>
#include <QDomNodeList>
#include <QBuffer>
#include <QImage>
#include <QImageWriter>
#include <QTextStream>

// MiniZip veya QuaZip kullanılabilir, şimdilik Qt ile basit çözüm
#include <QtCore/private/qzipreader_p.h>
#include <QtCore/private/qzipwriter_p.h>

namespace RaporSistemi {

// Word namespace'leri
static const QString NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
static const QString NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
static const QString NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
static const QString NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
static const QString NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

DocxWriter::DocxWriter() = default;

DocxWriter::~DocxWriter() = default;

bool DocxWriter::loadTemplate(const QString& templatePath) {
    if (!QFile::exists(templatePath)) {
        m_errorString = QString("Şablon dosyası bulunamadı: %1").arg(templatePath);
        return false;
    }

    // Şablonu çıkar
    if (!extractTemplate(templatePath)) {
        return false;
    }

    // document.xml'i yükle
    if (!loadDocumentXml()) {
        return false;
    }

    return true;
}

bool DocxWriter::extractTemplate(const QString& templatePath) {
    m_tempDir = std::make_unique<QTemporaryDir>();
    if (!m_tempDir->isValid()) {
        m_errorString = "Geçici dizin oluşturulamadı";
        return false;
    }

    QZipReader zip(templatePath);
    if (!zip.isReadable()) {
        m_errorString = QString("DOCX dosyası okunamadı (Arşiv bozuk veya erişim yok): %1").arg(templatePath);
        return false;
    }

    // Debug: Dosya içeriğini kontrol et
    QVector<QZipReader::FileInfo> allFiles = zip.fileInfoList();
    qDebug() << "Template Path:" << templatePath;
    qDebug() << "ZIP Entry Count:" << allFiles.count();
    qDebug() << "Temp Path:" << m_tempDir->path();

    if (allFiles.isEmpty()) {
         m_errorString = QString("DOCX arşivi boş görünüyor: %1").arg(templatePath);
         return false;
    }

    // ZIP içeriğini çıkar - MANUEL EXTRACTION (Hata Detayı İçin)
    QDir dir(m_tempDir->path());

    for (const QZipReader::FileInfo &info : allFiles) {
        QString destPath = dir.filePath(info.filePath);

        if (info.isDir) {
            if (!dir.mkpath(info.filePath)) {
                m_errorString = QString("Klasör oluşturulamadı: %1").arg(destPath);
                return false;
            }
        } else {
            // Parent klasörün var olduğundan emin ol
            QFileInfo destInfo(destPath);
            if (!dir.mkpath(destInfo.path())) {
                 m_errorString = QString("Dosya yolu oluşturulamadı: %1").arg(destInfo.path());
                 return false;
            }

            QFile file(destPath);
            if (!file.open(QIODevice::WriteOnly)) {
                m_errorString = QString("Dosya yazılamadı!\nHata: %1\nYol: %2")
                    .arg(file.errorString())
                    .arg(destPath);
                return false;
            }

            file.write(zip.fileData(info.filePath));
            file.close();

            if (file.error() != QFile::NoError) {
                 m_errorString = QString("Dosya yazma hatası: %1").arg(file.errorString());
                 return false;
            }
        }
    }

    return true;
}

bool DocxWriter::loadDocumentXml() {
    QString docPath = m_tempDir->path() + "/word/document.xml";

    QFile file(docPath);
    if (!file.open(QIODevice::ReadOnly)) {
        m_errorString = "document.xml açılamadı";
        return false;
    }

    // Binary olarak oku ve sakla
    m_xmlRawData = file.readAll();
    file.close();

    // DOM yükle (tablo işlemleri için)
    QString errorMsg;
    int errorLine, errorCol;
    if (!m_document.setContent(m_xmlRawData, true, &errorMsg, &errorLine, &errorCol)) {
        m_errorString = QString("XML parse hatası: %1 (satır %2, sütun %3)")
                            .arg(errorMsg).arg(errorLine).arg(errorCol);
        return false;
    }

    return true;
}

bool DocxWriter::saveDocumentXml() {
    QString docPath = m_tempDir->path() + "/word/document.xml";

    QFile file(docPath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        m_errorString = "document.xml yazılamadı";
        return false;
    }

    // DOM'u kaydet (m_xmlRawData yerine m_document kullanıyoruz)
    // Indent = 0 çünkü Word gereksiz boşlukları sevmez
    QTextStream stream(&file);
    m_document.save(stream, 0);
    file.close();

    if (file.error() != QFile::NoError) {
        m_errorString = QString("Dosya kaydetme hatası: %1").arg(file.errorString());
        return false;
    }

    return true;
}

void DocxWriter::replacePlaceholder(const QString& placeholder, const QString& value) {
    if (placeholder.isEmpty()) return;

    // Tüm w:t (text) elementlerini bul
    QDomNodeList texts = m_document.elementsByTagNameNS(NS_W, "t");

    for (int i = 0; i < texts.count(); ++i) {
        QDomElement textElem = texts.at(i).toElement();
        if (textElem.isNull()) continue;

        QString original = textElem.text();

        // Placeholder var mı?
        if (original.contains(placeholder)) {
            QString replaced = original;
            replaced.replace(placeholder, value);

            // Mevcut text node'u temizle
            while (textElem.hasChildNodes()) {
                textElem.removeChild(textElem.firstChild());
            }

            // Yeni text node ekle
            textElem.appendChild(m_document.createTextNode(replaced));

            qDebug() << "Placeholder replaced:" << placeholder << "->" << value;
        }
    }
}

QList<QDomElement> DocxWriter::findTables() {
    QList<QDomElement> tables;
    QDomNodeList tableNodes = m_document.elementsByTagNameNS(NS_W, "tbl");

    for (int i = 0; i < tableNodes.count(); ++i) {
        tables.append(tableNodes.at(i).toElement());
    }

    return tables;
}

void DocxWriter::setTableCell(int tableIndex, int rowIndex, int colIndex,
                              const QString& value, int fontSize) {
    auto tables = findTables();
    if (tableIndex >= tables.size()) return;

    QDomElement table = tables[tableIndex];

    // Doğrudan child row'ları al (nested tabloları dahil etmemek için)
    QList<QDomElement> rows;
    QDomNode child = table.firstChild();
    while (!child.isNull()) {
        if (child.isElement() && child.toElement().localName() == "tr") {
            rows.append(child.toElement());
        }
        child = child.nextSibling();
    }

    if (rowIndex >= rows.count()) return;

    QDomElement row = rows.at(rowIndex);

    // Doğrudan child cell'leri al
    QList<QDomElement> cells;
    child = row.firstChild();
    while (!child.isNull()) {
        if (child.isElement() && child.toElement().localName() == "tc") {
            cells.append(child.toElement());
        }
        child = child.nextSibling();
    }

    if (colIndex >= cells.count()) return;

    QDomElement cell = cells.at(colIndex);

    // Hücredeki paragrafı bul veya oluştur (doğrudan child)
    QDomElement para;
    QDomNode paraChild = cell.firstChild();
    while (!paraChild.isNull()) {
        if (paraChild.isElement() && paraChild.toElement().localName() == "p") {
            para = paraChild.toElement();
            break;
        }
        paraChild = paraChild.nextSibling();
    }

    if (para.isNull()) {
        para = m_document.createElementNS(NS_W, "w:p");
        cell.appendChild(para);
    } else {
        // Mevcut run'ları temizle (doğrudan children)
        QDomNode runChild = para.firstChild();
        while (!runChild.isNull()) {
            QDomNode next = runChild.nextSibling();
            if (runChild.isElement() && runChild.toElement().localName() == "r") {
                para.removeChild(runChild);
            }
            runChild = next;
        }
    }

    // === PYTHON İLE UYUMLU: Paragraf özellikleri (spacing, alignment) ===
    // Python: para.paragraph_format.space_after = Pt(0), space_before = Pt(0), line_spacing = 1.0
    // Mevcut pPr varsa kaldır, yenisini ekle
    QDomNode pPrChild = para.firstChild();
    while (!pPrChild.isNull()) {
        QDomNode next = pPrChild.nextSibling();
        if (pPrChild.isElement() && pPrChild.toElement().localName() == "pPr") {
            para.removeChild(pPrChild);
        }
        pPrChild = next;
    }

    // Yeni paragraf özellikleri oluştur
    QDomElement pPr = m_document.createElementNS(NS_W, "w:pPr");

    // Spacing: before=0, after=0, line=240 (240 twips = tek satır aralığı)
    QDomElement spacing = m_document.createElementNS(NS_W, "w:spacing");
    spacing.setAttribute("w:before", "0");
    spacing.setAttribute("w:after", "0");
    spacing.setAttribute("w:line", "240");
    spacing.setAttribute("w:lineRule", "auto");
    pPr.appendChild(spacing);

    // Paragrafı en başa ekle (pPr her zaman ilk element olmalı)
    para.insertBefore(pPr, para.firstChild());

    // Yeni run oluştur
    QDomElement run = m_document.createElementNS(NS_W, "w:r");

    // Run properties (font boyutu)
    QDomElement rPr = m_document.createElementNS(NS_W, "w:rPr");
    QDomElement sz = m_document.createElementNS(NS_W, "w:sz");
    sz.setAttribute("w:val", QString::number(fontSize * 2)); // Half-points
    rPr.appendChild(sz);

    QDomElement szCs = m_document.createElementNS(NS_W, "w:szCs");
    szCs.setAttribute("w:val", QString::number(fontSize * 2));
    rPr.appendChild(szCs);

    run.appendChild(rPr);

    // Text ekle
    QDomElement text = m_document.createElementNS(NS_W, "w:t");
    text.appendChild(m_document.createTextNode(value));
    run.appendChild(text);

    para.appendChild(run);
}

int DocxWriter::copyTableRow(int tableIndex, int sourceRowIndex) {
    auto tables = findTables();
    if (tableIndex >= tables.size()) return -1;

    QDomElement table = tables[tableIndex];

    // Doğrudan child row'ları al
    QList<QDomElement> rows;
    QDomNode child = table.firstChild();
    while (!child.isNull()) {
        if (child.isElement() && child.toElement().localName() == "tr") {
            rows.append(child.toElement());
        }
        child = child.nextSibling();
    }

    if (sourceRowIndex >= rows.count()) return -1;

    QDomElement sourceRow = rows.at(sourceRowIndex);
    QDomElement newRow = sourceRow.cloneNode(true).toElement();

    // === PYTHON İLE UYUMLU: Kopyalanan satırın içeriğini temizle ===
    // Python'daki _copy_row fonksiyonu: "Tüm hücrelerin içeriğini temizle ama formatı koru"
    QDomNode cellChild = newRow.firstChild();
    while (!cellChild.isNull()) {
        if (cellChild.isElement() && cellChild.toElement().localName() == "tc") {
            QDomElement cell = cellChild.toElement();

            // Hücredeki paragrafları bul ve içeriği temizle
            QDomNode paraChild = cell.firstChild();
            while (!paraChild.isNull()) {
                if (paraChild.isElement() && paraChild.toElement().localName() == "p") {
                    QDomElement para = paraChild.toElement();

                    // Paragraftaki run'ları bul ve text içeriğini temizle
                    QDomNode runChild = para.firstChild();
                    while (!runChild.isNull()) {
                        if (runChild.isElement() && runChild.toElement().localName() == "r") {
                            QDomElement run = runChild.toElement();

                            // Run içindeki text node'ları temizle
                            QDomNode textChild = run.firstChild();
                            while (!textChild.isNull()) {
                                if (textChild.isElement() && textChild.toElement().localName() == "t") {
                                    QDomElement textElem = textChild.toElement();
                                    // Text içeriğini temizle
                                    while (textElem.hasChildNodes()) {
                                        textElem.removeChild(textElem.firstChild());
                                    }
                                }
                                textChild = textChild.nextSibling();
                            }
                        }
                        runChild = runChild.nextSibling();
                    }
                }
                paraChild = paraChild.nextSibling();
            }
        }
        cellChild = cellChild.nextSibling();
    }

    // Satırı kaynak satırın arkasına ekle
    QDomNode nextSibling = sourceRow.nextSibling();
    if (nextSibling.isNull()) {
        table.appendChild(newRow);
    } else {
        table.insertBefore(newRow, nextSibling);
    }

    return sourceRowIndex + 1;
}

void DocxWriter::fillFirmaBilgileri(const FirmaBilgileri& firma) {
    // Şablondaki placeholder'lar {{}} olmadan yazılmış
    replacePlaceholder("firma_adi", firma.firmaAdi);
    replacePlaceholder("kontrol_adresi", firma.kontrolAdresi);
    replacePlaceholder("sgk_sicil", firma.sgkSicil);
    replacePlaceholder("rapor_numarasi", firma.raporNumarasi);
    replacePlaceholder("rapor_tarihi", firma.raporTarihi.toString("dd.MM.yyyy"));
    replacePlaceholder("sozlesme_id", firma.sozlesmeId);
    replacePlaceholder("baslangic_tarih_saat", firma.baslangicTarihSaat.toString("dd.MM.yyyy HH:mm"));
    replacePlaceholder("bitis_tarih_saat", firma.bitisTarihSaat.toString("dd.MM.yyyy HH:mm"));
    replacePlaceholder("bir_sonraki_kontrol", firma.birSonrakiKontrol.toString("dd.MM.yyyy"));
    replacePlaceholder("isim_soyisim", firma.kontrolEdenAdSoyad);
    replacePlaceholder("belge_no", firma.pkNo);

    // Teklif numarası (9.NOTLAR bölümünde)
    replacePlaceholder("tklf", firma.teklifNumarasi);
}

void DocxWriter::fillAnaDagitimPano(const AnaDagitimPano& pano, const QString& panoAdi) {
    // Şablondaki placeholder: PANO_adi1
    replacePlaceholder("PANO_adi1", panoAdi);

    // Şebeke Tipi Checkbox Mantığı
    // Şablondaki placeholder'lar: t_t, i_t, t_n, t_n_c-s, t_n-c, t_n-s
    // Python'da [X] ve [ ] formatı kullanılıyor
    QString sTip = pano.sebekeTipi.toUpper().replace(" ", "").replace("-", "");

    // TT sistemi
    replacePlaceholder("t_t", sTip == "TT" ? "[X]" : "[ ]");

    // IT sistemi
    replacePlaceholder("i_t", sTip == "IT" ? "[X]" : "[ ]");

    // TN sistemi (genel) - TN, TNC, TNS, TNCS hepsi TN kategorisinde
    bool isTN = sTip.startsWith("TN");
    replacePlaceholder("t_n", isTN ? "[X]" : "[ ]");

    // TN-C-S sistemi
    bool isTNCS = (sTip == "TNCS" || sTip.contains("TNCS"));
    replacePlaceholder("t_n_c-s", isTNCS ? "[X]" : "[ ]");

    // TN-C sistemi (TNCS değilse)
    bool isTNC = (sTip == "TNC" || (sTip.contains("TNC") && !isTNCS));
    replacePlaceholder("t_n-c", isTNC ? "[X]" : "[ ]");
    replacePlaceholder("t_n_c", isTNC ? "[X]" : "[ ]");  // Eski placeholder

    // TN-S sistemi
    bool isTNS = (sTip == "TNS" || sTip.contains("TNS"));
    replacePlaceholder("t_n-s", isTNS ? "[X]" : "[ ]");

    // SPD (Parafudr) checkbox'ları
    bool hasSPD = !pano.parafudrTip.isEmpty() && pano.parafudrTip != "-";
    replacePlaceholder("spd_evet", hasSPD ? "[X]" : "[ ]");
    replacePlaceholder("spd_hayir", hasSPD ? "[ ]" : "[X]");

    // Parafudr bilgileri
    replacePlaceholder("PARAFUDR_TIP", pano.parafudrTip);
    replacePlaceholder("PARAFUDR_Imax", pano.parafudrImax);

    // Ana Dağıtım Pano bilgileri
    replacePlaceholder("ADP_tip", pano.sigortaTipiAna);
    replacePlaceholder("ADP_anma", QString::number(pano.nominalAkimAna));

    // RCD bilgileri
    replacePlaceholder("RCD_tipi", pano.rcdBilgisi);
    replacePlaceholder("RCD_dayanim", pano.rcdAnmaAkimi);

    // Topraklama kesitleri
    replacePlaceholder("Sistem_top", pano.sistemTopraklamaKesiti);
    replacePlaceholder("Ana_top", pano.anaEspotansiyelKesiti);

    // Empedans ve akım
    replacePlaceholder("Z_E", pano.distCevrimEmpedansi);

    // I_f = 230 / Z_E (HER ZAMAN otomatik hesaplama - Python ile uyumlu)
    // Python'da I_f Excel'den okunmaz, her zaman Z_E'den hesaplanır
    // DÜZELTME: I_f her zaman A cinsinden yazılır (kA değil)
    QString hataAkimi;
    if (!pano.distCevrimEmpedansi.isEmpty()) {
        QString zEStr = pano.distCevrimEmpedansi;
        zEStr.replace(',', '.');  // Türkçe ondalık ayracı
        bool ok;
        double zE = zEStr.toDouble(&ok);
        if (ok && zE > 0) {
            double iF = 230.0 / zE;
            // Python'daki gibi - A cinsinden yaz (kA KULLANMA)
            hataAkimi = QString::number(iF, 'f', 1);  // Örn: 328.6 (A)
        }
    }
    replacePlaceholder("I_f", hataAkimi);

    // Potansiyel dengeleme için en büyük topraklama kesiti
    replacePlaceholder("en_buyuk_top_kesit", pano.enBuyukTopKesit);

    // Zemin izolasyonu uygunluk
    replacePlaceholder("zo_uygunluk", pano.zeminIzoUygunluk);

    // Temel topraklama direnci - Tablo 1'e (0. tablo), satır 14, sütun 9
    // Python: 'Temel Topraklama Direnci (Ohm)': (14, 9),
    if (!pano.topraklamaDirenci.isEmpty()) {
        setTableCell(0, 14, 9, pano.topraklamaDirenci, 9);
    }
}

void DocxWriter::fillCihazBilgileri(const FirmaBilgileri& firma) {
    // Termal kamera bilgileri
    // Önce yeni alanları dene, sonra eski alanları (geriye uyumluluk)
    QString termalAdi = !firma.termalCihazAdi.isEmpty() ? firma.termalCihazAdi : firma.cihaz1Adi;
    QString termalSeriNo = !firma.termalSeriNo.isEmpty() ? firma.termalSeriNo : firma.cihaz1SeriNo;
    QString termalKalTarihi = !firma.termalKalibrasyonTarihi.isEmpty() ? firma.termalKalibrasyonTarihi : firma.cihaz1KalibrasyonTarihi;
    QString termalKalGecerlilik = !firma.termalKalibrasyonGecerlilik.isEmpty() ? firma.termalKalibrasyonGecerlilik : firma.cihaz1KalibrasyonGecerlilik;
    QString termalKalNo = !firma.termalKalibrasyonNo.isEmpty() ? firma.termalKalibrasyonNo : firma.cihaz1KalibrasyonNo;

    replacePlaceholder("termal_cihaz_adi", termalAdi);
    replacePlaceholder("termal_seri_numarasi", termalSeriNo);
    replacePlaceholder("termal_kalibrasyon_tarihi", termalKalTarihi);
    replacePlaceholder("termal_kalibrasyon_gecerlilik", termalKalGecerlilik);
    replacePlaceholder("termal_kalibrasyon_no", termalKalNo);

    // Ölçüm cihazı bilgileri
    QString olcumAdi = !firma.olcumCihazAdi.isEmpty() ? firma.olcumCihazAdi : firma.cihaz2Adi;
    QString olcumSeriNo = !firma.olcumSeriNo.isEmpty() ? firma.olcumSeriNo : firma.cihaz2SeriNo;
    QString olcumKalTarihi = !firma.olcumKalibrasyonTarihi.isEmpty() ? firma.olcumKalibrasyonTarihi : firma.cihaz2KalibrasyonTarihi;
    QString olcumKalGecerlilik = !firma.olcumKalibrasyonGecerlilik.isEmpty() ? firma.olcumKalibrasyonGecerlilik : firma.cihaz2KalibrasyonGecerlilik;
    QString olcumKalNo = !firma.olcumKalibrasyonNo.isEmpty() ? firma.olcumKalibrasyonNo : firma.cihaz2KalibrasyonNo;

    replacePlaceholder("olcum_cihaz_adi", olcumAdi);
    replacePlaceholder("olcum_seri_numarasi", olcumSeriNo);
    replacePlaceholder("olcum_kalibrasyon_tarihi", olcumKalTarihi);
    replacePlaceholder("olcum_kalibrasyon_gecerlilik", olcumKalGecerlilik);
    replacePlaceholder("olcum_kalibrasyon_no", olcumKalNo);
}

void DocxWriter::fillGozleKontrol(const QVector<GozleKontrolMaddesi>& maddeler) {
    // Tablo 2 - Gözle kontrol
    // Şablondaki placeholder formatı: GK_01, GK_02, ... GK_29
    for (const auto& madde : maddeler) {
        // Madde numarasını 2 basamaklı formata çevir (1 -> 01, 10 -> 10)
        QString placeholder = QString("GK_%1").arg(madde.maddeNo, 2, 10, QChar('0'));

        // Sonucu yazdır (Uygun, Uygun Değil, Uygulanamaz)
        replacePlaceholder(placeholder, madde.sonuc);
    }
}

void DocxWriter::fillFlukeBilgileri(const QString& fotoTarihi, const QString& fotoNo) {
    // Fluke DOCX'ten gelen fotoğraf tarihi ve numarası
    // GK_24 = Fotoğraf Tarihi
    // GK_25 = Fotoğraf Numarası
    if (!fotoTarihi.isEmpty()) {
        replacePlaceholder("GK_24", fotoTarihi);
    }
    if (!fotoNo.isEmpty()) {
        replacePlaceholder("GK_25", fotoNo);
    }
}

void DocxWriter::fillFonksiyonTestleri(const QVector<FonksiyonTesti>& testler,
                                        const AnaDagitimPano& anaPano,
                                        const QString& panoAdi) {
    // === PANO_adi1 PLACEHOLDER (Tablo 2 başlık) ===
    replacePlaceholder("PANO_adi1", panoAdi);

    // Header placeholder'larını değiştir (Fonksiyon testleri bölümündeki ölçüm değerleri)
    // Z_ln'den 380/Z_ln otomatik hesapla
    QString kisaDevre;
    if (!anaPano.loopLN.isEmpty()) {
        bool ok;
        double zLn = anaPano.loopLN.toDouble(&ok);
        if (ok && zLn > 0) {
            int kd = qRound(380.0 / zLn);
            kisaDevre = QString::number(kd);
        }
    }

    // Önce 380/Z_ln, sonra diğerleri (Z_ln önce yazılırsa 380/Z_ln bozuluyor)
    replacePlaceholder("380/Z_ln", kisaDevre);
    replacePlaceholder("Z_x", anaPano.loopPeN);
    replacePlaceholder("Z_ln", anaPano.loopLN);
    replacePlaceholder("F_F", QString::number(anaPano.sistemGerilimi));

    // L-N gerilimi hesapla (F-F / sqrt(3))
    int gerilimLN = qRound(anaPano.sistemGerilimi / 1.732);
    replacePlaceholder("L_N", QString::number(gerilimLN));

    // N-PE genelde 0 veya çok düşük
    replacePlaceholder("N_PE", "0");

    // === 3 FAZ SİMETRİ ===
    // "3faz_simetri" placeholder'ı - varsayılan olarak "Simetrik" yazılır
    replacePlaceholder("3faz_simetri", QString::fromUtf8("Simetrik"));

    if (testler.isEmpty()) return;

    const int TABLE_INDEX = 2;  // Fonksiyon testleri tablosu (6.FONKSİYON KONTROL)
    const int FIRST_DATA_ROW = 10;  // Python'daki start_row = 10 (header'dan sonra veri satırları)

    auto tables = findTables();
    if (TABLE_INDEX >= tables.size()) return;

    QDomElement table = tables[TABLE_INDEX];

    // Doğrudan child row'ları sayarak toplam satır hesapla
    auto countDirectRows = [](QDomElement& tbl) -> int {
        int count = 0;
        QDomNode child = tbl.firstChild();
        while (!child.isNull()) {
            if (child.isElement() && child.toElement().localName() == "tr") {
                count++;
            }
            child = child.nextSibling();
        }
        return count;
    };

    int rowCount = countDirectRows(table);

    // Python'daki sütun eşlemesi:
    // 0: No., 1-2: Linye Adı (merged), 3: Açma eğrisi, 4: Kutup sayısı
    // 5-6: In(A) (merged), 7: Icu, 8-9: Faz kesiti (merged), 10: N/PEN Kesiti
    // 11-12: PE Kesiti (merged), 13-14: Ib (merged), 15: Iz, 16-17: IΔ mA (merged)
    // 18: TΔ ms, 19: Sonuç

    // Kablo Akım Taşıma Kapasiteleri (Iz hesabı için)
    // Python mantığı: parse_kesit -> base kesit bul, iz_table'dan değer al, factor ile çarp
    auto calcIz = [](const QString& fazKesiti) -> QString {
        static const QMap<double, int> izTable = {
            {0.75, 13}, {1.0, 16}, {1.5, 20}, {2.5, 27}, {4.0, 36},
            {6.0, 47}, {10.0, 65}, {16.0, 87}, {25.0, 115}, {35.0, 143},
            {50.0, 178}, {70.0, 220}, {95.0, 265}, {120.0, 310}, {150.0, 355},
            {185.0, 400}, {240.0, 480}, {300.0, 555}, {400.0, 770}, {500.0, 880}
        };

        QString val = fazKesiti.toLower().replace(',', '.').trimmed();
        if (val.isEmpty()) return "";

        int factor = 1;
        QString base = val;

        // "3x16" gibi çarpanlı değerler için
        if (val.contains('x')) {
            QStringList parts = val.split('x');
            if (parts.size() == 2) {
                factor = parts[0].toInt();
                if (factor <= 0) factor = 1;
                base = parts[1];
            }
        }

        double baseSize = base.toDouble();
        if (baseSize <= 0) return "";

        // Python mantığı: önce tablo'dan base değeri bul
        int baseIz = 0;

        // Tam eşleşme var mı?
        if (izTable.contains(baseSize)) {
            baseIz = izTable[baseSize];
        } else {
            // Daha büyük ilk değeri bul (Python: for k in sorted(iz_table.keys()): if size <= k)
            for (auto it = izTable.begin(); it != izTable.end(); ++it) {
                if (baseSize <= it.key()) {
                    baseIz = it.value();
                    break;
                }
            }
            // Hala bulunamadıysa son değeri kullan
            if (baseIz == 0) {
                baseIz = izTable.last();
            }
        }

        // Factor ile çarp (Python: base_iz * factor)
        return QString::number(baseIz * factor);
    };

    // Şablon satırını sakla (kopyalama için)
    int templateRow = FIRST_DATA_ROW;

    // Python: multi_pano_gui.py:4051-4064 - Ana pano RCD değerleri (fallback için)
    QString anaRcdMa = anaPano.rcdAnmaAkimi;  // Ana RCD Test Akimi
    QString anaRcdMs;  // Ana RCD Acma Suresi - anaPano'da mevcut değilse boş

    for (int i = 0; i < testler.size(); ++i) {
        const auto& test = testler[i];
        int rowIdx = FIRST_DATA_ROW + i;

        // Yeni satır gerekiyorsa kopyala
        if (rowIdx >= rowCount) {
            copyTableRow(TABLE_INDEX, templateRow);
            rowCount = countDirectRows(table);  // Yenile
        }

        // Iz hesapla (akimKapasitesi boşsa faz kesitinden)
        QString iz = test.akimKapasitesi > 0 ? QString::number(test.akimKapasitesi) : calcIz(test.fazKesiti);

        // === UYGUNLUK DEĞERLENDİRMESİ (Python: multi_pano_gui.py:4085-4140) ===
        QString sonuc = test.sonuc;
        QString rcdMa = test.rcdMa;
        QString rcdMs = test.rcdMs;

        // KAKR grubu kontrolü (Python: is_kakr_group = "KAKR" in linye_adi.upper())
        bool isKakrGroup = test.linye.toUpper().contains("KAKR");

        // 1) Iz < In kontrolü (KAKR grupları için geçerli değil)
        if (!isKakrGroup && test.akimKapasitesi > 0 && test.nominalAkim > 0 &&
            test.akimKapasitesi < test.nominalAkim) {
            sonuc = QString::fromUtf8("Uygun Değil");
        }

        // 2) In ≤ 32A ve 30mA KAKR yok kontrolü (ANA SİGORTA HARİÇ)
        // Python: if not is_ana_sigorta and in_val <= 32...
        if (!test.isAnaSigorta && test.nominalAkim > 0 && test.nominalAkim <= 32) {
            bool has30maKakr = test.kakrVar && test.rcd == "30mA";
            if (!has30maKakr && !test.kakrYok) {
                sonuc = QString::fromUtf8("Uygun Değil");
            }
        }

        // 3) KAKR Yok işaretli → mA/mS = 'x', Sonuç = Uygun Değil
        // Python: multi_pano_gui.py:4117-4121
        if (test.kakrYok) {
            rcdMa = "x";
            rcdMs = "x";
            sonuc = QString::fromUtf8("Uygun Değil");
        }
        // 4) RCD boş ve KAKR checkbox'ları işaretlenmemiş → Ana pano değerlerini kullan
        // Python: multi_pano_gui.py:4122-4125
        else if (rcdMa.isEmpty() && rcdMs.isEmpty() && !test.kakrVar) {
            rcdMa = anaRcdMa;
            rcdMs = anaRcdMs;
        }

        // Sonuç boşsa varsayılan "Uygun"
        if (sonuc.isEmpty()) sonuc = "Uygun";

        // === DOĞRU SÜTUN EŞLEMESİ (XML'deki gerçek hücre indeksleri) ===
        // Satır 9 header'da 14 hücre var:
        // [0]=No, [1]=Linye Adı, [2]=Açma eğrisi, [3]=Kutup, [4]=In(A)
        // [5]=Icu, [6]=Faz kesiti, [7]=N/PEN, [8]=PE kesiti
        // [9]=Ib, [10]=Iz, [11]=IΔ(mA), [12]=TΔ(ms), [13]=Sonuç
        setTableCell(TABLE_INDEX, rowIdx, 0, QString::number(i + 1), 6);   // No
        setTableCell(TABLE_INDEX, rowIdx, 1, test.linye, 6);               // Linye Adı
        setTableCell(TABLE_INDEX, rowIdx, 2, test.sigortaTipi, 6);         // Açma eğrisi
        setTableCell(TABLE_INDEX, rowIdx, 3, QString::number(test.kutupSayisi), 6);  // Kutup sayısı
        setTableCell(TABLE_INDEX, rowIdx, 4, QString::number(test.nominalAkim), 6);  // In(A)
        setTableCell(TABLE_INDEX, rowIdx, 5, test.icu.isEmpty() ? "6" : test.icu, 6); // Icu
        setTableCell(TABLE_INDEX, rowIdx, 6, test.fazKesiti, 6);           // Faz kesiti
        setTableCell(TABLE_INDEX, rowIdx, 7, test.notrKesiti, 6);          // N/PEN Kesiti
        setTableCell(TABLE_INDEX, rowIdx, 8, test.toprakKesiti, 6);        // PE Kesiti
        setTableCell(TABLE_INDEX, rowIdx, 9, test.ib, 6);                  // Ib
        setTableCell(TABLE_INDEX, rowIdx, 10, iz, 6);                      // Iz
        setTableCell(TABLE_INDEX, rowIdx, 11, test.rcd.isEmpty() ? rcdMa : test.rcd, 6);  // IΔ mA (RCD değeri veya hesaplanan)
        setTableCell(TABLE_INDEX, rowIdx, 12, rcdMs.isEmpty() ? test.rcdMs : rcdMs, 6);   // TΔ ms
        setTableCell(TABLE_INDEX, rowIdx, 13, sonuc, 6);                   // Sonuç
    }
}

void DocxWriter::fillPotansiyelDengelemeVeZemin(const PanoData& pano,
                                                  const QVector<FonksiyonTesti>& fonksiyonTestleri,
                                                  const QString& zeminIzolasyonDurum) {
    // === 6.2 POTANSİYEL DENGELEME (Tablo 3, index 3) ===
    // Pano adını da doldur (PANO_adi1)
    replacePlaceholder("PANO_adi1", pano.panoAdi);

    // En büyük topraklama kesitini fonksiyon testlerinden hesapla (Python mantığı)
    QString enBuyukTopKesit = pano.enBuyukTopKesit;
    if (enBuyukTopKesit.isEmpty() && !fonksiyonTestleri.isEmpty()) {
        double maxKesit = 0;
        QString maxKesitStr;
        for (const auto& test : fonksiyonTestleri) {
            QString kesitStr = test.toprakKesiti;
            if (kesitStr.isEmpty()) continue;

            QString val = kesitStr.toLower().replace(',', '.');
            double kesitVal = 0;
            if (val.contains('x')) {
                QStringList parts = val.split('x');
                if (parts.size() >= 2) {
                    kesitVal = parts.last().toDouble();
                }
            } else {
                kesitVal = val.toDouble();
            }

            if (kesitVal > maxKesit) {
                maxKesit = kesitVal;
                maxKesitStr = kesitStr;
            }
        }
        enBuyukTopKesit = maxKesitStr.isEmpty() ? "6" : maxKesitStr;
    }

    replacePlaceholder("en_buyuk_top_kesit", enBuyukTopKesit.isEmpty() ? "6" : enBuyukTopKesit);

    // === 6.3 ZEMİN İZOLASYONU ===
    // Zemin izolasyonu durumuna göre değerleri belirle (Python'daki mantık)
    QString enVal, boyVal, izoDirenci, izoUygunluk;
    QString durum = zeminIzolasyonDurum.isEmpty() ? pano.izoUygunluk : zeminIzolasyonDurum;

    if (durum == "Uygun" || durum == "UYGUN" || durum.isEmpty()) {
        enVal = pano.zeminEn.isEmpty() ? "1" : pano.zeminEn;
        boyVal = pano.zeminBoy.isEmpty() ? "1" : pano.zeminBoy;
        izoDirenci = QString::fromUtf8(">50MΩ");
        izoUygunluk = "UYGUN";
    } else if (durum.contains("Uygulanamaz", Qt::CaseInsensitive)) {
        enVal = "-";
        boyVal = "-";
        izoDirenci = "-";
        izoUygunluk = "UYGULANAMAZ";
    } else {
        enVal = "x";
        boyVal = "x";
        izoDirenci = "-";
        izoUygunluk = QString::fromUtf8("UYGUN DEĞİL");
    }

    qDebug() << "[DEBUG] 6.3 Zemin İzolasyonu:" << "en=" << enVal << "boy=" << boyVal
             << "izo_direnci=" << izoDirenci << "izo_uygunluk=" << izoUygunluk;

    // Placeholder değiştirme
    replacePlaceholder("en", enVal);
    replacePlaceholder("boy", boyVal);

    // İzo_direnci placeholder'ı için tüm olası varyasyonları dene
    replacePlaceholder(QString::fromUtf8("İzo_direnci"), izoDirenci);
    replacePlaceholder("Izo_direnci", izoDirenci);
    replacePlaceholder(QString::fromUtf8("ızo_direnci"), izoDirenci);
    replacePlaceholder("izo_direnci", izoDirenci);
    replacePlaceholder("İzo_dir1ci", izoDirenci);  // Görüntüdeki placeholder
    replacePlaceholder("Izo_dir1ci", izoDirenci);
    replacePlaceholder("izo_dir1ci", izoDirenci);

    replacePlaceholder("izo_uygunluk", izoUygunluk);
    replacePlaceholder("zo_uygunluk", izoUygunluk);

    // === DOĞRUDAN HÜCRE İNDEKSLERİ İLE DEĞİŞTİRME (Python'daki gibi) ===
    // Tablo 3 (index 3) - 6.3 Zemin İzolasyonu satırı
    const int TABLE_INDEX = 3;
    auto tables = findTables();
    if (TABLE_INDEX < tables.size()) {
        // 6.3 Zemin İzolasyonu satırında hücreleri direkt değiştir
        // Python: R6C4=izo_direnci, R6C5=izo_uygunluk
        setTableCell(TABLE_INDEX, 6, 4, izoDirenci, 9);
        setTableCell(TABLE_INDEX, 6, 5, izoUygunluk, 9);
    }
}

QString DocxWriter::addImage(const QString& imagePath) {
    if (!QFile::exists(imagePath)) {
        return {};
    }

    ++m_imageCounter;
    QString rId = QString("rId%1").arg(100 + m_imageCounter);  // Offset to avoid conflicts

    // Resmi media klasörüne kopyala
    QString mediaDir = m_tempDir->path() + "/word/media";
    QDir().mkpath(mediaDir);

    QFileInfo info(imagePath);
    QString destName = QString("image%1.%2").arg(m_imageCounter).arg(info.suffix());
    QString destPath = mediaDir + "/" + destName;

    QFile::copy(imagePath, destPath);
    m_addedImages.append(destPath);

    // document.xml.rels dosyasına relationship ekle
    QString relsPath = m_tempDir->path() + "/word/_rels/document.xml.rels";
    QFile relsFile(relsPath);

    if (relsFile.open(QIODevice::ReadWrite)) {
        QDomDocument relsDoc;
        relsDoc.setContent(&relsFile);

        QDomElement root = relsDoc.documentElement();
        QDomElement rel = relsDoc.createElement("Relationship");
        rel.setAttribute("Id", rId);
        rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        rel.setAttribute("Target", "media/" + destName);
        root.appendChild(rel);

        relsFile.seek(0);
        QTextStream stream(&relsFile);
        relsDoc.save(stream, 2);
        relsFile.close();
    }

    return rId;
}

void DocxWriter::addThermalImages(const QVector<TermalGoruntu>& images) {
    // Boş vector ise hiçbir şey yapma
    if (images.isEmpty()) {
        return;
    }

    // Tablo 5 (index 5) - 8.EKİPMAN FOTOĞRAFLARI
    const int TABLE_INDEX = 5;

    auto tables = findTables();
    if (TABLE_INDEX >= tables.size()) {
        qWarning() << "Tablo 5 (Ekipman Fotoğrafları) bulunamadı";
        return;
    }

    QDomElement table = tables[TABLE_INDEX];

    // Direct child row iteration - elementsByTagNameNS recursive olduğu için kullanmıyoruz
    QVector<QDomElement> directRows;
    QDomNode child = table.firstChild();
    while (!child.isNull()) {
        if (child.isElement() && child.toElement().localName() == "tr") {
            directRows.append(child.toElement());
        }
        child = child.nextSibling();
    }

    // Satır 1'deki ilk hücreye resimleri ekle (satır 0 başlık)
    if (directRows.count() < 2) {
        qWarning() << "Ekipman Fotoğrafları tablosunda yeterli satır yok";
        return;
    }

    QDomElement targetRow = directRows.at(1);

    // Direct child cell iteration
    QDomElement cell;
    QDomNode cellChild = targetRow.firstChild();
    while (!cellChild.isNull()) {
        if (cellChild.isElement() && cellChild.toElement().localName() == "tc") {
            cell = cellChild.toElement();
            break;
        }
        cellChild = cellChild.nextSibling();
    }

    if (cell.isNull()) {
        qWarning() << "Hedef satırda hücre bulunamadı";
        return;
    }

    // Hücredeki paragrafı bul veya oluştur - direct child iteration
    QDomElement para;
    QDomNode paraChild = cell.firstChild();
    while (!paraChild.isNull()) {
        if (paraChild.isElement() && paraChild.toElement().localName() == "p") {
            para = paraChild.toElement();
            break;
        }
        paraChild = paraChild.nextSibling();
    }

    if (para.isNull()) {
        para = m_document.createElementNS(NS_W, "w:p");
        cell.appendChild(para);
    } else {
        // Mevcut run'ları temizle (boş içerik varsa) - direct child iteration
        QVector<QDomNode> runsToRemove;
        QDomNode runChild = para.firstChild();
        while (!runChild.isNull()) {
            if (runChild.isElement() && runChild.toElement().localName() == "r") {
                runsToRemove.append(runChild);
            }
            runChild = runChild.nextSibling();
        }
        for (const auto& run : runsToRemove) {
            para.removeChild(run);
        }
    }

    // Paragraf ortala - direct child iteration for pPr
    QDomElement pPr;
    QDomNode pPrChild = para.firstChild();
    while (!pPrChild.isNull()) {
        if (pPrChild.isElement() && pPrChild.toElement().localName() == "pPr") {
            pPr = pPrChild.toElement();
            break;
        }
        pPrChild = pPrChild.nextSibling();
    }

    if (pPr.isNull()) {
        pPr = m_document.createElementNS(NS_W, "w:pPr");
        para.insertBefore(pPr, para.firstChild());
    }
    QDomElement jc = m_document.createElementNS(NS_W, "w:jc");
    jc.setAttribute("w:val", "center");
    pPr.appendChild(jc);

    // Görüntü sayısına göre boyut belirle (Python ile uyumlu)
    int imgCount = 0;
    for (const auto& img : images) {
        if (img.isValid()) imgCount++;
    }

    // EMU birimleri (1 cm = 360000 EMU)
    qint64 imgWidth, imgHeight;
    if (imgCount >= 4) {
        imgWidth = 4 * 360000;   // 4 cm
        imgHeight = 5 * 360000;  // 5 cm
    } else {
        imgWidth = 5 * 360000;   // 5 cm
        imgHeight = 6 * 360000;  // 6 cm
    }

    int addedCount = 0;
    for (const auto& img : images) {
        if (!img.isValid()) continue;

        QString rId = addImage(img.imagePath);
        if (rId.isEmpty()) continue;

        // DrawingML XML yapısı
        QDomElement run = m_document.createElementNS(NS_W, "w:r");
        QDomElement drawing = m_document.createElementNS(NS_W, "w:drawing");

        // wp:inline
        QDomElement inlineEl = m_document.createElementNS(NS_WP, "wp:inline");
        inlineEl.setAttribute("distT", "0");
        inlineEl.setAttribute("distB", "0");
        inlineEl.setAttribute("distL", "0");
        inlineEl.setAttribute("distR", "0");

        // wp:extent
        QDomElement extent = m_document.createElementNS(NS_WP, "wp:extent");
        extent.setAttribute("cx", QString::number(imgWidth));
        extent.setAttribute("cy", QString::number(imgHeight));
        inlineEl.appendChild(extent);

        // wp:docPr
        QDomElement docPr = m_document.createElementNS(NS_WP, "wp:docPr");
        docPr.setAttribute("id", QString::number(100 + addedCount));
        docPr.setAttribute("name", QString("Termal Görüntü %1").arg(addedCount + 1));
        inlineEl.appendChild(docPr);

        // a:graphic
        QDomElement graphic = m_document.createElementNS(NS_A, "a:graphic");

        // a:graphicData
        QDomElement graphicData = m_document.createElementNS(NS_A, "a:graphicData");
        graphicData.setAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");

        // pic:pic
        QDomElement pic = m_document.createElementNS(NS_PIC, "pic:pic");

        // pic:nvPicPr
        QDomElement nvPicPr = m_document.createElementNS(NS_PIC, "pic:nvPicPr");
        QDomElement cNvPr = m_document.createElementNS(NS_PIC, "pic:cNvPr");
        cNvPr.setAttribute("id", QString::number(100 + addedCount));
        cNvPr.setAttribute("name", QString("Termal %1").arg(addedCount + 1));
        nvPicPr.appendChild(cNvPr);
        QDomElement cNvPicPr = m_document.createElementNS(NS_PIC, "pic:cNvPicPr");
        nvPicPr.appendChild(cNvPicPr);
        pic.appendChild(nvPicPr);

        // pic:blipFill
        QDomElement blipFill = m_document.createElementNS(NS_PIC, "pic:blipFill");
        QDomElement blip = m_document.createElementNS(NS_A, "a:blip");
        blip.setAttributeNS(NS_R, "r:embed", rId);
        blipFill.appendChild(blip);
        QDomElement stretch = m_document.createElementNS(NS_A, "a:stretch");
        QDomElement fillRect = m_document.createElementNS(NS_A, "a:fillRect");
        stretch.appendChild(fillRect);
        blipFill.appendChild(stretch);
        pic.appendChild(blipFill);

        // pic:spPr
        QDomElement spPr = m_document.createElementNS(NS_PIC, "pic:spPr");
        QDomElement xfrm = m_document.createElementNS(NS_A, "a:xfrm");
        QDomElement off = m_document.createElementNS(NS_A, "a:off");
        off.setAttribute("x", "0");
        off.setAttribute("y", "0");
        xfrm.appendChild(off);
        QDomElement ext = m_document.createElementNS(NS_A, "a:ext");
        ext.setAttribute("cx", QString::number(imgWidth));
        ext.setAttribute("cy", QString::number(imgHeight));
        xfrm.appendChild(ext);
        spPr.appendChild(xfrm);
        QDomElement prstGeom = m_document.createElementNS(NS_A, "a:prstGeom");
        prstGeom.setAttribute("prst", "rect");
        spPr.appendChild(prstGeom);
        pic.appendChild(spPr);

        graphicData.appendChild(pic);
        graphic.appendChild(graphicData);
        inlineEl.appendChild(graphic);
        drawing.appendChild(inlineEl);
        run.appendChild(drawing);
        para.appendChild(run);

        // Görüntüler arası boşluk
        if (addedCount < imgCount - 1) {
            QDomElement spaceRun = m_document.createElementNS(NS_W, "w:r");
            QDomElement spaceText = m_document.createElementNS(NS_W, "w:t");
            spaceText.setAttribute("xml:space", "preserve");
            spaceText.appendChild(m_document.createTextNode("  "));
            spaceRun.appendChild(spaceText);
            para.appendChild(spaceRun);
        }

        addedCount++;
    }

    qDebug() << "Termal görüntü eklendi:" << addedCount << "adet";
}

QString DocxWriter::generateKusurAciklamasi(const QVector<Kusur>& kusurlar) {
    if (kusurlar.isEmpty()) {
        return "Kontrol edilen tesisatta herhangi bir kusur tespit edilmemiştir.";
    }

    QStringList aciklamalar;
    for (const auto& k : kusurlar) {
        aciklamalar.append(QString("- %1: %2 (%3)").arg(k.linye, k.kusurAciklamasi, k.kusurDerecesi));
    }

    return "Tespit edilen kusurlar:\n" + aciklamalar.join("\n");
}

void DocxWriter::fillSonuc(const QString& sonuc, const QString& aciklama) {
    replacePlaceholder("genel_sonuc", sonuc);
    replacePlaceholder("sonuc_aciklama", aciklama);

    // === KUSUR AÇIKLAMALARI (Tablo 4, index 4) ===
    // Python'daki gibi: Tablo 4, Satır 1, Hücre 0
    // Eğer kusur yoksa "Herhangi bir kusur tespit edilmemiştir." yaz
    const int KUSUR_TABLE_INDEX = 4;
    auto tables = findTables();
    if (KUSUR_TABLE_INDEX < tables.size()) {
        QString kusurMetni = aciklama.isEmpty()
            ? QString::fromUtf8("Herhangi bir kusur tespit edilmemiştir.")
            : aciklama;
        setTableCell(KUSUR_TABLE_INDEX, 1, 0, kusurMetni, 9);
    }
}

void DocxWriter::fillUygunluk(const QString& uygunlukDurumu) {
    // "Uygun" veya "Uygun Değil" durumuna göre metni belirle
    QString uygunlukMetni;
    if (uygunlukDurumu == "Uygun") {
        uygunlukMetni = QString::fromUtf8("kullanımı uygundur");
    } else {
        uygunlukMetni = QString::fromUtf8("kullanımı uygun değildir");
    }

    // Şablondaki "uygunluk" placeholder'ını değiştir
    replacePlaceholder("uygunluk", uygunlukMetni);

    // Eski checkbox-tabanlı placeholder'ları da destekle (geriye uyumluluk)
    if (uygunlukDurumu == "Uygun") {
        replacePlaceholder("uygun_checkbox", QString::fromUtf8("☑"));
        replacePlaceholder("uygun_degil_checkbox", QString::fromUtf8("☐"));
    } else {
        replacePlaceholder("uygun_checkbox", QString::fromUtf8("☐"));
        replacePlaceholder("uygun_degil_checkbox", QString::fromUtf8("☑"));
    }
}

void DocxWriter::addPageBreakBeforeTable(int tableIndex) {
    auto tables = findTables();
    if (tableIndex >= tables.size()) return;

    QDomElement table = tables[tableIndex];

    // Tablodan önce paragraf ekle ve sayfa sonu koy
    QDomElement pageBreakPara = m_document.createElementNS(NS_W, "w:p");
    QDomElement run = m_document.createElementNS(NS_W, "w:r");
    QDomElement br = m_document.createElementNS(NS_W, "w:br");
    br.setAttribute("w:type", "page");
    run.appendChild(br);
    pageBreakPara.appendChild(run);

    table.parentNode().insertBefore(pageBreakPara, table);
}

bool DocxWriter::createZip(const QString& outputPath) {
    QZipWriter zip(outputPath);
    if (!zip.isWritable()) {
        m_errorString = QString("Çıktı dosyası oluşturulamadı: %1").arg(outputPath);
        return false;
    }

    // Geçici dizindeki tüm dosyaları ZIP'e ekle
    QDirIterator it(m_tempDir->path(), QDirIterator::Subdirectories);
    QString basePath = m_tempDir->path() + "/";

    while (it.hasNext()) {
        it.next();
        QFileInfo info = it.fileInfo();

        if (info.isFile()) {
            QString relativePath = info.absoluteFilePath().mid(basePath.length());

            QFile file(info.absoluteFilePath());
            if (file.open(QIODevice::ReadOnly)) {
                zip.addFile(relativePath, file.readAll());
                file.close();
            }
        }
    }

    zip.close();
    return true;
}

bool DocxWriter::save(const QString& outputPath) {
    // document.xml'i kaydet
    if (!saveDocumentXml()) {
        return false;
    }

    // ZIP olarak paketle
    if (!createZip(outputPath)) {
        return false;
    }

    return true;
}

} // namespace RaporSistemi

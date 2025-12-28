/**
 * PdfParser.cpp
 *
 * PDF parse implementasyonu.
 */

#include "PdfParser.h"
#include <QPdfDocument>
#include <QPdfSelection>
#include <QFile>
#include <QDebug>
#include <QProcess>
#include <QCoreApplication>

namespace RaporSistemi {

PdfParser::PdfParser() {
    initPatterns();
}

PdfParser::~PdfParser() = default;

void PdfParser::initPatterns() {
    // Sözleşme ID: "Sözleşme Id: 12345678"
    m_patternSozlesmeId = QRegularExpression(
        R"(S[öo]zle[şs]me\s*[Ii]d\s*[:=]\s*(\d+))",
        QRegularExpression::CaseInsensitiveOption);

    // Firma adı: "İş Yeri Unvanı: XYZ LTD. ŞTİ."
    m_patternFirmaAdi = QRegularExpression(
        R"([İI][şs]\s*[Yy]eri\s*[Üü]nvan[ıi]\s*[:=]\s*(.+?)(?:\n|SGK))",
        QRegularExpression::CaseInsensitiveOption);

    // SGK Sicil: "SGK Sicil No: 1234567890123"
    m_patternSgk = QRegularExpression(
        R"(SGK\s*Sicil\s*(?:No|Numaras[ıi])\s*[:=]\s*([\d\-\.]+))",
        QRegularExpression::CaseInsensitiveOption);

    // Kontrol eden kişi
    m_patternKontrolEden = QRegularExpression(
        R"(Periyodik\s*Kontrol\s*Yapan\s*[:=]\s*(.+?)(?:\n|T\.C\.))",
        QRegularExpression::CaseInsensitiveOption);

    // TC Kimlik No
    m_patternTc = QRegularExpression(
        R"(T\.?C\.?\s*(?:Kimlik)?\s*(?:No|Numaras[ıi])?\s*[:=]?\s*(\d{11}))",
        QRegularExpression::CaseInsensitiveOption);

    // PK No
    m_patternPkNo = QRegularExpression(
        R"((?:PK|Periyodik\s*Kontrol)\s*(?:No|Numaras[ıi])\s*[:=]\s*([A-Z0-9\-]+))",
        QRegularExpression::CaseInsensitiveOption);

    // Tesisat sayısı
    m_patternTesisatSayisi = QRegularExpression(
        R"(Tesisat\s*Say[ıi]s[ıi]\s*[:=]\s*(\d+))",
        QRegularExpression::CaseInsensitiveOption);
}

QString PdfParser::extractText(const QString& pdfPath) {
    if (!QFile::exists(pdfPath)) {
        m_errorString = QString("PDF dosyası bulunamadı: %1").arg(pdfPath);
        return {};
    }

    QPdfDocument doc;
    QPdfDocument::Error error = doc.load(pdfPath);

    if (error != QPdfDocument::Error::None) {
        m_errorString = QString("PDF açılamadı: %1").arg(pdfPath);
        return {};
    }

    QString fullText;

    for (int i = 0; i < doc.pageCount(); ++i) {
        // Qt6.4+ için: QPdfDocument::getAllText()
        // Önceki sürümler için alternatif gerekebilir
        QPdfSelection selection = doc.getAllText(i);
        fullText += selection.text() + "\n";
    }

    return fullText;
}

QString PdfParser::findValue(const QString& text, const QRegularExpression& pattern) {
    QRegularExpressionMatch match = pattern.match(text);
    if (match.hasMatch() && match.lastCapturedIndex() >= 1) {
        return match.captured(1).trimmed();
    }
    return {};
}

QString PdfParser::formatSgkNo(const QString& sgkNo) {
    // Sadece rakamları al
    QString digits;
    for (const QChar& c : sgkNo) {
        if (c.isDigit()) {
            digits += c;
        }
    }

    // Python version format: 5-8-5-6 (or remainder)
    if (digits.length() >= 18) {
        return QString("%1-%2-%3-%4")
            .arg(digits.mid(0, 5))
            .arg(digits.mid(5, 8))
            .arg(digits.mid(13, 5))
            .arg(digits.mid(18));
    }

    return sgkNo;  // Formatlanamazsa olduğu gibi döndür
}

SozlesmeData PdfParser::parseSozlesme(const QString& pdfPath) {
    SozlesmeData data;

    // Önce Python script ile dene (daha güvenilir)
    data = parseSozlesmeViaPython(pdfPath);
    if (!data.sozlesmeId.isEmpty() || !data.firmaAdi.isEmpty()) {
        return data;
    }

    // Python başarısız olursa Qt ile dene
    QString text = extractText(pdfPath);
    if (text.isEmpty()) {
        return data;
    }

    // Tüm pattern'leri uygula
    data.sozlesmeId = findValue(text, m_patternSozlesmeId);
    data.firmaAdi = findValue(text, m_patternFirmaAdi);
    data.sgkSicil = formatSgkNo(findValue(text, m_patternSgk));
    data.kontrolEdenAdSoyad = findValue(text, m_patternKontrolEden);
    data.kontrolEdenTc = findValue(text, m_patternTc);
    data.pkNo = findValue(text, m_patternPkNo);

    QString tesisatStr = findValue(text, m_patternTesisatSayisi);
    if (!tesisatStr.isEmpty()) {
        data.tesisatSayisi = tesisatStr.toInt();
        if (data.tesisatSayisi < 1) data.tesisatSayisi = 1;
    }

    return data;
}

SozlesmeData PdfParser::parseSozlesmeViaPython(const QString& pdfPath) {
    SozlesmeData data;

    // Python script yolunu bul
    QString exePath = QCoreApplication::applicationDirPath();
    QString scriptPath = exePath + "/../../../sozlesme_parser.py";

    // Alternatif yollar
    if (!QFile::exists(scriptPath)) {
        scriptPath = exePath + "/../../sozlesme_parser.py";
    }
    if (!QFile::exists(scriptPath)) {
        scriptPath = "d:/YAPAY ZEKALILAR/rapor_sistemi/sozlesme_parser.py";
    }

    if (!QFile::exists(scriptPath)) {
        qDebug() << "Python parser bulunamadı:" << scriptPath;
        return data;
    }

    // Python çalıştır ve JSON çıktısı al
    QProcess process;
    process.start("python", {scriptPath, pdfPath});

    if (!process.waitForFinished(30000)) {
        qDebug() << "Python timeout";
        return data;
    }

    QString output = QString::fromUtf8(process.readAllStandardOutput());
    QString errorOutput = QString::fromUtf8(process.readAllStandardError());

    if (!errorOutput.isEmpty()) {
        qDebug() << "Python error:" << errorOutput;
    }

    // Çıktıyı parse et (key: value formatı)
    QStringList lines = output.split('\n');
    for (const QString& line : lines) {
        if (line.contains(':')) {
            int colonPos = line.indexOf(':');
            QString key = line.left(colonPos).trimmed();
            QString value = line.mid(colonPos + 1).trimmed();

            if (key == "sozlesme_id") data.sozlesmeId = value;
            else if (key == "sozlesme_baslangic") data.sozlesmeBaslangic = value;
            else if (key == "sozlesme_bitis") data.sozlesmeBitis = value;
            else if (key == "firma_unvan") data.firmaAdi = value;
            else if (key == "firma_adres") data.firmaAdres = value;
            else if (key == "firma_il") data.firmaIl = value;
            else if (key == "firma_sgk_no") data.sgkSicil = formatSgkNo(value);
            else if (key == "kontrol_eden_adsoyad") data.kontrolEdenAdSoyad = value;
            else if (key == "kontrol_eden_tc") data.kontrolEdenTc = value;
            else if (key == "pk_no") data.pkNo = value;
            else if (key == "tesisat_sayisi" && !value.isEmpty()) {
                data.tesisatSayisi = value.toInt();
                if (data.tesisatSayisi < 1) data.tesisatSayisi = 1;
            }
        }
    }

    return data;
}

} // namespace RaporSistemi

/**
 * PdfParser.h
 *
 * PDF sözleşme dosyasını parse eden modül.
 * Qt6 QPdfDocument kullanır.
 *
 * Python karşılığı: sozlesme_parser.py
 */

#ifndef PDFPARSER_H
#define PDFPARSER_H

#include <QString>
#include <QRegularExpression>
#include <memory>

class QPdfDocument;

namespace RaporSistemi {

/**
 * Sözleşme verileri
 */
struct SozlesmeData {
    QString sozlesmeId;
    QString sozlesmeBaslangic;
    QString sozlesmeBitis;
    QString firmaAdi;
    QString firmaAdres;
    QString firmaIl;
    QString sgkSicil;
    QString kontrolEdenAdSoyad;
    QString kontrolEdenTc;
    QString pkNo;              // Periyodik kontrol numarası
    int tesisatSayisi = 1;

    bool isValid() const {
        return !firmaAdi.isEmpty() && !sozlesmeId.isEmpty();
    }
};

class PdfParser {
public:
    PdfParser();
    ~PdfParser();

    /**
     * İSG-KATİP sözleşme PDF'ini parse eder.
     * @param pdfPath PDF dosyası yolu
     * @return Parse edilen veriler
     */
    SozlesmeData parseSozlesme(const QString& pdfPath);

    /**
     * Hata mesajı.
     */
    QString errorString() const { return m_errorString; }

private:
    /**
     * PDF'den tüm metni çıkarır.
     */
    QString extractText(const QString& pdfPath);

    /**
     * Regex ile değer bulur.
     */
    QString findValue(const QString& text, const QRegularExpression& pattern);

    /**
     * SGK numarasını formatlar.
     */
    QString formatSgkNo(const QString& sgkNo);

    QString m_errorString;

    // Önceden derlenmiş regex pattern'leri (performans için)
    QRegularExpression m_patternSozlesmeId;
    QRegularExpression m_patternFirmaAdi;
    QRegularExpression m_patternSgk;
    QRegularExpression m_patternKontrolEden;
    QRegularExpression m_patternTc;
    QRegularExpression m_patternPkNo;
    QRegularExpression m_patternTesisatSayisi;

    void initPatterns();

    /**
     * Python sozlesme_parser.py ile parse eder (daha güvenilir).
     */
    SozlesmeData parseSozlesmeViaPython(const QString& pdfPath);
};

} // namespace RaporSistemi

#endif // PDFPARSER_H

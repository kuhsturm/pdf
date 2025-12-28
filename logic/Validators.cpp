/**
 * Validators.cpp
 */

#include "Validators.h"
#include <QRegularExpression>

namespace RaporSistemi {

bool validateFonksiyonTesti(const FonksiyonTesti& test, QString* reason) {
    // KAKR grubu kontrolü (Python: is_kakr_group = "KAKR" in linye_adi.upper())
    // KAKR grupları için In>Iz kontrolü yapılmaz
    bool isKakrGroup = test.linye.toUpper().contains("KAKR");

    // 1. In > Iz kontrolü (KAKR grupları HARİÇ - Python ile uyumlu)
    if (!isKakrGroup && test.akimKapasitesi > 0 && test.nominalAkim > test.akimKapasitesi) {
        if (reason) *reason = QString("In (%1A) > Iz (%2A)")
            .arg(test.nominalAkim).arg(test.akimKapasitesi);
        return false;
    }

    // NOT: Python'da Ib > Iz kontrolü YOK - bu yüzden kaldırıldı
    // Python sadece In > Iz kontrolü yapıyor

    // 3. 32A altı KAKR kontrolü (Python ile uyumlu)
    // Python: has_30ma_kakr = kakr_checked and rcd_acma_val == "30mA"
    // ANA SİGORTA için bu kural geçerli değil (Python: if not is_ana_sigorta)
    if (!test.isAnaSigorta && test.nominalAkim > 0 && test.nominalAkim <= 32) {
        // Python mantığı: kakrVar VE rcd == "30mA" olmalı
        bool has30maKakr = test.kakrVar && test.rcd == "30mA";
        if (!has30maKakr && !test.kakrYok) {
            if (reason) *reason = "32A altı devreler için 30mA KAKR gerekli";
            return false;
        }
    }

    // 4. KAKR Yok işaretli
    if (test.kakrYok) {
        if (reason) *reason = "KAKR bulunmuyor";
        return false;
    }

    return true;
}

QVector<Kusur> validateAllTests(const QVector<FonksiyonTesti>& tests) {
    QVector<Kusur> kusurlar;

    for (const auto& test : tests) {
        QString reason;
        if (!validateFonksiyonTesti(test, &reason)) {
            Kusur k;
            k.linye = test.linye;
            k.kusurAciklamasi = reason;
            k.kusurDerecesi = "K2";  // Varsayılan
            kusurlar.append(k);
        }
    }

    return kusurlar;
}

QString determineGenelSonuc(const QVector<FonksiyonTesti>& tests) {
    for (const auto& test : tests) {
        if (!validateFonksiyonTesti(test)) {
            return "Uygun Değil";
        }
    }
    return "Uygun";
}

bool validateRaporNumarasi(const QString& raporNo) {
    // Format: TPK2024-1234 veya TPK2024-1234-1
    static QRegularExpression pattern(
        R"(^[A-Z]{2,4}\d{4}-\d{1,6}(-\d+)?$)",
        QRegularExpression::CaseInsensitiveOption);

    return pattern.match(raporNo).hasMatch();
}

bool validateTcKimlik(const QString& tc) {
    // 11 haneli olmalı
    if (tc.length() != 11) return false;

    // Tamamı rakam olmalı
    for (QChar c : tc) {
        if (!c.isDigit()) return false;
    }

    // İlk hane 0 olamaz
    if (tc[0] == '0') return false;

    // Algoritma kontrolü
    int digits[11];
    for (int i = 0; i < 11; ++i) {
        digits[i] = tc[i].digitValue();
    }

    // 10. hane kontrolü
    int sum1 = digits[0] + digits[2] + digits[4] + digits[6] + digits[8];
    int sum2 = digits[1] + digits[3] + digits[5] + digits[7];
    int check10 = (sum1 * 7 - sum2) % 10;
    if (check10 < 0) check10 += 10;

    if (digits[9] != check10) return false;

    // 11. hane kontrolü
    int sumAll = 0;
    for (int i = 0; i < 10; ++i) {
        sumAll += digits[i];
    }

    if (digits[10] != sumAll % 10) return false;

    return true;
}

} // namespace RaporSistemi

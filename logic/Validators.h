/**
 * Validators.h
 *
 * Doğrulama fonksiyonları.
 */

#ifndef VALIDATORS_H
#define VALIDATORS_H

#include <QString>
#include <QVector>
#include "DataModels.h"

namespace RaporSistemi {

/**
 * Fonksiyon testi satırını doğrular.
 * @return true = Uygun, false = Uygun Değil
 */
bool validateFonksiyonTesti(const FonksiyonTesti& test, QString* reason = nullptr);

/**
 * Tüm fonksiyon testlerini doğrular ve kusur listesi oluşturur.
 */
QVector<Kusur> validateAllTests(const QVector<FonksiyonTesti>& tests);

/**
 * Genel sonucu belirler.
 */
QString determineGenelSonuc(const QVector<FonksiyonTesti>& tests);

/**
 * Rapor numarası formatını doğrular.
 * Format: TPK2024-XXXX veya TPK2024-XXXX-N
 */
bool validateRaporNumarasi(const QString& raporNo);

/**
 * TC Kimlik numarasını doğrular (11 haneli, algoritma kontrolü).
 */
bool validateTcKimlik(const QString& tc);

} // namespace RaporSistemi

#endif // VALIDATORS_H

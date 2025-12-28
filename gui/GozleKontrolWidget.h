/**
 * GozleKontrolWidget.h
 *
 * Gözle Kontrol sekmesi - 27 kontrol maddesi
 * Python karşılığı: create_gk_tab() in multi_pano_gui.py
 */

#ifndef GOZLEKONTROLWIDGET_H
#define GOZLEKONTROLWIDGET_H

#include <QWidget>
#include <QComboBox>
#include <QMap>
#include <QString>
#include <QStringList>

namespace RaporSistemi {

class GozleKontrolWidget : public QWidget {
    Q_OBJECT

public:
    explicit GozleKontrolWidget(QWidget* parent = nullptr);

    /**
     * Tüm kontrol değerlerini döndürür.
     */
    QMap<QString, QString> getData() const;

    /**
     * Kontrol değerlerini ayarlar.
     */
    void setData(const QMap<QString, QString>& data);

    /**
     * Tüm alanları varsayılan değere döndürür.
     */
    void clear();

    /**
     * Kritik alanlardan (Zemin İzolasyonu, Pano Kapak, Aşırı Yük)
     * herhangi biri "Uygun Değil" mi kontrol eder.
     */
    bool hasUygunDegilCriticalField() const;

signals:
    void dataChanged();
    void criticalFieldChanged();  // Kritik GK alanı değiştiğinde

private:
    void setupUi();

    // 27 kontrol maddesi için ComboBox'lar
    QMap<QString, QComboBox*> m_entries;

    // Sabit kontrol maddeleri listesi
    static const QStringList KONTROL_MADDELERI;
    static const QStringList DEGERLENDIRME_VALUES;
};

} // namespace RaporSistemi

#endif // GOZLEKONTROLWIDGET_H

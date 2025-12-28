/**
 * LinieGroupDialog.h
 *
 * Linye grubu ekleme dialog'u.
 * Python karşılığı: add_multiple_ft_rows() in PanoDataFrame
 */

#ifndef LINIEGROUPDIALOG_H
#define LINIEGROUPDIALOG_H

#include <QDialog>
#include <QLineEdit>
#include <QSpinBox>
#include <QComboBox>
#include <QCheckBox>
#include <QRegularExpression>
#include "DataModels.h"

namespace RaporSistemi {

class LinieGroupDialog : public QDialog {
    Q_OBJECT

public:
    explicit LinieGroupDialog(QWidget* parent = nullptr);

    /**
     * Oluşturulan fonksiyon testlerini döndürür.
     */
    QVector<FonksiyonTesti> getTests() const { return m_tests; }

public slots:
    void addGroup();

private:
    void setupUi();

    // Linye adı ve adet
    QLineEdit* m_linyePrefix;
    QSpinBox* m_count;

    // Linye sigorta & kesit
    QComboBox* m_sigortaTipi;  // Eğri
    QComboBox* m_kutup;        // Kutup sayısı
    QComboBox* m_nominalAkim;  // In (A)
    QComboBox* m_icu;          // Icu (kA)
    QComboBox* m_fazKesiti;
    QComboBox* m_notrKesiti;
    QComboBox* m_toprakKesiti;

    // KAKR bilgileri
    QCheckBox* m_kakrVar;
    QComboBox* m_kakrIn;
    QComboBox* m_kakrIcu;
    QComboBox* m_kakrFazKesiti;
    QComboBox* m_kakrNotrKesiti;
    QComboBox* m_kakrToprakKesiti;
    QComboBox* m_rcd;
    QLineEdit* m_rcdMa;
    QLineEdit* m_rcdMs;

    QVector<FonksiyonTesti> m_tests;

private slots:
    void onInChanged(const QString& in);
    void onKakrInChanged(const QString& in);
};

} // namespace RaporSistemi

#endif // LINIEGROUPDIALOG_H

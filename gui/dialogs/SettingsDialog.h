/**
 * SettingsDialog.h
 *
 * Uygulama ayarları dialog'u.
 */

#ifndef SETTINGSDIALOG_H
#define SETTINGSDIALOG_H

#include <QDialog>
#include <QLineEdit>
#include <QCheckBox>

namespace RaporSistemi {

class SettingsDialog : public QDialog {
    Q_OBJECT

public:
    explicit SettingsDialog(QWidget* parent = nullptr);

public slots:
    void loadSettings();
    void saveSettings();

private:
    void setupUi();

    QLineEdit* m_templatePath;
    QLineEdit* m_outputPath;
    QLineEdit* m_kisiExcelPath;
    QCheckBox* m_autoSave;
    QCheckBox* m_darkMode;
};

} // namespace RaporSistemi

#endif // SETTINGSDIALOG_H

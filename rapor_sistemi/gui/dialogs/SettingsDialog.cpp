/**
 * SettingsDialog.cpp
 */

#include "SettingsDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QLabel>
#include <QPushButton>
#include <QFileDialog>
#include <QSettings>
#include <QGroupBox>

namespace RaporSistemi {

SettingsDialog::SettingsDialog(QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Ayarlar"));
    setupUi();
    loadSettings();
}

void SettingsDialog::setupUi() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    // Dosya yolları
    QGroupBox* pathGroup = new QGroupBox(tr("Dosya Yolları"));
    QGridLayout* pathGrid = new QGridLayout(pathGroup);

    pathGrid->addWidget(new QLabel(tr("Şablon Dosyası:")), 0, 0);
    m_templatePath = new QLineEdit();
    pathGrid->addWidget(m_templatePath, 0, 1);
    QPushButton* browseTemplate = new QPushButton(tr("..."));
    browseTemplate->setMaximumWidth(30);
    connect(browseTemplate, &QPushButton::clicked, [this]() {
        QString path = QFileDialog::getOpenFileName(this, tr("Şablon Seç"),
            QString(), tr("Word Dosyaları (*.docx)"));
        if (!path.isEmpty()) m_templatePath->setText(path);
    });
    pathGrid->addWidget(browseTemplate, 0, 2);

    pathGrid->addWidget(new QLabel(tr("Çıktı Klasörü:")), 1, 0);
    m_outputPath = new QLineEdit();
    pathGrid->addWidget(m_outputPath, 1, 1);
    QPushButton* browseOutput = new QPushButton(tr("..."));
    browseOutput->setMaximumWidth(30);
    connect(browseOutput, &QPushButton::clicked, [this]() {
        QString path = QFileDialog::getExistingDirectory(this, tr("Çıktı Klasörü Seç"));
        if (!path.isEmpty()) m_outputPath->setText(path);
    });
    pathGrid->addWidget(browseOutput, 1, 2);

    pathGrid->addWidget(new QLabel(tr("Kişi Bilgileri Excel:")), 2, 0);
    m_kisiExcelPath = new QLineEdit();
    pathGrid->addWidget(m_kisiExcelPath, 2, 1);

    mainLayout->addWidget(pathGroup);

    // Genel ayarlar
    QGroupBox* generalGroup = new QGroupBox(tr("Genel"));
    QVBoxLayout* generalLayout = new QVBoxLayout(generalGroup);

    m_autoSave = new QCheckBox(tr("Otomatik kaydet (5 dakikada bir)"));
    generalLayout->addWidget(m_autoSave);

    m_darkMode = new QCheckBox(tr("Koyu tema"));
    m_darkMode->setChecked(true);
    generalLayout->addWidget(m_darkMode);

    mainLayout->addWidget(generalGroup);

    // Butonlar
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    buttonLayout->addStretch();

    QPushButton* saveBtn = new QPushButton(tr("Kaydet"));
    connect(saveBtn, &QPushButton::clicked, this, &SettingsDialog::saveSettings);
    connect(saveBtn, &QPushButton::clicked, this, &QDialog::accept);
    buttonLayout->addWidget(saveBtn);

    QPushButton* cancelBtn = new QPushButton(tr("İptal"));
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
    buttonLayout->addWidget(cancelBtn);

    mainLayout->addLayout(buttonLayout);

    setMinimumWidth(500);
}

void SettingsDialog::loadSettings() {
    QSettings settings("RaporSistemi", "ElektrikRaporSistemi");

    m_templatePath->setText(settings.value("templatePath").toString());
    m_outputPath->setText(settings.value("outputPath").toString());
    m_kisiExcelPath->setText(settings.value("kisiExcelPath").toString());
    m_autoSave->setChecked(settings.value("autoSave", true).toBool());
    m_darkMode->setChecked(settings.value("darkMode", true).toBool());
}

void SettingsDialog::saveSettings() {
    QSettings settings("RaporSistemi", "ElektrikRaporSistemi");

    settings.setValue("templatePath", m_templatePath->text());
    settings.setValue("outputPath", m_outputPath->text());
    settings.setValue("kisiExcelPath", m_kisiExcelPath->text());
    settings.setValue("autoSave", m_autoSave->isChecked());
    settings.setValue("darkMode", m_darkMode->isChecked());
}

} // namespace RaporSistemi

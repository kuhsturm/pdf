/**
 * GozleKontrolWidget.cpp
 *
 * Python karşılığı: create_gk_tab() in PanoDataFrame
 */

#include "GozleKontrolWidget.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QScrollArea>
#include <QFrame>

namespace RaporSistemi {

// 27 kontrol maddesi (Python ile birebir aynı)
const QStringList GozleKontrolWidget::KONTROL_MADDELERI = {
    "Kablo Sebeke Tarafi", "Kablo Donanim Tarafi",
    "Pano Sabitlenmesi", "Dis Darbelere Karsi Koruma Onlemi",
    "Elektrik Panosu Etrafinda Yabanci Malzemeler", "Zemin Izolasyonu",
    "Topraklama Iletkeni", "Ana Potansiyel Dengeleme Iletkeni",
    "Ek Potansiyel Dengeleme Iletkeni", "Pano Kapak Baglantisi Kontrolu 6 mm2",
    "Elektriksel Olmayan Tesislere Yaklasma", "Bant Ayrilmasi",
    "Guvenlik Devre Ayrilmasi", "Pano Ic Kapak",
    "Semalar Talimatlar", "Koruma Cihaz ve Terminal Etiket",
    "Tehlike Isaretleri", "Kablo Yollari",
    "Kablo Renk Kodlari", "Tesisat Yontemi",
    "Yangin Engeli", "Kontak Gevsekligi Isinmasi",
    "Asiri Yuk Isinmasi", "Yangin Sondurme",
    "Ekipman Temizlik", "Korozyon Kontrolu",
    "Acil Durum Aydinlatma"
};

const QStringList GozleKontrolWidget::DEGERLENDIRME_VALUES = {
    "Uygun", "Uygun Değil", "Uygulanamaz"
};

GozleKontrolWidget::GozleKontrolWidget(QWidget* parent)
    : QWidget(parent)
{
    setupUi();
}

void GozleKontrolWidget::setupUi() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(4, 4, 4, 4);
    mainLayout->setSpacing(2);

    // Scrollable area
    QScrollArea* scrollArea = new QScrollArea();
    scrollArea->setWidgetResizable(true);
    scrollArea->setFrameShape(QFrame::NoFrame);

    QWidget* scrollContent = new QWidget();
    QVBoxLayout* contentLayout = new QVBoxLayout(scrollContent);
    contentLayout->setContentsMargins(4, 4, 4, 4);
    contentLayout->setSpacing(4);

    // Başlık
    QLabel* header = new QLabel(tr("<b>Gözle Kontrol Maddeleri (27 Madde)</b>"));
    header->setStyleSheet("color: #4CAF50; font-size: 12px;");
    contentLayout->addWidget(header);

    // 2'li gruplar halinde satırlar oluştur (Python ile aynı)
    for (int i = 0; i < KONTROL_MADDELERI.size(); i += 2) {
        QHBoxLayout* rowLayout = new QHBoxLayout();
        rowLayout->setSpacing(8);

        // Sol madde
        QString leftField = KONTROL_MADDELERI[i];

        QLabel* leftLabel = new QLabel(leftField.left(30) + ":");
        leftLabel->setMinimumWidth(180);
        leftLabel->setMaximumWidth(200);
        rowLayout->addWidget(leftLabel);

        if (leftField == "Tesisat Yontemi") {
            // Sabit değer: A1
            QLabel* fixedLabel = new QLabel("A1");
            fixedLabel->setStyleSheet("color: #4CAF50; font-weight: bold;");
            fixedLabel->setMinimumWidth(100);
            rowLayout->addWidget(fixedLabel);
            // Sabit değeri saklamak için görünmez combo
            QComboBox* combo = new QComboBox();
            combo->addItem("A1");
            combo->setVisible(false);
            m_entries[leftField] = combo;
        } else {
            QComboBox* leftCombo = new QComboBox();
            leftCombo->addItems(DEGERLENDIRME_VALUES);
            leftCombo->setCurrentText("Uygun");
            leftCombo->setMinimumWidth(100);
            leftCombo->setMaximumWidth(120);
            connect(leftCombo, &QComboBox::currentTextChanged, this, &GozleKontrolWidget::dataChanged);
            // Kritik alanlar için ek sinyal
            if (leftField == "Zemin Izolasyonu" ||
                leftField == "Pano Kapak Baglantisi Kontrolu 6 mm2" ||
                leftField == "Asiri Yuk Isinmasi") {
                connect(leftCombo, &QComboBox::currentTextChanged, this, &GozleKontrolWidget::criticalFieldChanged);
            }
            rowLayout->addWidget(leftCombo);
            m_entries[leftField] = leftCombo;
        }

        // Sağ madde (varsa)
        if (i + 1 < KONTROL_MADDELERI.size()) {
            QString rightField = KONTROL_MADDELERI[i + 1];

            // Boşluk
            rowLayout->addSpacing(20);

            QLabel* rightLabel = new QLabel(rightField.left(30) + ":");
            rightLabel->setMinimumWidth(180);
            rightLabel->setMaximumWidth(200);
            rowLayout->addWidget(rightLabel);

            if (rightField == "Tesisat Yontemi") {
                QLabel* fixedLabel = new QLabel("A1");
                fixedLabel->setStyleSheet("color: #4CAF50; font-weight: bold;");
                fixedLabel->setMinimumWidth(100);
                rowLayout->addWidget(fixedLabel);
                QComboBox* combo = new QComboBox();
                combo->addItem("A1");
                combo->setVisible(false);
                m_entries[rightField] = combo;
            } else {
                QComboBox* rightCombo = new QComboBox();
                rightCombo->addItems(DEGERLENDIRME_VALUES);
                rightCombo->setCurrentText("Uygun");
                rightCombo->setMinimumWidth(100);
                rightCombo->setMaximumWidth(120);
                connect(rightCombo, &QComboBox::currentTextChanged, this, &GozleKontrolWidget::dataChanged);
                // Kritik alanlar için ek sinyal
                if (rightField == "Zemin Izolasyonu" ||
                    rightField == "Pano Kapak Baglantisi Kontrolu 6 mm2" ||
                    rightField == "Asiri Yuk Isinmasi") {
                    connect(rightCombo, &QComboBox::currentTextChanged, this, &GozleKontrolWidget::criticalFieldChanged);
                }
                rowLayout->addWidget(rightCombo);
                m_entries[rightField] = rightCombo;
            }
        }

        rowLayout->addStretch();
        contentLayout->addLayout(rowLayout);
    }

    contentLayout->addStretch();
    scrollArea->setWidget(scrollContent);
    mainLayout->addWidget(scrollArea);
}

QMap<QString, QString> GozleKontrolWidget::getData() const {
    QMap<QString, QString> data;
    for (auto it = m_entries.constBegin(); it != m_entries.constEnd(); ++it) {
        data[it.key()] = it.value()->currentText();
    }
    return data;
}

void GozleKontrolWidget::setData(const QMap<QString, QString>& data) {
    for (auto it = data.constBegin(); it != data.constEnd(); ++it) {
        if (m_entries.contains(it.key())) {
            int idx = m_entries[it.key()]->findText(it.value());
            if (idx >= 0) {
                m_entries[it.key()]->setCurrentIndex(idx);
            }
        }
    }
}

void GozleKontrolWidget::clear() {
    for (auto it = m_entries.begin(); it != m_entries.end(); ++it) {
        if (it.key() == "Tesisat Yontemi") {
            continue; // Sabit değer
        }
        it.value()->setCurrentIndex(0); // "Uygun"
    }
}

bool GozleKontrolWidget::hasUygunDegilCriticalField() const {
    // Python'daki kritik alanlar:
    // - Zemin Izolasyonu
    // - Pano Kapak Baglantisi Kontrolu 6 mm2
    // - Asiri Yuk Isinmasi
    static const QStringList criticalFields = {
        "Zemin Izolasyonu",
        "Pano Kapak Baglantisi Kontrolu 6 mm2",
        "Asiri Yuk Isinmasi"
    };

    for (const QString& field : criticalFields) {
        if (m_entries.contains(field)) {
            if (m_entries[field]->currentText() == QString::fromUtf8("Uygun Değil")) {
                return true;
            }
        }
    }
    return false;
}

} // namespace RaporSistemi

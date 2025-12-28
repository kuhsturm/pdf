#include "LinieGroupDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QLabel>
#include <QPushButton>
#include <QGroupBox>
#include <QCheckBox>
#include <QScrollArea>
#include <QMessageBox>
#include <QMap>

namespace RaporSistemi {

LinieGroupDialog::LinieGroupDialog(QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Linye Grubu Ekle"));
    setMinimumSize(580, 800);
    setupUi();
}

void LinieGroupDialog::setupUi() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    QScrollArea* scroll = new QScrollArea();
    scroll->setWidgetResizable(true);
    scroll->setFrameShape(QFrame::NoFrame);

    QWidget* content = new QWidget();
    QVBoxLayout* contentLayout = new QVBoxLayout(content);

    // ===== GENEL AYARLAR =====
    QGroupBox* generalGroup = new QGroupBox(tr("Genel Ayarlar"));
    QGridLayout* genLayout = new QGridLayout(generalGroup);

    genLayout->addWidget(new QLabel(tr("Kaç Adet:")), 0, 0);
    m_count = new QSpinBox();
    m_count->setRange(1, 100);
    m_count->setValue(10);
    m_count->setFixedWidth(100);
    genLayout->addWidget(m_count, 0, 1);

    genLayout->addWidget(new QLabel(tr("Linye Şablonu:")), 1, 0);
    m_linyePrefix = new QLineEdit();
    m_linyePrefix->setPlaceholderText("Aydınlatma 1");
    m_linyePrefix->setText("Aydınlatma 1");
    genLayout->addWidget(m_linyePrefix, 1, 1);

    QLabel* hint = new QLabel(tr("(Sondaki sayı otomatik artar)"));
    hint->setStyleSheet("color: gray; font-size: 10px;");
    genLayout->addWidget(hint, 2, 1);

    contentLayout->addWidget(generalGroup);

    // ===== LİNYE SİGORTA & KESİT =====
    QGroupBox* linyeGroup = new QGroupBox();
    linyeGroup->setTitle(tr("━━━ LİNYE SİGORTA & KESİT ━━━"));
    linyeGroup->setStyleSheet("QGroupBox::title { color: #4CAF50; font-weight: bold; }");
    QGridLayout* linyeLayout = new QGridLayout(linyeGroup);

    linyeLayout->addWidget(new QLabel(tr("Eğri:")), 0, 0);
    m_sigortaTipi = new QComboBox();
    m_sigortaTipi->addItems({"B", "C", "D", "K", "Z"});
    m_sigortaTipi->setCurrentText("C");
    linyeLayout->addWidget(m_sigortaTipi, 0, 1);

    linyeLayout->addWidget(new QLabel(tr("Kutup:")), 0, 2);
    m_kutup = new QComboBox();
    m_kutup->addItems({"1", "2", "3", "4"});
    linyeLayout->addWidget(m_kutup, 0, 3);

    linyeLayout->addWidget(new QLabel(tr("In (A):")), 1, 0);
    m_nominalAkim = new QComboBox();
    m_nominalAkim->addItems({"1", "2", "3", "4", "6", "10", "13", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125", "160", "200", "250", "315", "400"});
    m_nominalAkim->setCurrentText("16");
    connect(m_nominalAkim, &QComboBox::currentTextChanged, this, &LinieGroupDialog::onInChanged);
    linyeLayout->addWidget(m_nominalAkim, 1, 1);

    linyeLayout->addWidget(new QLabel(tr("Icu (kA):")), 1, 2);
    m_icu = new QComboBox();
    m_icu->addItems({"3kA", "4.5kA", "6kA", "10kA", "25kA", "35kA", "55kA", "70kA"});
    m_icu->setCurrentText("6kA");
    linyeLayout->addWidget(m_icu, 1, 3);

    QHBoxLayout* kesitRow = new QHBoxLayout();
    QStringList kesitler = {"1.5", "2.5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120", "150", "185", "240"};

    m_fazKesiti = new QComboBox(); m_fazKesiti->addItems(kesitler); m_fazKesiti->setCurrentText("2.5");
    m_notrKesiti = new QComboBox(); m_notrKesiti->addItems(kesitler); m_notrKesiti->setCurrentText("2.5");
    m_toprakKesiti = new QComboBox(); m_toprakKesiti->addItems(kesitler); m_toprakKesiti->setCurrentText("2.5");

    kesitRow->addWidget(new QLabel(tr("Faz:"))); kesitRow->addWidget(m_fazKesiti);
    kesitRow->addWidget(new QLabel(tr("N:"))); kesitRow->addWidget(m_notrKesiti);
    kesitRow->addWidget(new QLabel(tr("PE:"))); kesitRow->addWidget(m_toprakKesiti);
    linyeLayout->addLayout(kesitRow, 2, 0, 1, 4);

    contentLayout->addWidget(linyeGroup);

    // ===== KAKR BİLGİLERİ =====
    QGroupBox* kakrGroup = new QGroupBox();
    kakrGroup->setTitle(tr("━━━ KAKR BİLGİLERİ ━━━"));
    kakrGroup->setStyleSheet("QGroupBox::title { color: #FF9800; font-weight: bold; }");
    QGridLayout* kakrLayout = new QGridLayout(kakrGroup);

    m_kakrVar = new QCheckBox(tr("KAKR Var (30mA)"));
    m_kakrVar->setChecked(true);
    kakrLayout->addWidget(m_kakrVar, 0, 0, 1, 4);

    kakrLayout->addWidget(new QLabel(tr("In (A):")), 1, 0);
    m_kakrIn = new QComboBox();
    m_kakrIn->addItems({"25", "32", "40", "50", "63", "80", "100", "125"});
    m_kakrIn->setCurrentText("40");
    connect(m_kakrIn, &QComboBox::currentTextChanged, this, &LinieGroupDialog::onKakrInChanged);
    kakrLayout->addWidget(m_kakrIn, 1, 1);

    kakrLayout->addWidget(new QLabel(tr("Icu (kA):")), 1, 2);
    m_kakrIcu = new QComboBox();
    m_kakrIcu->addItems({"6kA", "10kA", "25kA", "35kA"});
    m_kakrIcu->setCurrentText("10kA");
    kakrLayout->addWidget(m_kakrIcu, 1, 3);

    QHBoxLayout* kakrKesitRow = new QHBoxLayout();
    m_kakrFazKesiti = new QComboBox(); m_kakrFazKesiti->addItems(kesitler); m_kakrFazKesiti->setCurrentText("6");
    m_kakrNotrKesiti = new QComboBox(); m_kakrNotrKesiti->addItems(kesitler); m_kakrNotrKesiti->setCurrentText("6");
    m_kakrToprakKesiti = new QComboBox(); m_kakrToprakKesiti->addItems(kesitler); m_kakrToprakKesiti->setCurrentText("6");

    kakrKesitRow->addWidget(new QLabel(tr("Faz:"))); kakrKesitRow->addWidget(m_kakrFazKesiti);
    kakrKesitRow->addWidget(new QLabel(tr("N:"))); kakrKesitRow->addWidget(m_kakrNotrKesiti);
    kakrKesitRow->addWidget(new QLabel(tr("PE:"))); kakrKesitRow->addWidget(m_kakrToprakKesiti);
    kakrLayout->addLayout(kakrKesitRow, 2, 0, 1, 4);

    contentLayout->addWidget(kakrGroup);

    // ===== RCD TEST DEĞERLERİ =====
    QGroupBox* rcdGroup = new QGroupBox();
    rcdGroup->setTitle(tr("━━━ RCD TEST DEĞERLERİ ━━━"));
    rcdGroup->setStyleSheet("QGroupBox::title { color: #2196F3; font-weight: bold; }");
    QHBoxLayout* rcdLayout = new QHBoxLayout(rcdGroup);

    rcdLayout->addWidget(new QLabel(tr("IΔn:")));
    m_rcd = new QComboBox();
    m_rcd->addItems({"", "30mA", "100mA", "300mA"});
    m_rcd->setCurrentText("30mA");
    rcdLayout->addWidget(m_rcd);

    rcdLayout->addWidget(new QLabel(tr("mA:")));
    m_rcdMa = new QLineEdit();
    m_rcdMa->setFixedWidth(60);
    m_rcdMa->setText("25");
    rcdLayout->addWidget(m_rcdMa);

    rcdLayout->addWidget(new QLabel(tr("mS:")));
    m_rcdMs = new QLineEdit();
    m_rcdMs->setFixedWidth(60);
    m_rcdMs->setText("20");
    rcdLayout->addWidget(m_rcdMs);

    contentLayout->addWidget(rcdGroup);

    contentLayout->addStretch();
    scroll->setWidget(content);
    mainLayout->addWidget(scroll, 1);

    // Butonlar
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    QPushButton* addBtn = new QPushButton(tr("Ekle ve Devam"));
    addBtn->setStyleSheet("background-color: #2E7D32; color: white; font-weight: bold; padding: 10px 20px; border-radius: 4px;");
    connect(addBtn, &QPushButton::clicked, this, &LinieGroupDialog::addGroup);
    buttonLayout->addWidget(addBtn);

    QPushButton* addCloseBtn = new QPushButton(tr("Ekle ve Kapat"));
    addCloseBtn->setStyleSheet("background-color: #1976D2; color: white; font-weight: bold; padding: 10px 20px; border-radius: 4px;");
    connect(addCloseBtn, &QPushButton::clicked, [this]() {
        addGroup();
        accept();
    });
    buttonLayout->addWidget(addCloseBtn);

    QPushButton* closeBtn = new QPushButton(tr("İptal"));
    closeBtn->setMinimumHeight(40);
    connect(closeBtn, &QPushButton::clicked, this, &QDialog::reject);
    buttonLayout->addWidget(closeBtn);

    mainLayout->addLayout(buttonLayout);
}

void LinieGroupDialog::onInChanged(const QString& in) {
    static const QMap<int, QString> inToKesit = {
        {1, "1.5"}, {2, "1.5"}, {3, "1.5"}, {4, "1.5"}, {6, "1.5"},
        {10, "1.5"}, {13, "2.5"}, {16, "2.5"}, {20, "4"}, {25, "4"},
        {32, "6"}, {40, "10"}, {50, "16"}, {63, "16"}, {80, "25"},
        {100, "35"}, {125, "50"}, {160, "70"}, {200, "95"}, {250, "120"},
        {315, "150"}, {400, "240"}
    };
    int val = in.toInt();
    if (inToKesit.contains(val)) {
        QString kesit = inToKesit[val];
        m_fazKesiti->setCurrentText(kesit);
        m_notrKesiti->setCurrentText(kesit);
        m_toprakKesiti->setCurrentText(kesit);
    }
}

void LinieGroupDialog::onKakrInChanged(const QString& in) {
    static const QMap<int, QString> inToKesit = {
        {25, "4"}, {32, "6"}, {40, "10"}, {50, "16"}, {63, "16"}, {80, "25"}, {100, "35"}, {125, "50"}
    };
    int val = in.toInt();
    if (inToKesit.contains(val)) {
        QString kesit = inToKesit[val];
        m_kakrFazKesiti->setCurrentText(kesit);
        m_kakrNotrKesiti->setCurrentText(kesit);
        m_kakrToprakKesiti->setCurrentText(kesit);
    }
}

void LinieGroupDialog::addGroup() {
    QString prefix = m_linyePrefix->text();
    if (prefix.isEmpty()) prefix = "Linye 1";

    int count = m_count->value();

    // Naming logic search for digits at the end
    QRegularExpression re("^(.*?)(\\d+)$");
    QRegularExpressionMatch match = re.match(prefix);

    QString baseName = prefix;
    int startNum = 1;

    if (match.hasMatch()) {
        baseName = match.captured(1);
        startNum = match.captured(2).toInt();
    }

    // 1. First add KAKR row if checked
    if (m_kakrVar->isChecked()) {
        // Construct KAKR name from linye name (Python like)
        QString kakrName = "KAKR";
        QString cleanBase = baseName.trimmed();
        if (!cleanBase.isEmpty()) {
            kakrName = "KAKR " + cleanBase;
        }

        FonksiyonTesti kakrRow;
        kakrRow.linye = kakrName;
        kakrRow.sigortaTipi = "AAA"; // Python: AAA
        kakrRow.kutupSayisi = 4;     // Python: 4
        kakrRow.nominalAkim = m_kakrIn->currentText().toInt();
        kakrRow.icu = m_kakrIcu->currentText();
        kakrRow.ib = QString::number(kakrRow.nominalAkim * 0.7, 'f', 1);
        kakrRow.fazKesiti = m_kakrFazKesiti->currentText();
        kakrRow.notrKesiti = m_kakrNotrKesiti->currentText();
        kakrRow.toprakKesiti = m_kakrToprakKesiti->currentText();

        kakrRow.kakrVar = true;
        kakrRow.rcd = m_rcd->currentText();
        kakrRow.rcdMa = m_rcdMa->text();
        kakrRow.rcdMs = m_rcdMs->text();
        kakrRow.sonuc = "Uygun";

        auto [k, c] = parseKesit(kakrRow.fazKesiti);
        kakrRow.akimKapasitesi = kesitToIz(k, c);

        m_tests.append(kakrRow);
    }

    // 2. Add linye rows
    for (int i = 0; i < count; ++i) {
        FonksiyonTesti test;
        test.linye = QString("%1%2").arg(baseName).arg(startNum + i);
        test.sigortaTipi = m_sigortaTipi->currentText();
        test.kutupSayisi = m_kutup->currentText().toInt();
        test.nominalAkim = m_nominalAkim->currentText().toInt();
        test.icu = m_icu->currentText();
        test.ib = QString::number(test.nominalAkim * 0.7, 'f', 1);
        test.fazKesiti = m_fazKesiti->currentText();
        test.notrKesiti = m_notrKesiti->currentText();
        test.toprakKesiti = m_toprakKesiti->currentText();

        // Rows under KAKR also reference it in Python (kakrVar = True)
        if (m_kakrVar->isChecked()) {
            test.kakrVar = true;
            test.rcd = m_rcd->currentText();
            test.rcdMa = m_rcdMa->text();
            test.rcdMs = m_rcdMs->text();
        }

        test.sonuc = "Uygun";
        auto [k, c] = parseKesit(test.fazKesiti);
        test.akimKapasitesi = kesitToIz(k, c);

        m_tests.append(test);
    }

    // Update prefix for next usage
    m_linyePrefix->setText(QString("%1%2").arg(baseName).arg(startNum + count));
}

} // namespace RaporSistemi

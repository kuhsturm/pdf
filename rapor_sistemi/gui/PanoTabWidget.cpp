/**
 * PanoTabWidget.cpp
 */

#include "PanoTabWidget.h"
#include "FonksiyonTestleriTable.h"
#include "DragDropWidget.h"
#include "GozleKontrolWidget.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLabel>
#include <QScrollArea>
#include <QSplitter>
#include <QPushButton>
#include <cmath>

namespace RaporSistemi {

PanoTabWidget::PanoTabWidget(int panoIndex, QWidget* parent)
    : QWidget(parent)
    , m_panoIndex(panoIndex)
{
    setupUi();
    setupConnections();
}

PanoTabWidget::~PanoTabWidget() = default;

void PanoTabWidget::setupUi() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(4, 4, 4, 4);
    mainLayout->setSpacing(4);

    // Pano Adı (üstte)
    QHBoxLayout* topLayout = new QHBoxLayout();
    topLayout->addWidget(new QLabel(tr("Pano Adı:")));
    m_panoAdi = new QLineEdit();
    m_panoAdi->setPlaceholderText(QString("Pano %1 Adı").arg(m_panoIndex));
    m_panoAdi->setMinimumWidth(200);
    topLayout->addWidget(m_panoAdi);
    topLayout->addStretch();
    mainLayout->addLayout(topLayout);

    // ===== İÇ SEKMELER (Python ile aynı) =====
    m_innerTabs = new QTabWidget();
    m_innerTabs->setDocumentMode(true);

    // ----- 1. Sekme: Gözle Kontrol -----
    m_gozleKontrol = new GozleKontrolWidget(this);
    m_innerTabs->addTab(m_gozleKontrol, tr("Gözle Kontrol"));

    // ----- 2. Sekme: Fonksiyon Testleri (Parafudr/Empedans dahil) -----
    QWidget* ftWidget = new QWidget();
    QVBoxLayout* ftLayout = new QVBoxLayout(ftWidget);
    ftLayout->setContentsMargins(0, 0, 0, 0);
    ftLayout->setSpacing(4);

    // Parafudr/Empedans bölümü üstte
    QWidget* empedansWidget = setupAnaDagitimPano();
    ftLayout->addWidget(empedansWidget);

    // Fonksiyon testleri tablosu (esnek)
    setupFonksiyonTestleri();
    ftLayout->addWidget(m_fonksiyonTable, 1);

    m_innerTabs->addTab(ftWidget, tr("Fonksiyon Testleri"));

    // ----- 3. Sekme: Termal -----
    QWidget* termalWidget = setupTermalGoruntuler();
    m_innerTabs->addTab(termalWidget, tr("Termal"));

    mainLayout->addWidget(m_innerTabs, 1);
}

QWidget* PanoTabWidget::setupAnaDagitimPano() {
    // Python'daki gibi alt bölüm - GroupBox
    QGroupBox* group = new QGroupBox(this);
    group->setObjectName("anaDagitimGroup");
    group->setFlat(true);
    // Python'daki gibi belirgin yapmak için biraz stil ekleyelim
    group->setStyleSheet("QGroupBox#anaDagitimGroup { border-top: 1px solid #555; margin-top: 5px; padding-top: 5px; }");

    QVBoxLayout* mainLayout = new QVBoxLayout(group);
    mainLayout->setContentsMargins(5, 5, 5, 5);
    mainLayout->setSpacing(3);

    // ===== Satır 1: Parafudr Tipi ve Imax =====
    QHBoxLayout* row1 = new QHBoxLayout();
    row1->addWidget(new QLabel(tr("Parafudr Tipi")));
    m_parafudrTip = new QLineEdit(group);
    m_parafudrTip->setPlaceholderText("Örn: T1+T2");
    m_parafudrTip->setMaximumWidth(120);
    row1->addWidget(m_parafudrTip);

    row1->addWidget(new QLabel(tr("Parafudr Imax (kA)")));
    m_parafudrImax = new QLineEdit(group);
    m_parafudrImax->setPlaceholderText("Örn: 40");
    m_parafudrImax->setMaximumWidth(80);
    row1->addWidget(m_parafudrImax);
    row1->addStretch();
    mainLayout->addLayout(row1);

    // ===== Satır 2: Zx, RE, Zln =====
    QHBoxLayout* row2 = new QHBoxLayout();
    row2->addWidget(new QLabel(tr("Zx (Ω)")));
    m_zx = new QLineEdit(group);
    m_zx->setMaximumWidth(60);
    row2->addWidget(m_zx);

    row2->addWidget(new QLabel(tr("RE (Ω)")));
    m_re = new QLineEdit(group);
    m_re->setMaximumWidth(60);
    row2->addWidget(m_re);

    row2->addWidget(new QLabel(tr("Zln (Ω)")));
    m_zln = new QLineEdit(group);
    m_zln->setMaximumWidth(60);
    row2->addWidget(m_zln);
    row2->addStretch();
    mainLayout->addLayout(row2);

    // ===== Satır 3: F-F, L-N, N-PE, Ik3 =====
    QHBoxLayout* row3 = new QHBoxLayout();
    row3->addWidget(new QLabel(tr("F-F (V)")));
    m_ff = new QLineEdit(group);
    m_ff->setMaximumWidth(60);
    row3->addWidget(m_ff);

    row3->addWidget(new QLabel(tr("L-N (V)")));
    m_ln = new QLineEdit(group);
    m_ln->setMaximumWidth(60);
    row3->addWidget(m_ln);

    row3->addWidget(new QLabel(tr("N-PE (V)")));
    m_npe = new QLineEdit(group);
    m_npe->setMaximumWidth(60);
    row3->addWidget(m_npe);

    row3->addWidget(new QLabel(tr("Ik3")));
    m_ik3Auto = new QLineEdit(group);
    m_ik3Auto->setReadOnly(true);
    m_ik3Auto->setMaximumWidth(70);
    m_ik3Auto->setPlaceholderText("F-F/Zln");
    m_ik3Auto->setStyleSheet("background-color: rgba(100, 100, 100, 0.3);");
    row3->addWidget(m_ik3Auto);
    row3->addStretch();
    mainLayout->addLayout(row3);

    // ===== Satır 4: Sonuç =====
    QHBoxLayout* row4 = new QHBoxLayout();
    QLabel* sonucLabel = new QLabel(tr("📋 Sonuç:"));
    sonucLabel->setStyleSheet("font-weight: bold; color: #4caf50;");
    row4->addWidget(sonucLabel);
    m_uygunluk = new QComboBox(group);
    m_uygunluk->addItems({"Uygun", "Uygun Değil"});
    m_uygunluk->setMinimumWidth(120);
    row4->addWidget(m_uygunluk);
    row4->addStretch();
    mainLayout->addLayout(row4);

    // Kullanılmayan eski alanları oluştur (veri uyumluluğu için)
    m_sebekeTipi = new QComboBox(this); m_sebekeTipi->setVisible(false);
    m_enerjiSaglayan = new QLineEdit(this); m_enerjiSaglayan->setVisible(false);
    m_trafoGucu = new QLineEdit(this); m_trafoGucu->setVisible(false);
    m_sistemGerilimi = new QSpinBox(this); m_sistemGerilimi->setVisible(false);
    m_sistemFrekans = new QSpinBox(this); m_sistemFrekans->setVisible(false);
    m_topraklamaDirenci = new QLineEdit(this); m_topraklamaDirenci->setVisible(false);
    m_sigortaTipiAna = new QComboBox(this); m_sigortaTipiAna->setVisible(false);
    m_nominalAkimAna = new QSpinBox(this); m_nominalAkimAna->setVisible(false);
    m_rcdBilgisi = new QLineEdit(this); m_rcdBilgisi->setVisible(false);
    m_loopPeN = new QLineEdit(this); m_loopPeN->setVisible(false);
    m_loopLN = new QLineEdit(this); m_loopLN->setVisible(false);
    m_ik3 = new QLineEdit(this); m_ik3->setVisible(false);

    // Yeni Eklenen Alanlar (Veri yapısı bütünlüğü için)
    m_rcdAnmaAkimi = new QLineEdit(this); m_rcdAnmaAkimi->setVisible(false);
    m_rcdTestBilgisi = new QLineEdit(this); m_rcdTestBilgisi->setVisible(false);
    m_distCevrimEmpedansi = new QLineEdit(this); m_distCevrimEmpedansi->setVisible(false);
    m_hataAkimi = new QLineEdit(this); m_hataAkimi->setVisible(false);
    m_sistemTopraklamaKesiti = new QLineEdit(this); m_sistemTopraklamaKesiti->setVisible(false);
    m_anaEspotansiyelKesiti = new QLineEdit(this); m_anaEspotansiyelKesiti->setVisible(false);

    return group;
}

void PanoTabWidget::setupFonksiyonTestleri() {
    m_fonksiyonTable = new FonksiyonTestleriTable(this);
    connect(m_fonksiyonTable, &FonksiyonTestleriTable::dataChanged,
            this, &PanoTabWidget::dataChanged);
}

QWidget* PanoTabWidget::setupTermalGoruntuler() {
    QGroupBox* group = new QGroupBox(tr("Termal Görüntüler"), this);
    group->setObjectName("termalGroup");

    QVBoxLayout* mainLayout = new QVBoxLayout(group);

    // Drag & drop area
    m_termalImages = new DragDropWidget(group); // Parenting it to the group
    m_termalImages->setAcceptedExtensions({"docx", "jpg", "jpeg", "png"});
    m_termalImages->setPlaceholderText(tr("Fluke DOCX veya görüntü dosyalarını sürükleyin\nveya çift tıklayarak seçin"));
    m_termalImages->setMinimumHeight(150);

    connect(m_termalImages, &DragDropWidget::filesChanged,
            this, &PanoTabWidget::dataChanged);

    mainLayout->addWidget(m_termalImages);

    // Temizle butonu
    QPushButton* clearBtn = new QPushButton(tr("Temizle"), group);
    clearBtn->setMaximumWidth(100);
    connect(clearBtn, &QPushButton::clicked, m_termalImages, &DragDropWidget::clear);
    mainLayout->addWidget(clearBtn);

    // Alt satır: Zemin izolasyonu (GİZLİ - varsayılan değerler kullanılacak)
    // Görsel Kontrol'da "Zemin Izolasyonu" varsa bu değerler rapora yazılır
    QHBoxLayout* bottomLayout = new QHBoxLayout();

    // Zemin izolasyonu - GİZLİ widget (veri için tutulur ama gösterilmez)
    QGroupBox* zeminGroup = new QGroupBox(tr("Zemin İzolasyonu"), group);
    zeminGroup->setVisible(false);  // GİZLE
    QGridLayout* zeminGrid = new QGridLayout(zeminGroup);

    zeminGrid->addWidget(new QLabel(tr("En (m):")), 0, 0);
    m_zeminEn = new QLineEdit(zeminGroup);
    m_zeminEn->setText("1");  // VARSAYILAN: 1
    m_zeminEn->setMaximumWidth(60);
    zeminGrid->addWidget(m_zeminEn, 0, 1);

    zeminGrid->addWidget(new QLabel(tr("Boy (m):")), 0, 2);
    m_zeminBoy = new QLineEdit(zeminGroup);
    m_zeminBoy->setText("1");  // VARSAYILAN: 1
    m_zeminBoy->setMaximumWidth(60);
    zeminGrid->addWidget(m_zeminBoy, 0, 3);

    zeminGrid->addWidget(new QLabel(tr("İzo. Direnci:")), 1, 0);
    m_izoDirenci = new QLineEdit(zeminGroup);
    m_izoDirenci->setText(">50MΩ");  // VARSAYILAN: >50MΩ
    m_izoDirenci->setMaximumWidth(80);
    zeminGrid->addWidget(m_izoDirenci, 1, 1);

    zeminGrid->addWidget(new QLabel(tr("Uygunluk:")), 1, 2);
    m_izoUygunluk = new QComboBox(zeminGroup);
    m_izoUygunluk->addItems({"Uygun", "Uygun Değil"});
    m_izoUygunluk->setCurrentText("Uygun");  // VARSAYILAN: Uygun
    zeminGrid->addWidget(m_izoUygunluk, 1, 3);

    bottomLayout->addWidget(zeminGroup);
    bottomLayout->addStretch();

    mainLayout->addLayout(bottomLayout);

    return group;
}

void PanoTabWidget::setupConnections() {
    // Ik3 hesaplama: F-F / Zln
    connect(m_ff, &QLineEdit::textChanged, this, &PanoTabWidget::updateIk3);
    connect(m_zln, &QLineEdit::textChanged, this, &PanoTabWidget::updateIk3);

    // Pano adı değiştiğinde tab başlığını güncelle
    connect(m_panoAdi, &QLineEdit::textChanged, this, [this](const QString& text) {
        emit panoNameChanged(m_panoIndex, text);
        emit dataChanged();
    });

    // Data changed signals
    connect(m_sebekeTipi, &QComboBox::currentTextChanged, this, &PanoTabWidget::dataChanged);
    connect(m_zx, &QLineEdit::textChanged, this, &PanoTabWidget::dataChanged);
    connect(m_re, &QLineEdit::textChanged, this, &PanoTabWidget::dataChanged);
    connect(m_parafudrTip, &QLineEdit::textChanged, this, &PanoTabWidget::dataChanged);
    connect(m_uygunluk, &QComboBox::currentTextChanged, this, &PanoTabWidget::dataChanged);

    // OTOMATİK SONUÇ GÜNCELLEMESİ (Python: _check_and_update_sonuc)
    // FT satırı değişince sonucu kontrol et
    connect(m_fonksiyonTable, &FonksiyonTestleriTable::dataChanged,
            this, &PanoTabWidget::checkAndUpdateSonuc);

    // Kritik GK alanı değişince sonucu kontrol et
    connect(m_gozleKontrol, &GozleKontrolWidget::criticalFieldChanged,
            this, &PanoTabWidget::checkAndUpdateSonuc);
}

void PanoTabWidget::updateIk3() {
    // Ik3 = F-F / Zln (Python ile aynı formül)
    QString ffText = m_ff->text().trimmed().replace(',', '.');
    QString zlnText = m_zln->text().trimmed().replace(',', '.');

    bool ffOk, zlnOk;
    double ffVal = ffText.toDouble(&ffOk);
    double zlnVal = zlnText.toDouble(&zlnOk);

    if (ffOk && zlnOk && zlnVal > 0) {
        int ik3 = static_cast<int>(std::round(ffVal / zlnVal));
        m_ik3Auto->setText(QString::number(ik3));
    } else {
        m_ik3Auto->clear();
    }
}

void PanoTabWidget::checkAndUpdateSonuc() {
    // Python: _check_and_update_sonuc
    // Tüm kritik koşulları kontrol et ve sonucu otomatik güncelle
    bool shouldBeUygunDegil = false;

    // 1. Fonksiyon testlerinde "Uygun Değil" var mı?
    QVector<FonksiyonTesti> ftData = m_fonksiyonTable->getData();
    for (const FonksiyonTesti& test : ftData) {
        if (test.sonuc == QString::fromUtf8("Uygun Değil")) {
            shouldBeUygunDegil = true;
            break;
        }
    }

    // 2. Kritik GK alanlarını kontrol et (Zemin İzolasyonu, Pano Kapak, Aşırı Yük)
    if (!shouldBeUygunDegil && m_gozleKontrol->hasUygunDegilCriticalField()) {
        shouldBeUygunDegil = true;
    }

    // Sonucu güncelle - sadece "Uygun Değil" yapılır, geri "Uygun" yapılmaz
    // (Python davranışı ile aynı)
    if (shouldBeUygunDegil) {
        m_uygunluk->setCurrentText(QString::fromUtf8("Uygun Değil"));
    }
}

PanoData PanoTabWidget::getData() const {
    PanoData data;
    data.panoIndex = m_panoIndex;
    data.panoAdi = m_panoAdi->text();

    // Ana Dağıtım Pano - Görünür alanlardan
    data.anaDagitimPano.parafudrTip = m_parafudrTip->text();
    data.anaDagitimPano.parafudrImax = m_parafudrImax->text();
    data.anaDagitimPano.loopPeN = m_zx->text();  // Zx = loopPeN
    data.anaDagitimPano.distCevrimEmpedansi = m_re->text();  // RE = dış çevrim
    data.anaDagitimPano.loopLN = m_zln->text();  // Zln

    // Gerilimler
    int ff = m_ff->text().isEmpty() ? 400 : m_ff->text().toInt();
    data.anaDagitimPano.sistemGerilimi = ff;
    data.anaDagitimPano.ln = m_ln->text();    // L-N (V)
    data.anaDagitimPano.npe = m_npe->text();  // N-PE (V)

    // Sonuç/Uygunluk
    data.genelSonuc = m_uygunluk->currentText();

    // Gizli alanlardan (CommonInfoPanel'den doldurulacak)
    data.anaDagitimPano.sebekeTipi = m_sebekeTipi->currentText();
    data.anaDagitimPano.enerjiSaglayan = m_enerjiSaglayan->text();
    data.anaDagitimPano.trafoGucu = m_trafoGucu->text();
    data.anaDagitimPano.sistemFrekans = m_sistemFrekans->value();
    data.anaDagitimPano.topraklamaDirenci = m_topraklamaDirenci->text();
    data.anaDagitimPano.sigortaTipiAna = m_sigortaTipiAna->currentText();
    data.anaDagitimPano.nominalAkimAna = m_nominalAkimAna->value();
    data.anaDagitimPano.rcdBilgisi = m_rcdBilgisi->text();
    data.anaDagitimPano.ik3 = m_ik3->text();
    data.anaDagitimPano.rcdAnmaAkimi = m_rcdAnmaAkimi->text();
    data.anaDagitimPano.rcdTestBilgisi = m_rcdTestBilgisi->text();
    data.anaDagitimPano.hataAkimi = m_hataAkimi->text();
    data.anaDagitimPano.sistemTopraklamaKesiti = m_sistemTopraklamaKesiti->text();
    data.anaDagitimPano.anaEspotansiyelKesiti = m_anaEspotansiyelKesiti->text();

    // Fonksiyon testleri
    data.fonksiyonTestleri = m_fonksiyonTable->getData();

    // Termal görüntüler
    for (const QString& path : m_termalImages->getFiles()) {
        TermalGoruntu img;
        img.imagePath = path;
        img.tip = path.endsWith(".docx") ? "fluke" : "image";
        data.termalGoruntuler.append(img);
    }

    // Gözle Kontrol - Widget'tan al ve GozleKontrolMaddesi formatına çevir
    QMap<QString, QString> gkData = m_gozleKontrol->getData();
    static const QStringList gkOrder = {
        "Kablo Sebeke Tarafi", "Pano Sabitlenmesi",
        "Elektrik Panosu Etrafinda Yabanci Malzemeler", "Kablo Donanim Tarafi",
        "Dis Darbelere Karsi Koruma Onlemi", "Zemin Izolasyonu",
        "Topraklama Iletkeni", "Ek Potansiyel Dengeleme Iletkeni",
        "Ana Potansiyel Dengeleme Iletkeni", "Pano Kapak Baglantisi Kontrolu 6 mm2",
        "Elektriksel Olmayan Tesislere Yaklasma", "Guvenlik Devre Ayrilmasi",
        "Bant Ayrilmasi", "Pano Ic Kapak",
        "Semalar Talimatlar", "Tehlike Isaretleri",
        "Koruma Cihaz ve Terminal Etiket", "Kablo Yollari",
        "Tesisat Yontemi", "Kablo Renk Kodlari",
        "Yangin Engeli", "Kontak Gevsekligi Isinmasi",
        "Asiri Yuk Isinmasi", "", "", // GK_24, GK_25 = Fotoğraf tarihi/no (Fluke'dan)
        "Yangin Sondurme", "Korozyon Kontrolu",
        "Ekipman Temizlik", "Acil Durum Aydinlatma"
    };

    for (int i = 0; i < gkOrder.size() && i < 29; ++i) {
        GozleKontrolMaddesi madde;
        madde.maddeNo = i + 1;  // 1-indexed
        madde.maddeAdi = gkOrder[i];
        if (!gkOrder[i].isEmpty() && gkData.contains(gkOrder[i])) {
            madde.sonuc = gkData[gkOrder[i]];
        } else {
            madde.sonuc = "Uygun";  // Varsayılan
        }
        data.gozleKontrol.append(madde);
    }

    // Zemin izolasyonu
    data.zeminEn = m_zeminEn->text();
    data.zeminBoy = m_zeminBoy->text();
    data.izoDirenci = m_izoDirenci->text();
    data.izoUygunluk = m_izoUygunluk->currentText();

    return data;
}

void PanoTabWidget::setData(const PanoData& data) {
    m_panoAdi->setText(data.panoAdi);

    // Ana Dağıtım Pano
    int idx = m_sebekeTipi->findText(data.anaDagitimPano.sebekeTipi);
    if (idx >= 0) m_sebekeTipi->setCurrentIndex(idx);

    m_enerjiSaglayan->setText(data.anaDagitimPano.enerjiSaglayan);
    m_trafoGucu->setText(data.anaDagitimPano.trafoGucu);
    m_sistemGerilimi->setValue(data.anaDagitimPano.sistemGerilimi);
    m_sistemFrekans->setValue(data.anaDagitimPano.sistemFrekans);
    m_topraklamaDirenci->setText(data.anaDagitimPano.topraklamaDirenci);

    idx = m_sigortaTipiAna->findText(data.anaDagitimPano.sigortaTipiAna);
    if (idx >= 0) m_sigortaTipiAna->setCurrentIndex(idx);

    m_nominalAkimAna->setValue(data.anaDagitimPano.nominalAkimAna);
    m_rcdBilgisi->setText(data.anaDagitimPano.rcdBilgisi);
    m_loopPeN->setText(data.anaDagitimPano.loopPeN);
    m_loopLN->setText(data.anaDagitimPano.loopLN);

    // Yeni Alanlar
    m_rcdAnmaAkimi->setText(data.anaDagitimPano.rcdAnmaAkimi);
    m_rcdTestBilgisi->setText(data.anaDagitimPano.rcdTestBilgisi);
    m_distCevrimEmpedansi->setText(data.anaDagitimPano.distCevrimEmpedansi);
    m_hataAkimi->setText(data.anaDagitimPano.hataAkimi);
    m_sistemTopraklamaKesiti->setText(data.anaDagitimPano.sistemTopraklamaKesiti);
    m_anaEspotansiyelKesiti->setText(data.anaDagitimPano.anaEspotansiyelKesiti);

    // GÖRÜNÜR ALANLAR (yükleme sırasında eksik kalan)
    m_parafudrTip->setText(data.anaDagitimPano.parafudrTip);
    m_parafudrImax->setText(data.anaDagitimPano.parafudrImax);
    m_zx->setText(data.anaDagitimPano.loopPeN);       // Zx = loopPeN
    m_re->setText(data.anaDagitimPano.distCevrimEmpedansi);  // RE = distCevrimEmpedansi
    m_zln->setText(data.anaDagitimPano.loopLN);       // Zln = loopLN
    m_ff->setText(QString::number(data.anaDagitimPano.sistemGerilimi));  // F-F

    // Genel Sonuç/Uygunluk
    idx = m_uygunluk->findText(data.genelSonuc);
    if (idx >= 0) m_uygunluk->setCurrentIndex(idx);

    // Gözle Kontrol maddeleri
    QMap<QString, QString> gkData;
    for (const auto& gk : data.gozleKontrol) {
        gkData[gk.maddeAdi] = gk.sonuc;
    }
    m_gozleKontrol->setData(gkData);

    // Fonksiyon testleri
    m_fonksiyonTable->setData(data.fonksiyonTestleri);

    // L-N ve N-PE alanları
    m_ln->setText(data.anaDagitimPano.ln);
    m_npe->setText(data.anaDagitimPano.npe);

    // Termal görüntüler
    QStringList paths;
    for (const auto& img : data.termalGoruntuler) {
        paths.append(img.imagePath);
    }
    m_termalImages->setFiles(paths);

    // Zemin izolasyonu
    m_zeminEn->setText(data.zeminEn);
    m_zeminBoy->setText(data.zeminBoy);
    m_izoDirenci->setText(data.izoDirenci);

    idx = m_izoUygunluk->findText(data.izoUygunluk);
    if (idx >= 0) m_izoUygunluk->setCurrentIndex(idx);
}

void PanoTabWidget::clear() {
    m_panoAdi->clear();
    m_sebekeTipi->setCurrentIndex(0);
    m_enerjiSaglayan->setText("TEİAŞ Genel Müdürlüğü");
    m_trafoGucu->clear();
    m_sistemGerilimi->setValue(400);
    m_sistemFrekans->setValue(50);
    m_topraklamaDirenci->clear();
    m_sigortaTipiAna->setCurrentIndex(0);
    m_nominalAkimAna->setValue(250);
    m_rcdBilgisi->clear();
    m_loopPeN->clear();
    m_loopLN->clear();
    // Yeni alanlar
    m_rcdAnmaAkimi->clear();
    m_rcdTestBilgisi->clear();
    m_distCevrimEmpedansi->clear();
    m_hataAkimi->clear();
    m_sistemTopraklamaKesiti->clear();
    m_anaEspotansiyelKesiti->clear();

    m_fonksiyonTable->clear();
    m_termalImages->clear();
    m_zeminEn->clear();
    m_zeminBoy->clear();
    m_izoDirenci->clear();
    m_izoUygunluk->setCurrentIndex(0);
}

} // namespace RaporSistemi

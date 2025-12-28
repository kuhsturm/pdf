/**
 * MainWindow.cpp
 *
 * Ana pencere implementasyonu.
 */

#include "MainWindow.h"
#include "PanoTabWidget.h"
#include "core/PdfParser.h"
#include "core/KisiBilgileriReader.h"
#include "logic/ReportGenerator.h"

#include <QMenuBar>
#include <QToolBar>
#include <QStatusBar>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QSplitter>
#include <QScrollArea>
#include <QMessageBox>
#include <QFileDialog>
#include <QShortcut>
#include <QCloseEvent>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QFile>
#include <QFileInfo>
#include <QApplication>
#include <QDir>
#include <QStyle>
#include <QFrame>
#include <QTime>
#include <QToolButton>
#include <QDesktopServices> // Klasör açmak için
#include <QUrl>
#include <QStandardPaths> // Masaüstü yolu için

namespace RaporSistemi {

// ==================== MainWindow ====================

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
{
    setupUi();
    setupMenuBar();
    setupToolBar();
    setupStatusBar();
    setupShortcuts();

    // İlk panoyu ekle
    addNewPano();

    updateWindowTitle();
}

MainWindow::~MainWindow() = default;

void MainWindow::setupUi() {
    setMinimumSize(1400, 900);

    // Ana widget
    QWidget* centralWidget = new QWidget(this);
    setCentralWidget(centralWidget);

    // ===== PYTHON GİBİ YATAY LAYOUT: Sol panel + Sağ panel =====
    QHBoxLayout* mainLayout = new QHBoxLayout(centralWidget);
    mainLayout->setContentsMargins(5, 5, 5, 5);
    mainLayout->setSpacing(5);

    // ===== SOL PANEL (Sidebar) - Python'daki gibi 350px sabit =====
    m_commonInfoPanel = new CommonInfoPanel();
    m_commonInfoPanel->setFixedWidth(350);
    m_commonInfoPanel->setMinimumHeight(700);
    connect(m_commonInfoPanel, &CommonInfoPanel::dataChanged, this, &MainWindow::onProjectModified);
    connect(m_commonInfoPanel, &CommonInfoPanel::contractLoaded, this, &MainWindow::onContractLoad);
    connect(m_commonInfoPanel, &CommonInfoPanel::generateReportsRequested, this, &MainWindow::generateReports);
    mainLayout->addWidget(m_commonInfoPanel);

    // ===== SAĞ PANEL: Pano Tab'ları =====
    QVBoxLayout* rightLayout = new QVBoxLayout();
    rightLayout->setSpacing(4);

    // Pano sekmeleri
    m_panoTabs = new QTabWidget();
    m_panoTabs->setTabsClosable(true);
    m_panoTabs->setMovable(true);
    m_panoTabs->setDocumentMode(true);

    // Pano Ekle butonu (+)
    QToolButton* addBtn = new QToolButton();
    addBtn->setText("+");
    addBtn->setToolTip(tr("Yeni Pano Ekle"));
    connect(addBtn, &QToolButton::clicked, this, &MainWindow::addNewPano);
    m_panoTabs->setCornerWidget(addBtn, Qt::TopRightCorner);

    connect(m_panoTabs, &QTabWidget::tabCloseRequested, this, &MainWindow::onTabCloseRequested);
    connect(m_panoTabs, &QTabWidget::currentChanged, this, &MainWindow::onTabChanged);

    rightLayout->addWidget(m_panoTabs, 1);
    mainLayout->addLayout(rightLayout, 1);
}

void MainWindow::setupMenuBar() {
    QMenuBar* menuBar = this->menuBar();

    // Dosya menüsü
    QMenu* fileMenu = menuBar->addMenu(tr("&Dosya"));

    QAction* newAction = fileMenu->addAction(tr("&Yeni Proje"));
    newAction->setShortcut(QKeySequence::New);
    connect(newAction, &QAction::triggered, this, &MainWindow::newProject);

    QAction* openAction = fileMenu->addAction(tr("&Aç..."));
    openAction->setShortcut(QKeySequence::Open);
    connect(openAction, &QAction::triggered, this, &MainWindow::openProject);

    QAction* saveAction = fileMenu->addAction(tr("&Kaydet"));
    saveAction->setShortcut(QKeySequence::Save);
    connect(saveAction, &QAction::triggered, this, &MainWindow::saveProject);

    QAction* saveAsAction = fileMenu->addAction(tr("Farklı &Kaydet..."));
    saveAsAction->setShortcut(QKeySequence::SaveAs);
    connect(saveAsAction, &QAction::triggered, this, &MainWindow::saveProjectAs);

    fileMenu->addSeparator();

    QAction* loadContractAction = fileMenu->addAction(tr("Sözleşme &Yükle..."));
    connect(loadContractAction, &QAction::triggered, this, &MainWindow::loadContract);

    fileMenu->addSeparator();

    QAction* exitAction = fileMenu->addAction(tr("Çı&kış"));
    exitAction->setShortcut(QKeySequence::Quit);
    connect(exitAction, &QAction::triggered, this, &QWidget::close);

    // Pano menüsü
    QMenu* panoMenu = menuBar->addMenu(tr("&Pano"));

    QAction* addPanoAction = panoMenu->addAction(tr("Yeni Pano &Ekle"));
    addPanoAction->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_T));
    connect(addPanoAction, &QAction::triggered, this, &MainWindow::addNewPano);

    // Rapor menüsü
    QMenu* reportMenu = menuBar->addMenu(tr("&Rapor"));

    QAction* generateAction = reportMenu->addAction(tr("Tüm Raporları &Oluştur"));
    generateAction->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_G));
    connect(generateAction, &QAction::triggered, this, &MainWindow::generateReports);

    // Yardım menüsü
    QMenu* helpMenu = menuBar->addMenu(tr("&Yardım"));

    QAction* updatePersonAction = helpMenu->addAction(tr("Kişi Listesini &Güncelle..."));
    connect(updatePersonAction, &QAction::triggered, this, &MainWindow::updatePersonFile);

    helpMenu->addSeparator();

    QAction* aboutAction = helpMenu->addAction(tr("&Hakkında"));
    connect(aboutAction, &QAction::triggered, this, &MainWindow::showAboutDialog);
}

void MainWindow::setupToolBar() {
    QToolBar* toolBar = addToolBar(tr("Ana Araç Çubuğu"));
    toolBar->setMovable(false);
    toolBar->setIconSize(QSize(24, 24));

    // Toolbar butonları
    QAction* newAction = toolBar->addAction(
        style()->standardIcon(QStyle::SP_FileIcon), tr("Yeni"));
    connect(newAction, &QAction::triggered, this, &MainWindow::newProject);

    QAction* openAction = toolBar->addAction(
        style()->standardIcon(QStyle::SP_DialogOpenButton), tr("Aç"));
    connect(openAction, &QAction::triggered, this, &MainWindow::openProject);

    QAction* saveAction = toolBar->addAction(
        style()->standardIcon(QStyle::SP_DialogSaveButton), tr("Kaydet"));
    connect(saveAction, &QAction::triggered, this, &MainWindow::saveProject);

    toolBar->addSeparator();

    QAction* addPanoAction = toolBar->addAction(
        style()->standardIcon(QStyle::SP_FileDialogNewFolder), tr("Pano Ekle"));
    connect(addPanoAction, &QAction::triggered, this, &MainWindow::addNewPano);

    // Pano Kopyala butonu
    QAction* copyPanoAction = toolBar->addAction(
        style()->standardIcon(QStyle::SP_FileLinkIcon), tr("Pano Kopyala"));
    connect(copyPanoAction, &QAction::triggered, this, [this]() {
        duplicatePano(m_panoTabs->currentIndex());
    });

    toolBar->addSeparator();

    QAction* generateAction = toolBar->addAction(
        style()->standardIcon(QStyle::SP_MediaPlay), tr("Rapor Oluştur"));
    connect(generateAction, &QAction::triggered, this, &MainWindow::generateReports);
}

void MainWindow::setupStatusBar() {
    m_statusLabel = new QLabel(tr("Hazır"));
    statusBar()->addWidget(m_statusLabel);

    m_progressBar = new QProgressBar();
    m_progressBar->setMaximumWidth(200);
    m_progressBar->setVisible(false);
    statusBar()->addPermanentWidget(m_progressBar);
}

void MainWindow::setupShortcuts() {
    // Ek kısayollar burada tanımlanabilir
}

void MainWindow::addNewPano() {
    ++m_panoCounter;

    PanoTabWidget* panoTab = new PanoTabWidget(m_panoCounter, this);
    connect(panoTab, &PanoTabWidget::dataChanged, this, &MainWindow::onProjectModified);

    // Pano adı değiştiğinde tab başlığını güncelle
    connect(panoTab, &PanoTabWidget::panoNameChanged, this, [this, panoTab](int, const QString& newName) {
        int tabIndex = m_panoTabs->indexOf(panoTab);
        if (tabIndex >= 0) {
            QString title = newName.isEmpty() ? QString("Pano %1").arg(panoTab->panoIndex()) : newName;
            m_panoTabs->setTabText(tabIndex, title);
        }
    });

    QString tabName = QString("Pano %1").arg(m_panoCounter);
    int index = m_panoTabs->addTab(panoTab, tabName);
    m_panoTabs->setCurrentIndex(index);

    onProjectModified();
}

void MainWindow::removePano(int index) {
    if (m_panoTabs->count() <= 1) {
        QMessageBox::warning(this, tr("Uyarı"),
            tr("En az bir pano olmalıdır."));
        return;
    }

    QWidget* widget = m_panoTabs->widget(index);
    m_panoTabs->removeTab(index);
    widget->deleteLater();

    onProjectModified();
}

void MainWindow::duplicatePano(int index) {
    if (index < 0 || index >= m_panoTabs->count()) return;

    // Kaynak panoyu al
    PanoTabWidget* sourceTab = qobject_cast<PanoTabWidget*>(m_panoTabs->widget(index));
    if (!sourceTab) return;

    // Veriyi kopyala
    PanoData sourceData = sourceTab->getData();

    // Yeni pano oluştur
    ++m_panoCounter;
    PanoTabWidget* newTab = new PanoTabWidget(m_panoCounter, this);
    connect(newTab, &PanoTabWidget::dataChanged, this, &MainWindow::onProjectModified);

    // Pano adı değiştiğinde tab başlığını güncelle
    connect(newTab, &PanoTabWidget::panoNameChanged, this, [this, newTab](int, const QString& newName) {
        int tabIndex = m_panoTabs->indexOf(newTab);
        if (tabIndex >= 0) {
            QString title = newName.isEmpty() ? QString("Pano %1").arg(newTab->panoIndex()) : newName;
            m_panoTabs->setTabText(tabIndex, title);
        }
    });

    // Veriyi ayarla
    newTab->setData(sourceData);

    // Tab ekle - kopyalanan pano adını kullan
    QString tabName = sourceData.panoAdi.isEmpty() ?
        QString("Pano %1 (Kopya)").arg(m_panoCounter) : sourceData.panoAdi;
    int newIndex = m_panoTabs->addTab(newTab, tabName);
    m_panoTabs->setCurrentIndex(newIndex);

    onProjectModified();
    m_statusLabel->setText(tr("Pano kopyalandı."));
}

void MainWindow::onTabCloseRequested(int index) {
    if (m_panoTabs->count() <= 1) {
        QMessageBox::warning(this, tr("Uyarı"),
            tr("En az bir pano olmalıdır."));
        return;
    }

    int result = QMessageBox::question(this, tr("Pano Sil"),
        tr("Bu panoyu silmek istediğinizden emin misiniz?"),
        QMessageBox::Yes | QMessageBox::No);

    if (result == QMessageBox::Yes) {
        removePano(index);
    }
}

void MainWindow::onTabChanged(int index) {
    if (index >= 0) {
        m_statusLabel->setText(QString("Aktif: %1").arg(m_panoTabs->tabText(index)));
    }
}

void MainWindow::onProjectModified() {
    m_hasUnsavedChanges = true;
    updateWindowTitle();
}

void MainWindow::updateWindowTitle() {
    QString title = "Elektrik Rapor Sistemi";

    if (!m_currentProjectPath.isEmpty()) {
        QFileInfo info(m_currentProjectPath);
        title += " - " + info.fileName();
    }

    if (m_hasUnsavedChanges) {
        title += " *";
    }

    setWindowTitle(title);
}

void MainWindow::newProject() {
    if (m_hasUnsavedChanges) {
        int result = QMessageBox::question(this, tr("Kaydet?"),
            tr("Değişiklikler kaydedilmedi. Kaydetmek ister misiniz?"),
            QMessageBox::Yes | QMessageBox::No | QMessageBox::Cancel);

        if (result == QMessageBox::Yes) {
            saveProject();
        } else if (result == QMessageBox::Cancel) {
            return;
        }
    }

    clearAll();
    m_currentProjectPath.clear();
    m_hasUnsavedChanges = false;
    addNewPano();
    updateWindowTitle();
}

void MainWindow::clearAll() {
    m_commonInfoPanel->clear();

    while (m_panoTabs->count() > 0) {
        QWidget* widget = m_panoTabs->widget(0);
        m_panoTabs->removeTab(0);
        widget->deleteLater();
    }

    m_panoCounter = 0;
}

void MainWindow::onContractLoad(const QString& pdfPath) {
    // PDF dosyası seç (eğer yol boşsa)
    QString path = pdfPath;
    if (path.isEmpty()) {
        path = QFileDialog::getOpenFileName(this,
            tr("Sözleşme PDF Seç"), QString(),
            tr("PDF Dosyaları (*.pdf);;Tüm Dosyalar (*)"));
    }

    if (path.isEmpty()) return;

    m_statusLabel->setText(tr("Sözleşme yükleniyor..."));
    QApplication::processEvents();

    // PDF'i parse et
    PdfParser parser;
    SozlesmeData sozlesme = parser.parseSozlesme(path);

    if (sozlesme.sozlesmeId.isEmpty() && sozlesme.firmaAdi.isEmpty()) {
        QMessageBox::warning(this, tr("Uyarı"),
            tr("Sözleşme bilgileri okunamadı. PDF formatı desteklenmiyor olabilir."));
        m_statusLabel->setText(tr("Hazır"));
        return;
    }

    // Mevcut firma bilgilerini al
    FirmaBilgileri firma = m_commonInfoPanel->getFirmaBilgileri();

    // ===== Python field mapping (load_sozlesme_pdf ile aynı) =====
    // firma_unvan -> Firma Adi -> firmaAdi
    // firma_adres -> Tesis Adresi -> kontrolAdresi
    // firma_sgk_no -> SGK Sicil No -> sgkSicil
    // sozlesme_id -> Sozlesme ID -> sozlesmeId
    // kontrol_eden_adsoyad -> Kontrol Eden -> kontrolEdenAdSoyad
    // pk_no -> Belge No -> pkNo (NOT: kontrolEdenTc değil!)

    firma.firmaAdi = sozlesme.firmaAdi.trimmed();
    if (firma.firmaAdi.endsWith(':')) firma.firmaAdi.chop(1);

    // Adres ve İl birleştir
    firma.kontrolAdresi = sozlesme.firmaAdres;
    if (!sozlesme.firmaIl.isEmpty()) {
        if (!firma.kontrolAdresi.isEmpty()) firma.kontrolAdresi += " / ";
        firma.kontrolAdresi += sozlesme.firmaIl;
    }
    firma.kontrolAdresi = firma.kontrolAdresi.trimmed();
    if (firma.kontrolAdresi.endsWith(':')) firma.kontrolAdresi.chop(1);
    firma.sgkSicil = sozlesme.sgkSicil;
    firma.sozlesmeId = sozlesme.sozlesmeId;
    firma.kontrolEdenAdSoyad = sozlesme.kontrolEdenAdSoyad;
    firma.pkNo = sozlesme.pkNo;  // pk_no -> Belge No (Python'daki gibi)
    // kontrolEdenTc PDF'den gelmiyor, kisi_bilgileri'nden gelecek
    firma.tesisatSayisi = sozlesme.tesisatSayisi;

    // Tarihler (Eğer PDF'den geldiyse)
    if (!sozlesme.sozlesmeBaslangic.isEmpty()) {
        QDate baslangic = QDate::fromString(sozlesme.sozlesmeBaslangic, "dd.MM.yyyy");
        if (baslangic.isValid()) {
            firma.raporTarihi = baslangic;
            firma.baslangicTarihSaat = QDateTime(baslangic, QTime(8, 30));
        }
    }
    if (!sozlesme.sozlesmeBitis.isEmpty()) {
        QDate bitis = QDate::fromString(sozlesme.sozlesmeBitis, "dd.MM.yyyy");
        if (bitis.isValid()) {
            firma.bitisTarihSaat = QDateTime(bitis, QTime(17, 30));
        }
    }

    // Tarihler (Eğer PDF'den geldiyse)
    if (!sozlesme.sozlesmeBaslangic.isEmpty()) {
        QDate baslangic = QDate::fromString(sozlesme.sozlesmeBaslangic, "dd.MM.yyyy");
        if (baslangic.isValid()) {
            firma.raporTarihi = baslangic;
            firma.baslangicTarihSaat = QDateTime(baslangic, QTime(8, 30));
        }
    }
    if (!sozlesme.sozlesmeBitis.isEmpty()) {
        QDate bitis = QDate::fromString(sozlesme.sozlesmeBitis, "dd.MM.yyyy");
        if (bitis.isValid()) {
            firma.bitisTarihSaat = QDateTime(bitis, QTime(17, 30));
        }
    }

    // --- Kişi Bilgileri Dosya Yönetimi (Gömülü + Güncellenebilir) ---
    QString dataLocation = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir().mkpath(dataLocation);
    QString managedPath = dataLocation + "/kisi_bilgileri.xlsx";

    // Eğer AppData'da yoksa
    if (!QFile::exists(managedPath)) {
        // 1. Önce embedded (resource) var mı bak (resources.qrc)
        if (QFile::exists(":/resources/kisi_bilgileri.xlsx")) {
             QFile::copy(":/resources/kisi_bilgileri.xlsx", managedPath);
             QFile::setPermissions(managedPath, QFile::ReadOwner | QFile::WriteOwner);
        }
        // 2. Yoksa exe yanı
        else if (QFile::exists("kisi_bilgileri.xlsx")) {
             QFile::copy("kisi_bilgileri.xlsx", managedPath);
        }
    }

    // Kullanılacak path
    QString kisiBilgileriPath = managedPath;
    if (!QFile::exists(kisiBilgileriPath) && QFile::exists("kisi_bilgileri.xlsx")) {
         kisiBilgileriPath = "kisi_bilgileri.xlsx"; // Fallback to local
    }

    if (QFile::exists(kisiBilgileriPath)) {
        KisiBilgileriReader kisiReader;
        if (kisiReader.load(kisiBilgileriPath)) {
            // Check if person exists explicitly for feedback
            KisiBilgisi kb = kisiReader.getPersonByName(firma.kontrolEdenAdSoyad);
            if (!kb.isValid()) {
                QMessageBox::warning(this, tr("Uyarı"),
                    tr("'%1' kişisi '%2' dosyasında bulunamadı.\nCihaz bilgileri yüklenemedi.\n\nYol: %3")
                    .arg(firma.kontrolEdenAdSoyad)
                    .arg(QFileInfo(kisiBilgileriPath).fileName())
                    .arg(kisiBilgileriPath));
            } else {
                kisiReader.fillCihazBilgileri(firma.kontrolEdenAdSoyad, firma);
            }
        }
    } else {
         QMessageBox::warning(this, tr("Uyarı"),
             tr("Cihaz bilgileri dosyası bulunamadı!\n"
                "Program içine gömülü dosya oluşturulamadı.\nLütfen 'kisi_bilgileri.xlsx' dosyasını exe yanına koyun."));
    }

    // Paneli güncelle
    m_commonInfoPanel->setFirmaBilgileri(firma);

    // Tesisat sayısına göre panolar ekle
    if (sozlesme.tesisatSayisi > 1) {
        while (m_panoTabs->count() < sozlesme.tesisatSayisi) {
            addNewPano();
        }
    }

    m_statusLabel->setText(tr("Sözleşme yüklendi: %1").arg(sozlesme.firmaAdi));
    onProjectModified();
}

void MainWindow::updatePersonFile() {
    QString currentPath = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation) + "/kisi_bilgileri.xlsx";

    QString newPath = QFileDialog::getOpenFileName(this,
        tr("Yeni Kişi Listesi Seç (Excel)"), QString(),
        tr("Excel Dosyaları (*.xlsx)"));

    if (newPath.isEmpty()) return;

    // Validate file
    KisiBilgileriReader reader;
    if (!reader.load(newPath)) {
         QMessageBox::critical(this, tr("Hata"), tr("Seçilen dosya geçerli bir Excel dosyası değil veya okunamadı."));
         return;
    }

    // Copy/Overwrite
    QDir().mkpath(QFileInfo(currentPath).absolutePath());
    if (QFile::exists(currentPath)) {
        QFile::remove(currentPath);
    }

    if (QFile::copy(newPath, currentPath)) {
         QFile::setPermissions(currentPath, QFile::ReadOwner | QFile::WriteOwner);
         QMessageBox::information(this, tr("Başarılı"), tr("Kişi listesi başarıyla güncellendi!\nBundan sonraki işlemlerde bu liste kullanılacak."));
    } else {
         QMessageBox::critical(this, tr("Hata"), tr("Dosya kopyalanırken hata oluştu."));
    }
}

void MainWindow::openProject() {
    if (m_hasUnsavedChanges) {
        int result = QMessageBox::question(this, tr("Kaydet?"),
            tr("Mevcut değişiklikler kaydedilmedi. Kaydetmek ister misiniz?"),
            QMessageBox::Yes | QMessageBox::No | QMessageBox::Cancel);

        if (result == QMessageBox::Yes) {
            saveProject();
        } else if (result == QMessageBox::Cancel) {
            return;
        }
    }

    QString path = QFileDialog::getOpenFileName(this,
        tr("Proje Aç"), QString(),
        tr("Rapor Projesi (*.erp);;Tüm Dosyalar (*)"));

    if (path.isEmpty()) return;

    QFile file(path);
    if (!file.open(QIODevice::ReadOnly)) {
        QMessageBox::critical(this, tr("Hata"),
            tr("Dosya açılamadı: %1").arg(path));
        return;
    }

    QJsonDocument doc = QJsonDocument::fromJson(file.readAll());
    file.close();

    if (doc.isNull() || !doc.isObject()) {
        QMessageBox::critical(this, tr("Hata"), tr("Geçersiz proje dosyası."));
        return;
    }

    QJsonObject root = doc.object();
    Proje proje;

    // FirmaBilgileri
    QJsonObject firma = root["firmaBilgileri"].toObject();
    proje.firmaBilgileri.firmaAdi = firma["firmaAdi"].toString();
    proje.firmaBilgileri.kontrolAdresi = firma["kontrolAdresi"].toString();
    proje.firmaBilgileri.sgkSicil = firma["sgkSicil"].toString();
    proje.firmaBilgileri.sozlesmeId = firma["sozlesmeId"].toString();
    proje.firmaBilgileri.raporNumarasi = firma["raporNumarasi"].toString();
    proje.firmaBilgileri.raporTarihi = QDate::fromString(firma["raporTarihi"].toString(), "dd.MM.yyyy");
    proje.firmaBilgileri.baslangicTarihSaat = QDateTime::fromString(firma["baslangicTarihSaat"].toString(), "dd.MM.yyyy HH:mm");
    proje.firmaBilgileri.bitisTarihSaat = QDateTime::fromString(firma["bitisTarihSaat"].toString(), "dd.MM.yyyy HH:mm");
    proje.firmaBilgileri.kontrolEdenAdSoyad = firma["kontrolEdenAdSoyad"].toString();
    proje.firmaBilgileri.kontrolEdenTc = firma["kontrolEdenTc"].toString();
    proje.firmaBilgileri.pkNo = firma["pkNo"].toString();
    proje.firmaBilgileri.teklifNumarasi = firma["teklifNumarasi"].toString();
    proje.firmaBilgileri.tesisatSayisi = firma["tesisatSayisi"].toInt();
    // Cihaz bilgileri - TAM
    proje.firmaBilgileri.termalCihazAdi = firma["termalCihazAdi"].toString();
    proje.firmaBilgileri.termalKalibrasyonTarihi = firma["termalKalibrasyonTarihi"].toString();
    proje.firmaBilgileri.termalKalibrasyonGecerlilik = firma["termalKalibrasyonGecerlilik"].toString();
    proje.firmaBilgileri.termalSeriNo = firma["termalSeriNo"].toString();
    proje.firmaBilgileri.termalKalibrasyonNo = firma["termalKalibrasyonNo"].toString();

    proje.firmaBilgileri.olcumCihazAdi = firma["olcumCihazAdi"].toString();
    proje.firmaBilgileri.olcumKalibrasyonTarihi = firma["olcumKalibrasyonTarihi"].toString();
    proje.firmaBilgileri.olcumKalibrasyonGecerlilik = firma["olcumKalibrasyonGecerlilik"].toString();
    proje.firmaBilgileri.olcumSeriNo = firma["olcumSeriNo"].toString();
    proje.firmaBilgileri.olcumKalibrasyonNo = firma["olcumKalibrasyonNo"].toString();

    // Eski format uyumluluk
    proje.firmaBilgileri.cihaz1Adi = firma["cihaz1Adi"].toString();
    proje.firmaBilgileri.cihaz1SeriNo = firma["cihaz1SeriNo"].toString();
    proje.firmaBilgileri.cihaz2Adi = firma["cihaz2Adi"].toString();
    proje.firmaBilgileri.cihaz2SeriNo = firma["cihaz2SeriNo"].toString();

    // AnaPanoBilgileri
    QJsonObject anaPano = root["anaPanoBilgileri"].toObject();
    proje.anaPanoBilgileri.sebekeTipi = anaPano["sebekeTipi"].toString();
    proje.anaPanoBilgileri.topraklamaDirenci = anaPano["topraklamaDirenci"].toString();
    proje.anaPanoBilgileri.sigortaTipiAna = anaPano["sigortaTipiAna"].toString();
    proje.anaPanoBilgileri.nominalAkimAna = anaPano["nominalAkimAna"].toInt();
    proje.anaPanoBilgileri.rcdBilgisi = anaPano["rcdBilgisi"].toString();
    proje.anaPanoBilgileri.rcdAnmaAkimi = anaPano["rcdAnmaAkimi"].toString();

    // Panolar
    QJsonArray panolarArray = root["panolar"].toArray();
    for (const QJsonValue& panoVal : panolarArray) {
        QJsonObject panoObj = panoVal.toObject();
        PanoData pano;
        pano.panoAdi = panoObj["panoAdi"].toString();
        pano.raporNumarasi = panoObj["raporNumarasi"].toString();
        // Not: PanoData'da 'sonuc' alanı yok, potansiyelSonuc kullan
        pano.potansiyelSonuc = panoObj["potansiyelSonuc"].toString();

        // AnaDagitimPano
        QJsonObject adp = panoObj["anaDagitimPano"].toObject();
        pano.anaDagitimPano.sebekeTipi = adp["sebekeTipi"].toString();
        pano.anaDagitimPano.topraklamaDirenci = adp["topraklamaDirenci"].toString();
        pano.anaDagitimPano.sigortaTipiAna = adp["sigortaTipiAna"].toString();
        pano.anaDagitimPano.nominalAkimAna = adp["nominalAkimAna"].toInt();
        pano.anaDagitimPano.loopPeN = adp["loopPeN"].toString();
        pano.anaDagitimPano.loopLN = adp["loopLN"].toString();
        pano.anaDagitimPano.ik3 = adp["ik3"].toString();
        pano.anaDagitimPano.distCevrimEmpedansi = adp["distCevrimEmpedansi"].toString();
        pano.anaDagitimPano.rcdAnmaAkimi = adp["rcdAnmaAkimi"].toString();
        // EKSİK ALANLAR
        pano.anaDagitimPano.parafudrTip = adp["parafudrTip"].toString();
        pano.anaDagitimPano.parafudrImax = adp["parafudrImax"].toString();
        pano.anaDagitimPano.sistemGerilimi = adp["sistemGerilimi"].toInt();
        pano.anaDagitimPano.enerjiSaglayan = adp["enerjiSaglayan"].toString();
        pano.anaDagitimPano.trafoGucu = adp["trafoGucu"].toString();
        pano.anaDagitimPano.sistemFrekans = adp["sistemFrekans"].toInt();
        pano.anaDagitimPano.rcdBilgisi = adp["rcdBilgisi"].toString();
        pano.anaDagitimPano.rcdTestBilgisi = adp["rcdTestBilgisi"].toString();
        pano.anaDagitimPano.hataAkimi = adp["hataAkimi"].toString();
        pano.anaDagitimPano.sistemTopraklamaKesiti = adp["sistemTopraklamaKesiti"].toString();
        pano.anaDagitimPano.anaEspotansiyelKesiti = adp["anaEspotansiyelKesiti"].toString();
        pano.anaDagitimPano.ln = adp["ln"].toString();
        pano.anaDagitimPano.npe = adp["npe"].toString();

        // Genel sonuç
        pano.genelSonuc = panoObj["genelSonuc"].toString();

        // FonksiyonTestleri
        QJsonArray ftArray = panoObj["fonksiyonTestleri"].toArray();
        for (const QJsonValue& ftVal : ftArray) {
            QJsonObject ftObj = ftVal.toObject();
            FonksiyonTesti ft;
            ft.siraNo = ftObj["siraNo"].toInt();
            ft.linye = ftObj["linye"].toString();
            ft.sigortaTipi = ftObj["sigortaTipi"].toString();
            ft.kutupSayisi = ftObj["kutupSayisi"].toInt();
            ft.nominalAkim = ftObj["nominalAkim"].toInt();
            ft.icu = ftObj["icu"].toString();
            ft.ib = ftObj["ib"].toString();
            ft.fazKesiti = ftObj["fazKesiti"].toString();
            ft.notrKesiti = ftObj["notrKesiti"].toString();
            ft.toprakKesiti = ftObj["toprakKesiti"].toString();
            ft.akimKapasitesi = ftObj["akimKapasitesi"].toInt();
            ft.sonuc = ftObj["sonuc"].toString();
            ft.kakrVar = ftObj["kakrVar"].toBool();
            ft.rcd = ftObj["rcd"].toString();
            ft.rcdMa = ftObj["rcdMa"].toString();
            ft.rcdMs = ftObj["rcdMs"].toString();
            ft.kakrYok = ftObj["kakrYok"].toBool();
            ft.isAnaSigorta = ftObj["isAnaSigorta"].toBool();
            pano.fonksiyonTestleri.append(ft);
        }

        // TermalGoruntuler
        QJsonArray tgArray = panoObj["termalGoruntuler"].toArray();
        for (const QJsonValue& tgVal : tgArray) {
            QJsonObject tgObj = tgVal.toObject();
            TermalGoruntu tg;
            tg.imagePath = tgObj["imagePath"].toString();
            tg.tip = tgObj["tip"].toString();
            pano.termalGoruntuler.append(tg);
        }

        // GozleKontrol
        QJsonArray gkArray = panoObj["gozleKontrol"].toArray();
        for (const QJsonValue& gkVal : gkArray) {
            QJsonObject gkObj = gkVal.toObject();
            GozleKontrolMaddesi gk;
            gk.maddeAdi = gkObj["maddeAdi"].toString();
            gk.sonuc = gkObj["sonuc"].toString();
            pano.gozleKontrol.append(gk);
        }

        proje.panolar.append(pano);
    }

    // Projeyi yükle
    loadProjectData(proje);

    m_currentProjectPath = path;
    m_hasUnsavedChanges = false;
    updateWindowTitle();
    m_statusLabel->setText(tr("Proje yüklendi: %1").arg(QFileInfo(path).fileName()));
}

void MainWindow::saveProject() {
    if (m_currentProjectPath.isEmpty()) {
        saveProjectAs();
        return;
    }

    Proje proje = collectAllData();

    // Proje -> JSON dönüşümü
    QJsonObject root;
    root["version"] = "1.0";

    // FirmaBilgileri
    QJsonObject firma;
    firma["firmaAdi"] = proje.firmaBilgileri.firmaAdi;
    firma["kontrolAdresi"] = proje.firmaBilgileri.kontrolAdresi;
    firma["sgkSicil"] = proje.firmaBilgileri.sgkSicil;
    firma["sozlesmeId"] = proje.firmaBilgileri.sozlesmeId;
    firma["raporNumarasi"] = proje.firmaBilgileri.raporNumarasi;
    firma["raporTarihi"] = proje.firmaBilgileri.raporTarihi.toString("dd.MM.yyyy");
    firma["baslangicTarihSaat"] = proje.firmaBilgileri.baslangicTarihSaat.toString("dd.MM.yyyy HH:mm");
    firma["bitisTarihSaat"] = proje.firmaBilgileri.bitisTarihSaat.toString("dd.MM.yyyy HH:mm");
    firma["kontrolEdenAdSoyad"] = proje.firmaBilgileri.kontrolEdenAdSoyad;
    firma["kontrolEdenTc"] = proje.firmaBilgileri.kontrolEdenTc;
    firma["pkNo"] = proje.firmaBilgileri.pkNo;
    firma["teklifNumarasi"] = proje.firmaBilgileri.teklifNumarasi;
    firma["tesisatSayisi"] = proje.firmaBilgileri.tesisatSayisi;
    // Cihaz bilgileri - TAM
    firma["termalCihazAdi"] = proje.firmaBilgileri.termalCihazAdi;
    firma["termalKalibrasyonTarihi"] = proje.firmaBilgileri.termalKalibrasyonTarihi;
    firma["termalKalibrasyonGecerlilik"] = proje.firmaBilgileri.termalKalibrasyonGecerlilik;
    firma["termalSeriNo"] = proje.firmaBilgileri.termalSeriNo;
    firma["termalKalibrasyonNo"] = proje.firmaBilgileri.termalKalibrasyonNo;

    firma["olcumCihazAdi"] = proje.firmaBilgileri.olcumCihazAdi;
    firma["olcumKalibrasyonTarihi"] = proje.firmaBilgileri.olcumKalibrasyonTarihi;
    firma["olcumKalibrasyonGecerlilik"] = proje.firmaBilgileri.olcumKalibrasyonGecerlilik;
    firma["olcumSeriNo"] = proje.firmaBilgileri.olcumSeriNo;
    firma["olcumKalibrasyonNo"] = proje.firmaBilgileri.olcumKalibrasyonNo;

    // Eski format uyumluluk
    firma["cihaz1Adi"] = proje.firmaBilgileri.cihaz1Adi;
    firma["cihaz1SeriNo"] = proje.firmaBilgileri.cihaz1SeriNo;
    firma["cihaz2Adi"] = proje.firmaBilgileri.cihaz2Adi;
    firma["cihaz2SeriNo"] = proje.firmaBilgileri.cihaz2SeriNo;
    root["firmaBilgileri"] = firma;

    // AnaPanoBilgileri
    QJsonObject anaPano;
    anaPano["sebekeTipi"] = proje.anaPanoBilgileri.sebekeTipi;
    anaPano["topraklamaDirenci"] = proje.anaPanoBilgileri.topraklamaDirenci;
    anaPano["sigortaTipiAna"] = proje.anaPanoBilgileri.sigortaTipiAna;
    anaPano["nominalAkimAna"] = proje.anaPanoBilgileri.nominalAkimAna;
    anaPano["rcdBilgisi"] = proje.anaPanoBilgileri.rcdBilgisi;
    anaPano["rcdAnmaAkimi"] = proje.anaPanoBilgileri.rcdAnmaAkimi;
    root["anaPanoBilgileri"] = anaPano;

    // Panolar
    QJsonArray panolar;
    for (const PanoData& pano : proje.panolar) {
        QJsonObject panoObj;
        panoObj["panoAdi"] = pano.panoAdi;
        panoObj["raporNumarasi"] = pano.raporNumarasi;
        panoObj["potansiyelSonuc"] = pano.potansiyelSonuc;

        // AnaDagitimPano
        QJsonObject adp;
        adp["sebekeTipi"] = pano.anaDagitimPano.sebekeTipi;
        adp["topraklamaDirenci"] = pano.anaDagitimPano.topraklamaDirenci;
        adp["sigortaTipiAna"] = pano.anaDagitimPano.sigortaTipiAna;
        adp["nominalAkimAna"] = pano.anaDagitimPano.nominalAkimAna;
        adp["loopPeN"] = pano.anaDagitimPano.loopPeN;
        adp["loopLN"] = pano.anaDagitimPano.loopLN;
        adp["ik3"] = pano.anaDagitimPano.ik3;
        adp["distCevrimEmpedansi"] = pano.anaDagitimPano.distCevrimEmpedansi;
        adp["rcdAnmaAkimi"] = pano.anaDagitimPano.rcdAnmaAkimi;
        // EKSİK ALANLAR
        adp["parafudrTip"] = pano.anaDagitimPano.parafudrTip;
        adp["parafudrImax"] = pano.anaDagitimPano.parafudrImax;
        adp["sistemGerilimi"] = pano.anaDagitimPano.sistemGerilimi;
        adp["enerjiSaglayan"] = pano.anaDagitimPano.enerjiSaglayan;
        adp["trafoGucu"] = pano.anaDagitimPano.trafoGucu;
        adp["sistemFrekans"] = pano.anaDagitimPano.sistemFrekans;
        adp["rcdBilgisi"] = pano.anaDagitimPano.rcdBilgisi;
        adp["rcdTestBilgisi"] = pano.anaDagitimPano.rcdTestBilgisi;
        adp["hataAkimi"] = pano.anaDagitimPano.hataAkimi;
        adp["sistemTopraklamaKesiti"] = pano.anaDagitimPano.sistemTopraklamaKesiti;
        adp["anaEspotansiyelKesiti"] = pano.anaDagitimPano.anaEspotansiyelKesiti;
        adp["ln"] = pano.anaDagitimPano.ln;
        adp["npe"] = pano.anaDagitimPano.npe;
        panoObj["anaDagitimPano"] = adp;

        // Genel sonuç
        panoObj["genelSonuc"] = pano.genelSonuc;

        // FonksiyonTestleri
        QJsonArray ftArray;
        for (const FonksiyonTesti& ft : pano.fonksiyonTestleri) {
            QJsonObject ftObj;
            ftObj["siraNo"] = ft.siraNo;
            ftObj["linye"] = ft.linye;
            ftObj["sigortaTipi"] = ft.sigortaTipi;
            ftObj["kutupSayisi"] = ft.kutupSayisi;
            ftObj["nominalAkim"] = ft.nominalAkim;
            ftObj["icu"] = ft.icu;
            ftObj["ib"] = ft.ib;
            ftObj["fazKesiti"] = ft.fazKesiti;
            ftObj["notrKesiti"] = ft.notrKesiti;
            ftObj["toprakKesiti"] = ft.toprakKesiti;
            ftObj["akimKapasitesi"] = ft.akimKapasitesi;
            ftObj["sonuc"] = ft.sonuc;
            ftObj["kakrVar"] = ft.kakrVar;
            ftObj["rcd"] = ft.rcd;
            ftObj["rcdMa"] = ft.rcdMa;
            ftObj["rcdMs"] = ft.rcdMs;
            ftObj["kakrYok"] = ft.kakrYok;
            ftObj["isAnaSigorta"] = ft.isAnaSigorta;
            ftArray.append(ftObj);
        }
        panoObj["fonksiyonTestleri"] = ftArray;

        // TermalGoruntuler
        QJsonArray tgArray;
        for (const TermalGoruntu& tg : pano.termalGoruntuler) {
            QJsonObject tgObj;
            tgObj["imagePath"] = tg.imagePath;
            tgObj["tip"] = tg.tip;
            tgArray.append(tgObj);
        }
        panoObj["termalGoruntuler"] = tgArray;

        // GozleKontrol
        QJsonArray gkArray;
        for (const GozleKontrolMaddesi& gk : pano.gozleKontrol) {
            QJsonObject gkObj;
            gkObj["maddeAdi"] = gk.maddeAdi;
            gkObj["sonuc"] = gk.sonuc;
            gkArray.append(gkObj);
        }
        panoObj["gozleKontrol"] = gkArray;

        panolar.append(panoObj);
    }
    root["panolar"] = panolar;

    // Dosyaya yaz
    QFile file(m_currentProjectPath);
    if (!file.open(QIODevice::WriteOnly)) {
        QMessageBox::critical(this, tr("Hata"),
            tr("Dosya kaydedilemedi: %1").arg(m_currentProjectPath));
        return;
    }

    QJsonDocument doc(root);
    file.write(doc.toJson());
    file.close();

    m_hasUnsavedChanges = false;
    updateWindowTitle();
    m_statusLabel->setText(tr("Proje kaydedildi: %1").arg(QFileInfo(m_currentProjectPath).fileName()));
}

void MainWindow::saveProjectAs() {
    QString path = QFileDialog::getSaveFileName(this,
        tr("Projeyi Kaydet"), QString(),
        tr("Rapor Projesi (*.erp)"));

    if (path.isEmpty()) return;

    if (!path.endsWith(".erp")) {
        path += ".erp";
    }

    m_currentProjectPath = path;
    saveProject();
}

void MainWindow::loadContract() {
    QString path = QFileDialog::getOpenFileName(this,
        tr("Sözleşme PDF Seç"), QString(),
        tr("PDF Dosyaları (*.pdf);;Tüm Dosyalar (*)"));

    if (!path.isEmpty()) {
        onContractLoad(path);
    }
}

void MainWindow::generateReports() {
    Proje proje = collectAllData();

    if (proje.panolar.isEmpty()) {
         QMessageBox::warning(this, tr("Uyarı"), tr("Rapor oluşturulacak pano bulunamadı."));
         return;
    }

    // Klasör seçimi
    QString defaultDir = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation) + "/Raporlar";
    if (!QDir(defaultDir).exists()) QDir().mkpath(defaultDir);

    QString outputDir = QFileDialog::getExistingDirectory(this, tr("Rapor Kayıt Klasörü"), defaultDir);
    if (outputDir.isEmpty()) return;

    m_progressBar->setVisible(true);
    m_progressBar->setMaximum(proje.panolar.count());
    m_statusLabel->setText(tr("Raporlar oluşturuluyor..."));
    QApplication::processEvents();

    ReportGenerator generator;
    generator.setOutputDirectory(outputDir);

    // Şablon yolu (önce exe yanı, sonra bir üst klasör)
    QString templatePath = QCoreApplication::applicationDirPath() + "/sablon/rapor_sablonu.docx";
    if (!QFile::exists(templatePath)) {
        templatePath = QCoreApplication::applicationDirPath() + "/../sablon/rapor_sablonu.docx";
    }

    // Debug: Şablon yoksa uyarı ver
    // if (!QFile::exists(templatePath)) ...

    generator.setTemplatePath(templatePath);

    QStringList createdFiles = generator.generateAllReports(proje);

    m_progressBar->setVisible(false);

    if (createdFiles.isEmpty()) {
        QMessageBox::critical(this, tr("Hata"),
            tr("Rapor oluşturulamadı!\nDetay: %1").arg(generator.errorString()));
        m_statusLabel->setText(tr("Rapor oluşturma başarısız."));
    } else {
        QMessageBox::information(this, tr("Başarılı"),
            tr("%1 adet rapor oluşturuldu.\n\nKayıt Yeri:\n%2")
            .arg(createdFiles.count())
            .arg(outputDir));

        m_statusLabel->setText(tr("Raporlar tamamlandı."));

        // Klasörü aç
        QDesktopServices::openUrl(QUrl::fromLocalFile(outputDir));
    }
}

void MainWindow::generateSingleReport(int panoIndex) {
    // TODO: Tek rapor oluşturma
}

Proje MainWindow::collectAllData() const {
    Proje proje;
    proje.firmaBilgileri = m_commonInfoPanel->getFirmaBilgileri();
    proje.anaPanoBilgileri = m_commonInfoPanel->getAnaPanoBilgileri();

    // Proje görseli varsa al
    QString projeGorseliPath = m_commonInfoPanel->getProjeGorseliPath();

    // Rapor numarası tabanı (TPK2026-1234)
    QString baseRaporNo = proje.firmaBilgileri.raporNumarasi;

    for (int i = 0; i < m_panoTabs->count(); ++i) {
        PanoTabWidget* tab = qobject_cast<PanoTabWidget*>(m_panoTabs->widget(i));
        if (tab) {
            PanoData panoData = tab->getData();

            // Rapor numarası formatı: TPK2026-1234-1, TPK2026-1234-2, ...
            if (!baseRaporNo.isEmpty()) {
                panoData.raporNumarasi = QString("%1-%2").arg(baseRaporNo).arg(i + 1);
            }

            // Proje görseli varsa termal görüntülere ekle (Python gibi)
            if (!projeGorseliPath.isEmpty() && QFile::exists(projeGorseliPath)) {
                TermalGoruntu projeGorsel;
                projeGorsel.imagePath = projeGorseliPath;
                projeGorsel.tip = "proje_gorseli";
                panoData.termalGoruntuler.append(projeGorsel);
            }

            proje.panolar.append(panoData);
        }
    }

    return proje;
}

void MainWindow::loadProjectData(const Proje& proje) {
    clearAll();

    m_commonInfoPanel->setFirmaBilgileri(proje.firmaBilgileri);
    m_commonInfoPanel->setAnaPanoBilgileri(proje.anaPanoBilgileri);

    for (const PanoData& pano : proje.panolar) {
        ++m_panoCounter;
        PanoTabWidget* tab = new PanoTabWidget(m_panoCounter, this);
        connect(tab, &PanoTabWidget::dataChanged, this, &MainWindow::onProjectModified);

        // Pano adı değiştiğinde tab başlığını güncelle
        connect(tab, &PanoTabWidget::panoNameChanged, this, [this, tab](int, const QString& newName) {
            int tabIndex = m_panoTabs->indexOf(tab);
            if (tabIndex >= 0) {
                QString title = newName.isEmpty() ? QString("Pano %1").arg(tab->panoIndex()) : newName;
                m_panoTabs->setTabText(tabIndex, title);
            }
        });

        tab->setData(pano);

        QString tabName = pano.panoAdi.isEmpty() ?
            QString("Pano %1").arg(m_panoCounter) : pano.panoAdi;
        m_panoTabs->addTab(tab, tabName);
    }
}

void MainWindow::showAboutDialog() {
    QMessageBox::about(this, tr("Hakkında"),
        tr("<h2>Elektrik Rapor Sistemi</h2>"
           "<p>Versiyon 1.0.0 (C++ Edition)</p>"
           "<p>Qt %1 ile oluşturuldu.</p>"
           "<p>© 2024</p>").arg(QT_VERSION_STR));
}

void MainWindow::showSettingsDialog() {
    // TODO: Ayarlar dialog
}

// ==================== CommonInfoPanel ====================

CommonInfoPanel::CommonInfoPanel(QWidget* parent)
    : QWidget(parent)
{
    setupUi();
}

void CommonInfoPanel::setupUi() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(4, 4, 4, 4);
    mainLayout->setSpacing(4);

    // ===== BAŞLIK =====
    QLabel* title = new QLabel(tr("<b>🏢 Ortak Bilgiler</b>"));
    title->setAlignment(Qt::AlignCenter);
    mainLayout->addWidget(title);

    // ===== İKİ SEKMELİ YAPI =====
    QTabWidget* tabs = new QTabWidget();

    // ----- 1. SEKME: FİRMA -----
    QWidget* firmaWidget = new QWidget();
    QVBoxLayout* firmaLayout = new QVBoxLayout(firmaWidget);
    firmaLayout->setSpacing(4);

    // Scroll Area ekle (Çok fazla alan olduğu için)
    QScrollArea* scrollArea = new QScrollArea();
    scrollArea->setWidgetResizable(true);
    QWidget* scrollContent = new QWidget();
    QVBoxLayout* scrollLayout = new QVBoxLayout(scrollContent);
    scrollLayout->setSpacing(4);
    scrollArea->setWidget(scrollContent);

    // Helper lambda for creating rows
    auto addRow = [&](const QString& label, QWidget* widget) {
        QHBoxLayout* row = new QHBoxLayout();
        QLabel* lbl = new QLabel(label);
        lbl->setFixedWidth(120); // Python'daki gibi sabit genişlik
        row->addWidget(lbl);
        row->addWidget(widget);
        scrollLayout->addLayout(row);
    };

    // 1. Firma Adı
    m_firmaAdi = new QLineEdit();
    addRow(tr("Firma Adı:"), m_firmaAdi);

    // 2. Tesis Adresi
    m_kontrolAdresi = new QLineEdit();
    addRow(tr("Tesis Adresi:"), m_kontrolAdresi);

    // 3. SGK Sicil No
    m_sgkSicil = new QLineEdit();
    addRow(tr("SGK Sicil No:"), m_sgkSicil);

    // 4. Sözleşme ID
    m_sozlesmeId = new QLineEdit();
    addRow(tr("Sözleşme ID:"), m_sozlesmeId);

    // 5. Teklif Numarası
    m_teklifNumarasi = new QLineEdit();
    addRow(tr("Teklif Numarası:"), m_teklifNumarasi);

    // 6. Rapor Tarihi
    m_raporTarihi = new QDateEdit();
    m_raporTarihi->setDate(QDate::currentDate());
    m_raporTarihi->setCalendarPopup(true);
    m_raporTarihi->setDisplayFormat("dd.MM.yyyy");
    addRow(tr("Rapor Tarihi:"), m_raporTarihi);

    // 7. Kontrol Başlangıç
    m_baslangicTarih = new QDateEdit();
    m_baslangicTarih->setDate(QDate::currentDate());
    m_baslangicTarih->setCalendarPopup(true);
    m_baslangicTarih->setDisplayFormat("dd.MM.yyyy");
    addRow(tr("Kontrol Başlangıç:"), m_baslangicTarih);

    // 8. Kontrol Bitiş
    m_bitisTarih = new QDateEdit();
    m_bitisTarih->setDate(QDate::currentDate());
    m_bitisTarih->setCalendarPopup(true);
    m_bitisTarih->setDisplayFormat("dd.MM.yyyy");
    addRow(tr("Kontrol Bitiş:"), m_bitisTarih);

    // 9. Bir Sonraki Kontrol (OTOMATİK: başlangıç + 1 yıl)
    m_birSonrakiKontrol = new QDateEdit();
    m_birSonrakiKontrol->setDate(QDate::currentDate().addYears(1));
    m_birSonrakiKontrol->setCalendarPopup(true);
    m_birSonrakiKontrol->setDisplayFormat("dd.MM.yyyy");
    addRow(tr("Bir Sonraki Kontrol:"), m_birSonrakiKontrol);

    // Başlangıç tarihi değişince "Bir Sonraki Kontrol" otomatik +1 yıl olsun
    connect(m_baslangicTarih, &QDateEdit::dateChanged, this, [this](const QDate& date) {
        m_birSonrakiKontrol->setDate(date.addYears(1));
    });

    // 10. Kontrol Eden
    m_kontrolEdenAdSoyad = new QLineEdit();
    addRow(tr("Kontrol Eden:"), m_kontrolEdenAdSoyad);

    // 11. Belge No (Python'daki 'Belge No' = PDF'deki 'pk_no')
    m_pkNo = new QLineEdit();
    addRow(tr("Belge No:"), m_pkNo);

    // Rapor No Prefix
    scrollLayout->addSpacing(10);
    QHBoxLayout* prefixRow = new QHBoxLayout();
    prefixRow->addWidget(new QLabel(tr("Rapor No Prefix:")));
    m_raporNumarasi = new QLineEdit();
    m_raporNumarasi->setPlaceholderText("TPK2025-4001");
    prefixRow->addWidget(m_raporNumarasi);
    prefixRow->addWidget(new QLabel("-Y"));
    scrollLayout->addLayout(prefixRow);

    QLabel* infoLbl = new QLabel(tr("💡 Her pano için rapor numarası otomatik artar"));
    infoLbl->setStyleSheet("color: gray; font-size: 10px;");
    scrollLayout->addWidget(infoLbl);

    // === Tesis Belgeleri ===
    scrollLayout->addSpacing(15);
    QLabel* tesisLbl = new QLabel(tr("📋 Tesis Belgeleri"));
    tesisLbl->setStyleSheet("color: #03a9f4; font-weight: bold;");
    scrollLayout->addWidget(tesisLbl);

    // Proje Var/Yok
    QComboBox* cmbProje = new QComboBox();
    cmbProje->addItems({tr("Var"), tr("Yok")});
    addRow(tr("Tesise ait proje var mı?"), cmbProje);

    // Proje Görseli (Python'daki proje_gorseli_path karşılığı)
    QHBoxLayout* projeGorselRow = new QHBoxLayout();
    QLabel* projeGorselLbl = new QLabel(tr("   └─ Proje Görseli:"));
    projeGorselLbl->setFixedWidth(120);
    projeGorselLbl->setStyleSheet("color: #888888;");
    projeGorselRow->addWidget(projeGorselLbl);

    m_projeGorselBtn = new QPushButton(tr("📷 Görsel Ekle"));
    m_projeGorselBtn->setStyleSheet("background-color: #2E7D32; color: white; padding: 4px 8px;");
    m_projeGorselBtn->setFixedWidth(100);
    connect(m_projeGorselBtn, &QPushButton::clicked, this, &CommonInfoPanel::selectProjeGorseli);
    projeGorselRow->addWidget(m_projeGorselBtn);

    m_projeGorselLabel = new QLabel();
    m_projeGorselLabel->setStyleSheet("color: #4CAF50; font-size: 10px;");
    projeGorselRow->addWidget(m_projeGorselLabel);
    projeGorselRow->addStretch();
    scrollLayout->addLayout(projeGorselRow);

    // Tek Hat Var/Yok
    QComboBox* cmbTekHat = new QComboBox();
    cmbTekHat->addItems({tr("Var"), tr("Yok")});
    addRow(tr("Tek hat şeması var mı?"), cmbTekHat);

    // Yapı Cinsi
    QComboBox* cmbYapi = new QComboBox();
    cmbYapi->addItems({tr("Ev"), tr("Ticari"), tr("Endüstri"), tr("Diğer")});
    cmbYapi->setCurrentText("Ticari");
    addRow(tr("Yapı cinsi:"), cmbYapi);

    // === Termal Kamera ===
    scrollLayout->addSpacing(15);
    QLabel* termalLbl = new QLabel(tr("📷 Termal Kamera"));
    termalLbl->setStyleSheet("color: #ff9800; font-weight: bold;");
    scrollLayout->addWidget(termalLbl);

    m_termalCihazAdi = new QLineEdit();
    addRow(tr("Cihaz Adı"), m_termalCihazAdi);

    m_termalKalibrasyonTarihi = new QLineEdit();
    addRow(tr("Kalibrasyon Tarihi"), m_termalKalibrasyonTarihi);

    m_termalKalibrasyonGecerlilik = new QLineEdit();
    addRow(tr("Kalibrasyon Geçerlilik"), m_termalKalibrasyonGecerlilik);

    m_termalSeriNo = new QLineEdit();
    addRow(tr("Seri No"), m_termalSeriNo);

    m_termalKalibrasyonNo = new QLineEdit();
    addRow(tr("Kalibrasyon No"), m_termalKalibrasyonNo);

    // === Ölçüm Cihazı ===
    scrollLayout->addSpacing(15);
    QLabel* olcumLbl = new QLabel(tr("🔌 Ölçüm Cihazı"));
    olcumLbl->setStyleSheet("color: #2196f3; font-weight: bold;");
    scrollLayout->addWidget(olcumLbl);

    m_olcumCihazAdi = new QLineEdit();
    addRow(tr("Cihaz Adı"), m_olcumCihazAdi);

    m_olcumKalibrasyonTarihi = new QLineEdit();
    addRow(tr("Kalibrasyon Tarihi"), m_olcumKalibrasyonTarihi);

    m_olcumKalibrasyonGecerlilik = new QLineEdit();
    addRow(tr("Kalibrasyon Geçerlilik"), m_olcumKalibrasyonGecerlilik);

    m_olcumSeriNo = new QLineEdit();
    addRow(tr("Seri No"), m_olcumSeriNo);

    m_olcumKalibrasyonNo = new QLineEdit();
    addRow(tr("Kalibrasyon No"), m_olcumKalibrasyonNo);

    // Spacer
    scrollLayout->addStretch();

    firmaLayout->addWidget(scrollArea);

    // --- Hidden Fields ---
    m_kontrolEdenTc = new QLineEdit(); // Not shown in UI but kept for data
    m_kontrolEdenTc->setVisible(false);
    m_birSonrakiKontrol = new QDateEdit();
    m_birSonrakiKontrol->setVisible(false);

    tabs->addTab(firmaWidget, tr("Firma"));

    // ----- 2. SEKME: ANA PANO -----
    QWidget* anaPanoWidget = new QWidget();
    QVBoxLayout* anaPanoLayout = new QVBoxLayout(anaPanoWidget);
    anaPanoLayout->setSpacing(4);

    // Helper lambda for ana pano rows
    auto addPanoRow = [&](const QString& label, QWidget* widget) {
        QHBoxLayout* row = new QHBoxLayout();
        QLabel* lbl = new QLabel(label);
        lbl->setFixedWidth(160); // Biraz daha geniş
        row->addWidget(lbl);
        row->addWidget(widget);
        anaPanoLayout->addLayout(row);
    };

    // Enerji Sağlayan Kuruluş
    m_enerjiSaglayan = new QLineEdit();
    m_enerjiSaglayan->setText("TEDAŞ");
    addPanoRow(tr("Enerji Sağlayan Kuruluş:"), m_enerjiSaglayan);

    // Şebeke Tipi
    m_sebekeTipi = new QComboBox();
    m_sebekeTipi->addItems({"TN-S", "TN-C", "TN-C-S", "TT", "IT"});
    addPanoRow(tr("Şebeke Tipi:"), m_sebekeTipi);

    // Temel Topraklama Direnci
    m_temelTopraklamaDirenci = new QLineEdit();
    addPanoRow(tr("Temel Topraklama Direnci (Ω):"), m_temelTopraklamaDirenci);

    // Dış Çevrim Empedansı Z_E
    m_disCevrimEmpedansi = new QLineEdit();
    addPanoRow(tr("Dış Çevrim Empedansı Z_E (Ω):"), m_disCevrimEmpedansi);

    // Ana Kesici Tipi
    m_anaKesiciTipi = new QComboBox();
    m_anaKesiciTipi->addItems({"C", "B", "D", "K"});
    addPanoRow(tr("Ana Kesici Tipi:"), m_anaKesiciTipi);

    // Ana Kesici Nominal Akımı
    m_anaKesiciNominalAkim = new QLineEdit();
    addPanoRow(tr("Ana Kesici Nominal Akımı:"), m_anaKesiciNominalAkim);

    // Ana RCD Tipi
    m_anaRcdTipi = new QComboBox();
    m_anaRcdTipi->addItems({"TOROİD", "KAKR"});
    addPanoRow(tr("Ana RCD Tipi:"), m_anaRcdTipi);

    // Ana RCD Anma Akımı
    m_anaRcdAnmaAkimi = new QLineEdit();
    m_anaRcdAnmaAkimi->setPlaceholderText("Örn: 30mA");
    addPanoRow(tr("Ana RCD Anma Akımı:"), m_anaRcdAnmaAkimi);

    // Ana RCD Test Bilgisi
    m_anaRcdTestBjlgisi = new QLineEdit();
    m_anaRcdTestBjlgisi->setPlaceholderText("Örn: 24mA / 15ms");
    addPanoRow(tr("Ana RCD Test (Akım/Süre):"), m_anaRcdTestBjlgisi);

    // Sistem Topraklama Kesiti
    m_sistemTopraklamaKesiti = new QComboBox();
    m_sistemTopraklamaKesiti->setEditable(true);
    m_sistemTopraklamaKesiti->addItems({"6", "10", "16", "25", "35", "50", "70", "95", "120"});
    addPanoRow(tr("Sistem Topraklama Kesiti (mm²):"), m_sistemTopraklamaKesiti);

    // Ana Eşpotansiyel Kesiti
    m_anaEspotansiyelKesiti = new QComboBox();
    m_anaEspotansiyelKesiti->setEditable(true);
    m_anaEspotansiyelKesiti->addItems({"4", "6", "10", "16", "25", "35", "50"});
    addPanoRow(tr("Ana Eşpotansiyel Kesiti (mm²):"), m_anaEspotansiyelKesiti);

    anaPanoLayout->addStretch();
    tabs->addTab(anaPanoWidget, tr("Ana Pano"));

    mainLayout->addWidget(tabs, 1);

    // ===== ALT BUTONLAR =====
    QHBoxLayout* btnRow = new QHBoxLayout();

    QPushButton* kaydetBtn = new QPushButton(tr("📁 Kaydet"));
    btnRow->addWidget(kaydetBtn);

    QPushButton* acBtn = new QPushButton(tr("📂 Aç"));
    btnRow->addWidget(acBtn);

    mainLayout->addLayout(btnRow);

    // Legacy fields initialization (to prevent crash)
    m_cihaz1Adi = new QLineEdit(); m_cihaz1Adi->setVisible(false);
    m_cihaz1SeriNo = new QLineEdit(); m_cihaz1SeriNo->setVisible(false);
    m_cihaz1Kalibrasyon = new QLineEdit(); m_cihaz1Kalibrasyon->setVisible(false);
    m_cihaz2Adi = new QLineEdit(); m_cihaz2Adi->setVisible(false);
    m_cihaz2SeriNo = new QLineEdit(); m_cihaz2SeriNo->setVisible(false);
    m_cihaz2Kalibrasyon = new QLineEdit(); m_cihaz2Kalibrasyon->setVisible(false);
    m_cihaz3Adi = new QLineEdit(); m_cihaz3Adi->setVisible(false);
    m_cihaz3SeriNo = new QLineEdit(); m_cihaz3SeriNo->setVisible(false);
    m_cihaz3Kalibrasyon = new QLineEdit(); m_cihaz3Kalibrasyon->setVisible(false);

    // Add them to layout so they have a parent (though not strictly necessary for crash prevention if parented)
    firmaLayout->addWidget(m_cihaz1Adi);
    firmaLayout->addWidget(m_cihaz1SeriNo);
    firmaLayout->addWidget(m_cihaz1Kalibrasyon);

    // Sözleşme PDF Yükle
    QPushButton* loadBtn = new QPushButton(tr("📄 Sözleşme PDF Yükle"));
    loadBtn->setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 8px;");
    connect(loadBtn, &QPushButton::clicked, this, &CommonInfoPanel::loadContract);
    mainLayout->addWidget(loadBtn);

    // TOPLU RAPOR OLUŞTUR
    QPushButton* topluBtn = new QPushButton(tr("📊 TOPLU RAPOR OLUŞTUR"));
    topluBtn->setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 12px; font-size: 14px;");
    connect(topluBtn, &QPushButton::clicked, this, &CommonInfoPanel::generateReportsRequested);
    mainLayout->addWidget(topluBtn);

    // Connect signals for new fields
    connect(m_enerjiSaglayan, &QLineEdit::textChanged, this, &CommonInfoPanel::dataChanged);
    // ... add others if needed, typically dataChanged triggers auto-save or UI update
}

FirmaBilgileri CommonInfoPanel::getFirmaBilgileri() const {
    FirmaBilgileri firma;
    firma.firmaAdi = m_firmaAdi->text();
    firma.kontrolAdresi = m_kontrolAdresi->text();
    firma.sgkSicil = m_sgkSicil->text();
    firma.raporNumarasi = m_raporNumarasi->text();
    firma.sozlesmeId = m_sozlesmeId->text();
    firma.pkNo = m_pkNo->text();
    firma.teklifNumarasi = m_teklifNumarasi->text();  // tklf için
    firma.raporTarihi = m_raporTarihi->date();
    firma.baslangicTarihSaat = QDateTime(m_baslangicTarih->date(), QTime(8, 30));
    firma.bitisTarihSaat = QDateTime(m_bitisTarih->date(), QTime(17, 30));
    firma.birSonrakiKontrol = m_birSonrakiKontrol->date();
    firma.kontrolEdenAdSoyad = m_kontrolEdenAdSoyad->text();
    firma.kontrolEdenTc = m_kontrolEdenTc->text();

    // Termal Kamera
    firma.termalCihazAdi = m_termalCihazAdi->text();
    firma.termalKalibrasyonTarihi = m_termalKalibrasyonTarihi->text();
    firma.termalKalibrasyonGecerlilik = m_termalKalibrasyonGecerlilik->text();
    firma.termalSeriNo = m_termalSeriNo->text();
    firma.termalKalibrasyonNo = m_termalKalibrasyonNo->text();

    // Ölçüm Cihazı
    firma.olcumCihazAdi = m_olcumCihazAdi->text();
    firma.olcumKalibrasyonTarihi = m_olcumKalibrasyonTarihi->text();
    firma.olcumKalibrasyonGecerlilik = m_olcumKalibrasyonGecerlilik->text();
    firma.olcumSeriNo = m_olcumSeriNo->text();
    firma.olcumKalibrasyonNo = m_olcumKalibrasyonNo->text();

    // Cihaz bilgileri (Legacy)
    firma.cihaz1Adi = m_cihaz1Adi->text();
    firma.cihaz1SeriNo = m_cihaz1SeriNo->text();
    firma.cihaz1KalibrasyonTarihi = m_cihaz1Kalibrasyon->text();
    firma.cihaz2Adi = m_cihaz2Adi->text();
    firma.cihaz2SeriNo = m_cihaz2SeriNo->text();
    firma.cihaz2KalibrasyonTarihi = m_cihaz2Kalibrasyon->text();

    return firma;
}

void CommonInfoPanel::setFirmaBilgileri(const FirmaBilgileri& firma) {
    m_firmaAdi->setText(firma.firmaAdi);
    m_kontrolAdresi->setText(firma.kontrolAdresi);
    m_sgkSicil->setText(firma.sgkSicil);
    m_raporNumarasi->setText(firma.raporNumarasi);
    m_sozlesmeId->setText(firma.sozlesmeId);
    m_pkNo->setText(firma.pkNo);
    m_teklifNumarasi->setText(firma.teklifNumarasi);  // tklf için
    m_raporTarihi->setDate(firma.raporTarihi);
    m_baslangicTarih->setDate(firma.baslangicTarihSaat.date());
    m_bitisTarih->setDate(firma.bitisTarihSaat.date());
    m_birSonrakiKontrol->setDate(firma.birSonrakiKontrol);
    m_kontrolEdenAdSoyad->setText(firma.kontrolEdenAdSoyad);
    m_kontrolEdenTc->setText(firma.kontrolEdenTc);

    // Termal
    m_termalCihazAdi->setText(firma.termalCihazAdi);
    m_termalKalibrasyonTarihi->setText(firma.termalKalibrasyonTarihi);
    m_termalKalibrasyonGecerlilik->setText(firma.termalKalibrasyonGecerlilik);
    m_termalSeriNo->setText(firma.termalSeriNo);
    m_termalKalibrasyonNo->setText(firma.termalKalibrasyonNo);

    // Ölçüm
    m_olcumCihazAdi->setText(firma.olcumCihazAdi);
    m_olcumKalibrasyonTarihi->setText(firma.olcumKalibrasyonTarihi);
    m_olcumKalibrasyonGecerlilik->setText(firma.olcumKalibrasyonGecerlilik);
    m_olcumSeriNo->setText(firma.olcumSeriNo);
    m_olcumKalibrasyonNo->setText(firma.olcumKalibrasyonNo);

    // Eski alanları da güncelle (ne olur ne olmaz)
    m_cihaz1Adi->setText(firma.cihaz1Adi);
    m_cihaz1SeriNo->setText(firma.cihaz1SeriNo);
    m_cihaz1Kalibrasyon->setText(firma.cihaz1KalibrasyonTarihi);
    m_cihaz2Adi->setText(firma.cihaz2Adi);
    m_cihaz2SeriNo->setText(firma.cihaz2SeriNo);
    m_cihaz2Kalibrasyon->setText(firma.cihaz2KalibrasyonTarihi);
}

AnaDagitimPano CommonInfoPanel::getAnaPanoBilgileri() const {
    AnaDagitimPano data;
    data.enerjiSaglayan = m_enerjiSaglayan->text();
    data.sebekeTipi = m_sebekeTipi->currentText();
    data.topraklamaDirenci = m_temelTopraklamaDirenci->text();
    data.distCevrimEmpedansi = m_disCevrimEmpedansi->text();
    data.sigortaTipiAna = m_anaKesiciTipi->currentText();
    data.nominalAkimAna = m_anaKesiciNominalAkim->text().toInt();
    data.rcdBilgisi = m_anaRcdTipi->currentText();
    data.rcdAnmaAkimi = m_anaRcdAnmaAkimi->text();
    data.hataAkimi = m_anaRcdTestBjlgisi->text(); // Test bilgisi olarak kullanıyoruz
    data.sistemTopraklamaKesiti = m_sistemTopraklamaKesiti->currentText();
    data.anaEspotansiyelKesiti = m_anaEspotansiyelKesiti->currentText();
    return data;
}

void CommonInfoPanel::setAnaPanoBilgileri(const AnaDagitimPano& data) {
    m_enerjiSaglayan->setText(data.enerjiSaglayan);
    m_sebekeTipi->setCurrentText(data.sebekeTipi);
    m_temelTopraklamaDirenci->setText(data.topraklamaDirenci);
    m_disCevrimEmpedansi->setText(data.distCevrimEmpedansi);
    m_anaKesiciTipi->setCurrentText(data.sigortaTipiAna);
    m_anaKesiciNominalAkim->setText(QString::number(data.nominalAkimAna));
    m_anaRcdTipi->setCurrentText(data.rcdBilgisi);
    m_anaRcdAnmaAkimi->setText(data.rcdAnmaAkimi);
    m_anaRcdTestBjlgisi->setText(data.hataAkimi);
    m_sistemTopraklamaKesiti->setCurrentText(data.sistemTopraklamaKesiti);
    m_anaEspotansiyelKesiti->setCurrentText(data.anaEspotansiyelKesiti);
}

void CommonInfoPanel::clear() {
    m_firmaAdi->clear();
    m_kontrolAdresi->clear();
    m_sgkSicil->clear();
    m_raporNumarasi->clear();
    m_sozlesmeId->clear();
    m_pkNo->clear();
    m_raporTarihi->setDate(QDate::currentDate());
    m_baslangicTarih->setDate(QDate::currentDate());
    m_bitisTarih->setDate(QDate::currentDate());
    m_birSonrakiKontrol->setDate(QDate::currentDate().addYears(1));
    m_kontrolEdenAdSoyad->clear();
    m_kontrolEdenTc->clear();

    // Ana Pano temizle
    m_enerjiSaglayan->setText("TEDAŞ");
    m_sebekeTipi->setCurrentIndex(0);
    m_temelTopraklamaDirenci->clear();
    m_disCevrimEmpedansi->clear();
    m_anaKesiciNominalAkim->clear();
    m_anaRcdAnmaAkimi->clear();
    m_anaRcdTestBjlgisi->clear();

    // Cihaz bilgileri
    m_cihaz1Adi->clear();
    m_cihaz1SeriNo->clear();
    m_cihaz1Kalibrasyon->clear();
    m_cihaz2Adi->clear();
    m_cihaz2SeriNo->clear();
    m_cihaz2Kalibrasyon->clear();
    m_cihaz3Adi->clear();
    m_cihaz3SeriNo->clear();
    m_cihaz3Kalibrasyon->clear();

    // Proje görseli temizle
    m_projeGorseliPath.clear();
    m_projeGorselBtn->setText(tr("📷 Görsel Ekle"));
    m_projeGorselBtn->setStyleSheet("background-color: #2E7D32; color: white; padding: 4px 8px;");
    m_projeGorselLabel->clear();
}

void CommonInfoPanel::loadContract() {
    emit contractLoaded(QString());
}

void CommonInfoPanel::selectProjeGorseli() {
    QString filePath = QFileDialog::getOpenFileName(
        this,
        tr("Proje Görselini Seçin"),
        QString(),
        tr("Görüntü dosyaları (*.png *.jpg *.jpeg *.bmp *.gif);;Tüm dosyalar (*.*)")
    );

    if (!filePath.isEmpty()) {
        m_projeGorseliPath = filePath;

        QFileInfo fileInfo(filePath);
        QString filename = fileInfo.fileName();
        if (filename.length() > 15) {
            filename = filename.left(12) + "...";
        }

        m_projeGorselBtn->setText(QString::fromUtf8("✓ %1").arg(filename));
        m_projeGorselBtn->setStyleSheet("background-color: #1B5E20; color: white; padding: 4px 8px;");
        m_projeGorselLabel->setText(QString::fromUtf8("✓ %1").arg(fileInfo.fileName()));

        emit dataChanged();
    }
}

} // namespace RaporSistemi

/**
 * MainWindow.h
 *
 * Ana uygulama penceresi.
 * Python karşılığı: multi_pano_gui.py (4192 satır)
 */

#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTabWidget>
#include <QLineEdit>
#include <QDateEdit>
#include <QComboBox>
#include <QPushButton>
#include <QLabel>
#include <QProgressBar>
#include <QTimer>
#include <memory>
#include "DataModels.h"

namespace RaporSistemi {

class PanoTabWidget;
class CommonInfoPanel;

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow();

signals:
    void projectChanged();
    void reportGenerationStarted();
    void reportGenerationFinished(bool success);

public slots:
    // Pano işlemleri
    void addNewPano();
    void removePano(int index);
    void duplicatePano(int index);

    // Proje işlemleri
    void newProject();
    void openProject();
    void saveProject();
    void saveProjectAs();

    // Rapor işlemleri
    void generateReports();
    void generateSingleReport(int panoIndex);

    // Sözleşme yükleme
    void loadContract();

private slots:
    void onTabCloseRequested(int index);
    void onTabChanged(int index);
    void updateWindowTitle();
    void showAboutDialog();
    void showSettingsDialog();
    void onProjectModified();
    void onContractLoad(const QString& pdfPath);
    void updatePersonFile();

private:
    void setupUi();
    void setupMenuBar();
    void setupToolBar();
    void setupStatusBar();
    void setupShortcuts();

    // Veri işlemleri
    Proje collectAllData() const;
    void loadProjectData(const Proje& proje);
    void clearAll();

    // UI bileşenleri
    QTabWidget* m_panoTabs;
    CommonInfoPanel* m_commonInfoPanel;
    QProgressBar* m_progressBar;
    QLabel* m_statusLabel;

    // Proje durumu
    QString m_currentProjectPath;
    bool m_hasUnsavedChanges = false;

    // Pano sayacı
    int m_panoCounter = 0;
};

/**
 * Ortak bilgiler paneli (Firma bilgileri, cihaz bilgileri vs.)
 */
class CommonInfoPanel : public QWidget {
    Q_OBJECT

public:
    explicit CommonInfoPanel(QWidget* parent = nullptr);

    FirmaBilgileri getFirmaBilgileri() const;
    void setFirmaBilgileri(const FirmaBilgileri& firma);
    AnaDagitimPano getAnaPanoBilgileri() const;
    void setAnaPanoBilgileri(const AnaDagitimPano& data);
    QString getProjeGorseliPath() const { return m_projeGorseliPath; }
    void clear();

signals:
    void dataChanged();
    void contractLoaded(const QString& path);
    void generateReportsRequested();

public slots:
    void loadContract();

private:
    void setupUi();

    // Firma bilgileri alanları
    QLineEdit* m_firmaAdi;
    QLineEdit* m_kontrolAdresi;
    QLineEdit* m_sgkSicil;
    QLineEdit* m_raporNumarasi;
    QDateEdit* m_raporTarihi;
    QLineEdit* m_sozlesmeId;
    QDateEdit* m_baslangicTarih;
    QDateEdit* m_bitisTarih;
    QDateEdit* m_birSonrakiKontrol;

    // Kontrol eden kişi
    QLineEdit* m_kontrolEdenAdSoyad;
    QLineEdit* m_kontrolEdenTc;
    QLineEdit* m_pkNo;
    QLineEdit* m_teklifNumarasi;  // tklf placeholder için

    // Termal Kamera
    QLineEdit* m_termalCihazAdi;
    QLineEdit* m_termalKalibrasyonTarihi;
    QLineEdit* m_termalKalibrasyonGecerlilik;
    QLineEdit* m_termalSeriNo;
    QLineEdit* m_termalKalibrasyonNo;

    // Ölçüm Cihazı
    QLineEdit* m_olcumCihazAdi;
    QLineEdit* m_olcumKalibrasyonTarihi;
    QLineEdit* m_olcumKalibrasyonGecerlilik;
    QLineEdit* m_olcumSeriNo;
    QLineEdit* m_olcumKalibrasyonNo;

    // Eski alanlar (silinebilir ama şimdilik tutalım)
    // Eski alanlar (silinebilir ama şimdilik tutalım)
    QLineEdit* m_cihaz1Adi;
    QLineEdit* m_cihaz1SeriNo;
    QLineEdit* m_cihaz1Kalibrasyon;
    QLineEdit* m_cihaz2Adi;
    QLineEdit* m_cihaz2SeriNo;
    QLineEdit* m_cihaz2Kalibrasyon;
    QLineEdit* m_cihaz3Adi;
    QLineEdit* m_cihaz3SeriNo;
    QLineEdit* m_cihaz3Kalibrasyon;

    // === YENİ: Ana Pano (Global) Bilgileri ===
    QLineEdit* m_enerjiSaglayan;
    QComboBox* m_sebekeTipi;
    QLineEdit* m_temelTopraklamaDirenci;
    QLineEdit* m_disCevrimEmpedansi;
    QComboBox* m_anaKesiciTipi;
    QLineEdit* m_anaKesiciNominalAkim;
    QComboBox* m_anaRcdTipi;
    QLineEdit* m_anaRcdAnmaAkimi;
    QLineEdit* m_anaRcdTestBjlgisi; // Test akımı + süresi
    QComboBox* m_sistemTopraklamaKesiti;
    QComboBox* m_anaEspotansiyelKesiti;

    // Proje görseli
    QString m_projeGorseliPath;
    QPushButton* m_projeGorselBtn;
    QLabel* m_projeGorselLabel;

private slots:
    void selectProjeGorseli();
};

} // namespace RaporSistemi

#endif // MAINWINDOW_H

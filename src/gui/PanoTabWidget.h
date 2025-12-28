/**
 * PanoTabWidget.h
 *
 * Tek bir pano için veri girişi widget'ı.
 * Python'daki gibi 3 iç sekme: Gözle Kontrol, Fonksiyon Testleri, Termal
 */

#ifndef PANOTABWIDGET_H
#define PANOTABWIDGET_H

#include "DataModels.h"
#include <QCheckBox>
#include <QComboBox>
#include <QGroupBox>
#include <QLineEdit>
#include <QSpinBox>
#include <QTabWidget>
#include <QWidget>
#include <memory>

namespace RaporSistemi {

class FonksiyonTestleriTable;
class DragDropWidget;
class GozleKontrolWidget;

class PanoTabWidget : public QWidget {
  Q_OBJECT

public:
  explicit PanoTabWidget(int panoIndex, QWidget *parent = nullptr);
  ~PanoTabWidget();

  /**
   * Pano verilerini döndürür.
   */
  PanoData getData() const;

  /**
   * Pano verilerini ayarlar.
   */
  void setData(const PanoData &data);

  /**
   * Tüm alanları temizler.
   */
  void clear();

  int panoIndex() const { return m_panoIndex; }

signals:
  void dataChanged();
  void deleteRequested();
  void
  panoNameChanged(int panoIndex,
                  const QString &newName); // Tab başlığını güncellemek için

private:
  void setupUi();
  QWidget *setupAnaDagitimPano();
  void setupFonksiyonTestleri();
  QWidget *setupTermalGoruntuler();
  void setupConnections();

  // Iz hesaplama
  void updateIk3();

  // Otomatik sonuç güncelleme (Python: _check_and_update_sonuc)
  void checkAndUpdateSonuc();

  int m_panoIndex;

  // Ana Dağıtım Pano
  QLineEdit *m_panoAdi;
  QComboBox *m_sebekeTipi;
  QLineEdit *m_enerjiSaglayan;
  QLineEdit *m_trafoGucu;
  QSpinBox *m_sistemGerilimi;
  QSpinBox *m_sistemFrekans;
  QLineEdit *m_topraklamaDirenci;
  QComboBox *m_sigortaTipiAna;
  QSpinBox *m_nominalAkimAna;
  QLineEdit *m_rcdBilgisi;
  QLineEdit *m_rcdAnmaAkimi;   // EKLENDI
  QLineEdit *m_rcdTestBilgisi; // EKLENDI
  QLineEdit *m_loopPeN;
  QLineEdit *m_loopLN;
  QLineEdit *m_ik3;

  // Yeni Alanlar
  QLineEdit *m_distCevrimEmpedansi;    // Z_E
  QLineEdit *m_hataAkimi;              // I_f
  QLineEdit *m_sistemTopraklamaKesiti; // Sistem_top
  QLineEdit *m_anaEspotansiyelKesiti;  // Ana_top

  // Empedans Alanları (Python ile aynı)
  QLineEdit *m_zx;  // Zx (Ω)
  QLineEdit *m_re;  // RE (Ω)
  QLineEdit *m_zln; // Zln (Ω)
  QLineEdit *m_ff;  // F-F (V)
  QLineEdit *m_ln;  // L-N (V)
  QLineEdit *m_npe; // N-PE (V)

  QLineEdit *m_ik3Auto; // Ik3 = F-F / Zln

  // Proje ve Şema Durumu
  QCheckBox *m_chkProjeVar;        // Proje Var mı?
  QCheckBox *m_chkTekHatSemasiVar; // Tek Hat Şeması Var mı?

  // Parafudr
  QLineEdit *m_parafudrTip;  // Parafudr Tipi
  QLineEdit *m_parafudrImax; // Parafudr Imax (kA)

  // Uygunluk
  QComboBox *m_uygunluk; // Uygun / Uygun Değil

  // İç sekmeler (Python'daki gibi)
  QTabWidget *m_innerTabs;

  // Gözle Kontrol
  GozleKontrolWidget *m_gozleKontrol;

  // Fonksiyon testleri
  FonksiyonTestleriTable *m_fonksiyonTable;

  // Termal görüntüler
  DragDropWidget *m_termalImages;

  // Zemin izolasyonu
  QLineEdit *m_zeminEn;
  QLineEdit *m_zeminBoy;
  QLineEdit *m_izoDirenci;
  QComboBox *m_izoUygunluk;

  // Potansiyel dengeleme
  QLineEdit *m_enBuyukTopKesit;
};

} // namespace RaporSistemi

#endif // PANOTABWIDGET_H

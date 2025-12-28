/**
 * FonksiyonTestleriTable.h
 *
 * Fonksiyon testleri tablosu.
 * QTableView + QAbstractTableModel kullanır (performans için).
 */

#ifndef FONKSIYONTESTLERITABEL_H
#define FONKSIYONTESTLERITABEL_H

#include <QTableView>
#include <QAbstractTableModel>
#include <QStyledItemDelegate>
#include <QComboBox>
#include <QSpinBox>
#include "DataModels.h"

namespace RaporSistemi {

/**
 * Fonksiyon testleri veri modeli
 */
class FonksiyonTestleriModel : public QAbstractTableModel {
    Q_OBJECT

public:
    // Python'daki tüm sütunlar (20 sütun)
    enum Column {
        ColSiraNo = 0,      // #
        ColLinye,           // Linye Adı
        ColAcmaEgrisi,      // Eğri (B, C, D, K, Z, AAA)
        ColKutupSayisi,     // Kut (1, 2, 3, 4)
        ColIn,              // In (A)
        ColIcu,             // Icu (kA)
        ColIb,              // Ib (In * 0.7)
        ColFazKesiti,       // Faz (mm²)
        ColNotrKesiti,      // N (mm²)
        ColToprakKesiti,    // PE (mm²)
        ColIz,              // Iz (akım taşıma kapasitesi)
        ColSonuc,           // Sonuç (Uygun/Uygun Değil)
        ColKakrVar,         // K (checkbox)
        ColRcdAcma,         // IΔn (30mA/300mA)
        ColRcdMa,           // mA (test değeri)
        ColRcdMs,           // mS (test zamanı)
        ColKakrYok,         // Yok (checkbox)
        ColCount
    };

    explicit FonksiyonTestleriModel(QObject* parent = nullptr);

    // QAbstractTableModel overrides
    int rowCount(const QModelIndex& parent = {}) const override;
    int columnCount(const QModelIndex& parent = {}) const override;
    QVariant data(const QModelIndex& index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex& index, const QVariant& value, int role = Qt::EditRole) override;
    QVariant headerData(int section, Qt::Orientation orientation, int role) const override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;

    // Veri işlemleri
    void addRow();
    void addRows(int count);
    void insertRow(int position);
    void insertRows(int position, int count);
    void insertAnaSigortaRow();  // Python: add_ft_row(is_ana_sigorta=True)
    void removeRow(int position);
    void removeRows(const QVector<int>& positions);
    void clear();

    // Data get/set
    QVector<FonksiyonTesti> getData() const { return m_data; }
    void setData(const QVector<FonksiyonTesti>& data);

signals:
    void modelDataChanged();

private:
    void updateIz(int row);
    void validateSonuc(int row);

    QVector<FonksiyonTesti> m_data;
    QStringList m_sigortaTipleri;
    QStringList m_rcdValues;
    QStringList m_kesitValues;
};

/**
 * ComboBox delegate for table cells
 */
class ComboBoxDelegate : public QStyledItemDelegate {
    Q_OBJECT
public:
    explicit ComboBoxDelegate(const QStringList& items, QObject* parent = nullptr);

    QWidget* createEditor(QWidget* parent, const QStyleOptionViewItem& option,
                         const QModelIndex& index) const override;
    void setEditorData(QWidget* editor, const QModelIndex& index) const override;
    void setModelData(QWidget* editor, QAbstractItemModel* model,
                     const QModelIndex& index) const override;

private:
    QStringList m_items;
};

/**
 * CheckBox delegate for KAKR Yok column
 */
class CheckBoxDelegate : public QStyledItemDelegate {
    Q_OBJECT
public:
    using QStyledItemDelegate::QStyledItemDelegate;

    void paint(QPainter* painter, const QStyleOptionViewItem& option,
              const QModelIndex& index) const override;
    bool editorEvent(QEvent* event, QAbstractItemModel* model,
                    const QStyleOptionViewItem& option, const QModelIndex& index) override;
};

/**
 * Ana tablo widget'ı
 */
class FonksiyonTestleriTable : public QWidget {
    Q_OBJECT

public:
    explicit FonksiyonTestleriTable(QWidget* parent = nullptr);

    QVector<FonksiyonTesti> getData() const;
    void setData(const QVector<FonksiyonTesti>& data);
    void clear();

signals:
    void dataChanged();

public slots:
    void addRow();
    void addMultipleRows();
    void removeSelectedRows();
    void addLinieGroup();
    void autoFillSelectedRcd();
    void copySelectedRows();
    void pasteRows();

private:
    void setupUi();
    void setupContextMenu();

    QTableView* m_tableView;
    FonksiyonTestleriModel* m_model;
    QVector<FonksiyonTesti> m_clipboard;
};

} // namespace RaporSistemi

#endif // FONKSIYONTESTLERITABEL_H

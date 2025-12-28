/**
 * FonksiyonTestleriTable.cpp
 */

#include "FonksiyonTestleriTable.h"
#include "dialogs/LinieGroupDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QPushButton>
#include <QHeaderView>
#include <QMenu>
#include <QInputDialog>
#include <QMessageBox>
#include <QPainter>
#include <QMouseEvent>
#include <QApplication>
#include <QCheckBox>
#include <QRandomGenerator> // Added
#include <algorithm> // Added
#include <random> // Added

namespace RaporSistemi {

// ==================== FonksiyonTestleriModel ====================

FonksiyonTestleriModel::FonksiyonTestleriModel(QObject* parent)
    : QAbstractTableModel(parent)
{
    // Varsayılan değerler
    m_sigortaTipleri = {"B", "C", "D", "Kompakt", "NH"};
    m_rcdValues = {"", "30mA", "100mA", "300mA", "500mA"};
    m_kesitValues = {"1.5", "2.5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120",
                     "2x16", "2x25", "2x35", "3x16", "3x25", "3x35", "4x16", "4x25"};
}

int FonksiyonTestleriModel::rowCount(const QModelIndex& parent) const {
    Q_UNUSED(parent);
    return m_data.size();
}

int FonksiyonTestleriModel::columnCount(const QModelIndex& parent) const {
    Q_UNUSED(parent);
    return ColCount;
}

QVariant FonksiyonTestleriModel::data(const QModelIndex& index, int role) const {
    if (!index.isValid() || index.row() >= m_data.size()) {
        return {};
    }

    const FonksiyonTesti& test = m_data[index.row()];

    if (role == Qt::DisplayRole || role == Qt::EditRole) {
        switch (index.column()) {
            case ColSiraNo: return test.siraNo;
            case ColLinye: return test.linye;
            case ColAcmaEgrisi: return test.sigortaTipi;  // "Eğri" alanı
            case ColKutupSayisi: return test.kutupSayisi;
            case ColIn: return test.nominalAkim;
            case ColIcu: return test.icu;
            case ColIb: return test.ib;
            case ColFazKesiti: return test.fazKesiti;
            case ColNotrKesiti: return test.notrKesiti;
            case ColToprakKesiti: return test.toprakKesiti;
            case ColIz: return test.akimKapasitesi;
            case ColSonuc: return test.sonuc;
            case ColKakrVar: return test.kakrVar;
            case ColRcdAcma: return test.rcd;
            case ColRcdMa: return test.rcdMa;
            case ColRcdMs: return test.rcdMs;
            case ColKakrYok: return test.kakrYok;
        }
    }
    else if (role == Qt::BackgroundRole) {
        // Sonuç hücresini renklendir
        if (index.column() == ColSonuc) {
            if (test.sonuc == "Uygun Değil") {
                return QColor(180, 60, 60);  // Kırmızımsı
            } else if (test.sonuc == "Uygun") {
                return QColor(60, 140, 60);  // Yeşilimsi
            }
        }
    }
    else if (role == Qt::TextAlignmentRole) {
        // Numerik sütunları ortala
        if (index.column() == ColIn || index.column() == ColIz ||
            index.column() == ColSiraNo || index.column() == ColKutupSayisi ||
            index.column() == ColIb || index.column() == ColRcdMa || index.column() == ColRcdMs) {
            return Qt::AlignCenter;
        }
    }

    return {};
}

bool FonksiyonTestleriModel::setData(const QModelIndex& index, const QVariant& value, int role) {
    if (!index.isValid() || index.row() >= m_data.size() || role != Qt::EditRole) {
        return false;
    }

    FonksiyonTesti& test = m_data[index.row()];

    switch (index.column()) {
        case ColLinye: test.linye = value.toString(); break;
        case ColAcmaEgrisi: test.sigortaTipi = value.toString(); break;
        case ColKutupSayisi: test.kutupSayisi = value.toInt(); break;
        case ColIn:
            test.nominalAkim = value.toInt();
            // Ib otomatik hesapla: In * 0.7
            test.ib = QString::number(test.nominalAkim * 0.7, 'f', 1);
            break;
        case ColIcu: test.icu = value.toString(); break;
        case ColFazKesiti:
            test.fazKesiti = value.toString();
            updateIz(index.row());
            break;
        case ColNotrKesiti: test.notrKesiti = value.toString(); break;
        case ColToprakKesiti: test.toprakKesiti = value.toString(); break;
        case ColSonuc: test.sonuc = value.toString(); break;
        case ColKakrVar: test.kakrVar = value.toBool(); break;
        case ColRcdAcma: test.rcd = value.toString(); break;
        case ColRcdMa: test.rcdMa = value.toString(); break;
        case ColRcdMs: test.rcdMs = value.toString(); break;
        case ColKakrYok: test.kakrYok = value.toBool(); break;
        default: return false;
    }

    // Sonuç doğrulama
    validateSonuc(index.row());

    emit dataChanged(index, index);
    emit modelDataChanged();
    return true;
}

QVariant FonksiyonTestleriModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role != Qt::DisplayRole || orientation != Qt::Horizontal) {
        return {};
    }

    // Python sütun başlıkları ile birebir aynı
    static const QStringList headers = {
        "#", "Linye", "Eğri", "Kut", "In", "Icu", "Ib",
        "Faz", "N", "PE", "Iz", "Sonuç", "K", "IΔn", "mA", "mS", "Yok"
    };

    if (section < headers.size()) {
        return headers[section];
    }
    return {};
}

Qt::ItemFlags FonksiyonTestleriModel::flags(const QModelIndex& index) const {
    if (!index.isValid()) {
        return Qt::NoItemFlags;
    }

    Qt::ItemFlags flags = Qt::ItemIsEnabled | Qt::ItemIsSelectable;

    // Otomatik hesaplanan sütunlar düzenlenemez
    if (index.column() != ColIz &&
        index.column() != ColIb &&
        index.column() != ColSiraNo) {
        flags |= Qt::ItemIsEditable;
    }

    return flags;
}

void FonksiyonTestleriModel::addRow() {
    int row = m_data.size();
    beginInsertRows({}, row, row);

    FonksiyonTesti test;
    test.siraNo = row + 1;
    test.sonuc = "Uygun";
    m_data.append(test);

    endInsertRows();
    emit modelDataChanged();
}

void FonksiyonTestleriModel::addRows(int count) {
    if (count <= 0) return;

    int firstRow = m_data.size();
    beginInsertRows({}, firstRow, firstRow + count - 1);

    for (int i = 0; i < count; ++i) {
        FonksiyonTesti test;
        test.siraNo = firstRow + i + 1;
        test.sonuc = "Uygun";
        m_data.append(test);
    }

    endInsertRows();
    emit modelDataChanged();
}

void FonksiyonTestleriModel::insertRow(int position) {
    if (position < 0) position = 0;
    if (position > m_data.size()) position = m_data.size();

    beginInsertRows({}, position, position);

    FonksiyonTesti test;
    test.sonuc = "Uygun";
    m_data.insert(position, test);

    // Sıra numaralarını güncelle
    for (int i = 0; i < m_data.size(); ++i) {
        m_data[i].siraNo = i + 1;
    }

    endInsertRows();
    emit modelDataChanged();
}

void FonksiyonTestleriModel::insertAnaSigortaRow() {
    // Python: multi_pano_gui.py:680-681, 1001-1005
    // Her pano açıldığında ilk satır ANA SİGORTA olarak eklenir
    // is_ana_sigorta=True → 32A KAKR kuralından muaf

    beginInsertRows({}, 0, 0);

    FonksiyonTesti test;
    test.linye = QString::fromUtf8("ANA SİGORTA");  // Python: e.insert(0, "ANA SİGORTA")
    test.sigortaTipi = "C";  // Varsayılan
    test.sonuc = "Uygun";
    test.isAnaSigorta = true;  // 32A KAKR kuralından muaf (Python: entries['_is_ana_sigorta'] = True)

    m_data.insert(0, test);

    // Sıra numaralarını güncelle
    for (int i = 0; i < m_data.size(); ++i) {
        m_data[i].siraNo = i + 1;
    }

    endInsertRows();
    emit modelDataChanged();
}

void FonksiyonTestleriModel::insertRows(int position, int count) {
    if (count <= 0) return;
    if (position < 0) position = 0;
    if (position > m_data.size()) position = m_data.size();

    beginInsertRows({}, position, position + count - 1);

    for (int i = 0; i < count; ++i) {
        FonksiyonTesti test;
        test.sonuc = "Uygun";
        m_data.insert(position + i, test);
    }

    // Sıra numaralarını güncelle
    for (int i = 0; i < m_data.size(); ++i) {
        m_data[i].siraNo = i + 1;
    }

    endInsertRows();
    emit modelDataChanged();
}

void FonksiyonTestleriModel::removeRow(int position) {
    if (position < 0 || position >= m_data.size()) return;

    beginRemoveRows({}, position, position);
    m_data.removeAt(position);
    endRemoveRows();

    // Sıra numaralarını güncelle
    for (int i = 0; i < m_data.size(); ++i) {
        m_data[i].siraNo = i + 1;
    }

    emit modelDataChanged();
}

void FonksiyonTestleriModel::removeRows(const QVector<int>& positions) {
    if (positions.isEmpty()) return;

    // Büyükten küçüğe sırala
    QVector<int> sorted = positions;
    std::sort(sorted.begin(), sorted.end(), std::greater<int>());

    for (int pos : sorted) {
        if (pos >= 0 && pos < m_data.size()) {
            beginRemoveRows({}, pos, pos);
            m_data.removeAt(pos);
            endRemoveRows();
        }
    }

    // Sıra numaralarını güncelle
    for (int i = 0; i < m_data.size(); ++i) {
        m_data[i].siraNo = i + 1;
    }

    emit modelDataChanged();
}

void FonksiyonTestleriModel::clear() {
    if (m_data.isEmpty()) return;

    beginResetModel();
    m_data.clear();
    endResetModel();

    emit modelDataChanged();
}

void FonksiyonTestleriModel::setData(const QVector<FonksiyonTesti>& data) {
    beginResetModel();
    m_data = data;

    // Sıra numaralarını ayarla
    for (int i = 0; i < m_data.size(); ++i) {
        m_data[i].siraNo = i + 1;
    }

    endResetModel();
}

void FonksiyonTestleriModel::updateIz(int row) {
    if (row < 0 || row >= m_data.size()) return;

    FonksiyonTesti& test = m_data[row];
    auto [kesit, carpan] = parseKesit(test.fazKesiti);
    test.akimKapasitesi = kesitToIz(kesit, carpan);

    QModelIndex idx = index(row, ColIz);
    emit QAbstractTableModel::dataChanged(idx, idx);
}

void FonksiyonTestleriModel::validateSonuc(int row) {
    if (row < 0 || row >= m_data.size()) return;

    FonksiyonTesti& test = m_data[row];
    bool uygunDegil = false;
    QString kusur;

    // Python: multi_pano_gui.py:483-487, 4088-4100
    // KAKR grupları için Iz < In kontrolü yapılmaz
    bool isKakrGroup = test.isKakrGroup();

    // 1. In > Iz kontrolü (KAKR grupları için geçerli değil)
    if (!isKakrGroup && test.akimKapasitesi > 0 && test.nominalAkim > test.akimKapasitesi) {
        uygunDegil = true;
        kusur = QString::fromUtf8("Iz(%1A) < In(%2A) - Kablo akım taşıma kapasitesi yetersiz")
            .arg(test.akimKapasitesi).arg(test.nominalAkim);
    }

    // NOT: Python'da Ib > Iz kontrolü YOK, sadece In > Iz var
    // Bu yüzden Ib > Iz kontrolü kaldırıldı


    // 3. 32A ve altı KAKR kontrolü (Python: multi_pano_gui.py:490-500)
    // ÖNEMLİ: ANA SİGORTA için bu kural GEÇERSİZ (Python: if not is_ana_sigorta)
    if (!test.isAnaSigorta && test.nominalAkim > 0 && test.nominalAkim <= 32) {
        // 30mA KAKR var mı kontrol et (Python: has_30ma_kakr = kakr_checked and rcd_acma_val == "30mA")
        bool has30maKakr = test.kakrVar && test.rcd == "30mA";

        if (!has30maKakr && !test.kakrYok) {
            uygunDegil = true;
            if (kusur.isEmpty()) {
                kusur = QString::fromUtf8("In≤32A için 30mA KAKR gerekli");
            }
        }
    }

    // 4. KAKR Yok işaretli ise → Uygun Değil (Python: multi_pano_gui.py:503-504)
    if (test.kakrYok) {
        uygunDegil = true;
        if (kusur.isEmpty()) {
            kusur = QString::fromUtf8("KAKR (kaçak akım rölesi) yok");
        }
    }

    // Sonucu güncelle
    test.sonuc = uygunDegil ? QString::fromUtf8("Uygun Değil") : "Uygun";
    test.aciklama = kusur;

    // UI güncelle
    QModelIndex idxSonuc = index(row, ColSonuc);
    emit QAbstractTableModel::dataChanged(idxSonuc, idxSonuc);
}

// ==================== Delegates ====================

ComboBoxDelegate::ComboBoxDelegate(const QStringList& items, QObject* parent)
    : QStyledItemDelegate(parent)
    , m_items(items)
{}

QWidget* ComboBoxDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem&,
                                        const QModelIndex&) const {
    QComboBox* combo = new QComboBox(parent);
    combo->addItems(m_items);
    combo->setEditable(true);
    return combo;
}

void ComboBoxDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    QComboBox* combo = qobject_cast<QComboBox*>(editor);
    if (combo) {
        QString value = index.data(Qt::EditRole).toString();
        int idx = combo->findText(value);
        if (idx >= 0) {
            combo->setCurrentIndex(idx);
        } else {
            combo->setEditText(value);
        }

        // TAB ile gelince mevcut metin seçili olsun, yeni yazı eskisini değiştirsin
        if (combo->lineEdit()) {
            combo->lineEdit()->selectAll();
        }
    }
}

void ComboBoxDelegate::setModelData(QWidget* editor, QAbstractItemModel* model,
                                    const QModelIndex& index) const {
    QComboBox* combo = qobject_cast<QComboBox*>(editor);
    if (combo) {
        model->setData(index, combo->currentText(), Qt::EditRole);
    }
}

void CheckBoxDelegate::paint(QPainter* painter, const QStyleOptionViewItem& option,
                             const QModelIndex& index) const {
    bool checked = index.data(Qt::EditRole).toBool();

    QStyleOptionButton checkboxOption;
    checkboxOption.rect = option.rect;
    checkboxOption.state = QStyle::State_Enabled;
    if (checked) {
        checkboxOption.state |= QStyle::State_On;
    } else {
        checkboxOption.state |= QStyle::State_Off;
    }

    // Merkeze hizala
    QRect checkRect = QApplication::style()->subElementRect(
        QStyle::SE_CheckBoxIndicator, &checkboxOption);
    checkboxOption.rect.moveCenter(option.rect.center());

    QApplication::style()->drawControl(QStyle::CE_CheckBox, &checkboxOption, painter);
}

bool CheckBoxDelegate::editorEvent(QEvent* event, QAbstractItemModel* model,
                                   const QStyleOptionViewItem&, const QModelIndex& index) {
    if (event->type() == QEvent::MouseButtonRelease) {
        bool checked = index.data(Qt::EditRole).toBool();
        model->setData(index, !checked, Qt::EditRole);
        return true;
    }
    return false;
}

// ==================== FonksiyonTestleriTable ====================

FonksiyonTestleriTable::FonksiyonTestleriTable(QWidget* parent)
    : QWidget(parent)
{
    setupUi();
    setupContextMenu();
}

void FonksiyonTestleriTable::setupUi() {
    QVBoxLayout* layout = new QVBoxLayout(this);
    layout->setContentsMargins(0, 0, 0, 0);

    // Toolbar
    QHBoxLayout* toolbar = new QHBoxLayout();

    // ===== PYTHON GİBİ RENKLİ BUTONLAR =====
    QPushButton* addBtn = new QPushButton(tr("+ Satır"));
    addBtn->setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 6px 12px;");
    connect(addBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::addRow);
    toolbar->addWidget(addBtn);

    QPushButton* groupBtn = new QPushButton(tr("++ Linye Grubu"));
    groupBtn->setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 6px 12px;");
    connect(groupBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::addLinieGroup);
    toolbar->addWidget(groupBtn);

    QPushButton* insertBtn = new QPushButton(tr("↕ Araya Ekle"));
    insertBtn->setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 6px 12px;");
    connect(insertBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::addMultipleRows);
    toolbar->addWidget(insertBtn);

    QPushButton* testBtn = new QPushButton(tr("● DENE"));
    testBtn->setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 6px 12px;");
    connect(testBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::autoFillSelectedRcd);
    toolbar->addWidget(testBtn);

    QPushButton* copyBtn = new QPushButton(tr("📋 Kopyala"));
    copyBtn->setStyleSheet("background-color: #9C27B0; color: white; padding: 6px 12px;");
    connect(copyBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::copySelectedRows);
    toolbar->addWidget(copyBtn);

    QPushButton* pasteBtn = new QPushButton(tr("📥 Yapıştır"));
    pasteBtn->setStyleSheet("background-color: #00BCD4; color: white; padding: 6px 12px;");
    connect(pasteBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::pasteRows);
    toolbar->addWidget(pasteBtn);

    QPushButton* removeBtn = new QPushButton(tr("🗑️ Seçileni Sil"));
    removeBtn->setStyleSheet("background-color: #E91E63; color: white; font-weight: bold; padding: 6px 12px;");
    connect(removeBtn, &QPushButton::clicked, this, &FonksiyonTestleriTable::removeSelectedRows);
    toolbar->addWidget(removeBtn);

    toolbar->addStretch();
    layout->addLayout(toolbar);

    // Tablo
    m_tableView = new QTableView();
    m_model = new FonksiyonTestleriModel(this);
    m_tableView->setModel(m_model);

    // Delegate'ler - Python sütunları ile aynı
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColAcmaEgrisi,
        new ComboBoxDelegate({"B", "C", "D", "K", "Z", "AAA"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColKutupSayisi,
        new ComboBoxDelegate({"1", "2", "3", "4"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColIn,
        new ComboBoxDelegate({"6", "10", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125", "160", "200", "250", "315", "400"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColIcu,
        new ComboBoxDelegate({"3kA", "4.5kA", "6kA", "10kA", "25kA", "35kA", "55kA", "70kA"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColFazKesiti,
        new ComboBoxDelegate({"1.5", "2.5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120", "150", "185", "240", "2x16", "2x25", "2x35", "3x70", "3x95"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColNotrKesiti,
        new ComboBoxDelegate({"1.5", "2.5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120", "150"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColToprakKesiti,
        new ComboBoxDelegate({"1.5", "2.5", "4", "6", "10", "16", "25", "35", "50"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColSonuc,
        new ComboBoxDelegate({"Uygun", "Uygun Değil"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColKakrVar,
        new CheckBoxDelegate(this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColRcdAcma,
        new ComboBoxDelegate({"", "30mA", "100mA", "300mA", "500mA"}, this));
    m_tableView->setItemDelegateForColumn(FonksiyonTestleriModel::ColKakrYok,
        new CheckBoxDelegate(this));

    // Görünüm ayarları
    m_tableView->setSelectionBehavior(QAbstractItemView::SelectItems);  // Hücre seçimi
    m_tableView->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_tableView->setAlternatingRowColors(true);
    m_tableView->verticalHeader()->setDefaultSectionSize(28);
    m_tableView->setFocusPolicy(Qt::StrongFocus);

    // Aktif hücre vurgulama stili - TAB ile geçişte görünür
    m_tableView->setStyleSheet(
        "QTableView {"
        "    gridline-color: #ccc;"
        "    selection-background-color: #d0e8ff;"  // Seçili satır açık mavi
        "    selection-color: #000;"
        "}"
        "QTableView::item:selected {"
        "    background-color: #d0e8ff;"            // Seçili hücre
        "    color: #000;"
        "}"
        "QTableView::item:focus {"
        "    background-color: #fff3cd;"            // Aktif hücre sarımsı
        "    border: 2px solid #ff9800;"            // Turuncu kenarlık
        "    color: #000;"
        "}"
        "QTableView::item:selected:focus {"
        "    background-color: #ffe082;"            // Seçili + aktif
        "    border: 2px solid #ff5722;"            // Koyu turuncu kenarlık
        "    color: #000;"
        "}"
    );

    // Sütun genişlikleri - SABİT MİNİMUM GENİŞLİKLER
    QHeaderView* header = m_tableView->horizontalHeader();
    header->setMinimumSectionSize(30);

    // Interactive mode - kullanıcı sütun genişliğini değiştirebilir
    header->setSectionResizeMode(QHeaderView::Interactive);

    // Linye sütunu genişleyebilir (Stretch)
    header->setSectionResizeMode(FonksiyonTestleriModel::ColLinye, QHeaderView::Stretch);

    // Son sütun genişlesin (Yok checkbox)
    header->setStretchLastSection(true);

    // SABİT SÜTUN GENİŞLİKLERİ - Her sütun görünür olacak
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColSiraNo, 30);      // #
    // Linye Stretch olduğu için genişlik vermiyoruz
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColAcmaEgrisi, 45);  // Eğri
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColKutupSayisi, 40); // Kut
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColIn, 40);          // In
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColIcu, 55);         // Icu
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColIb, 45);          // Ib
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColFazKesiti, 50);   // Faz
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColNotrKesiti, 40);  // N
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColToprakKesiti, 40);// PE
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColIz, 40);          // Iz
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColSonuc, 75);       // Sonuç
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColKakrVar, 30);     // K (checkbox)
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColRcdAcma, 55);     // IΔn
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColRcdMa, 45);       // mA
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColRcdMs, 45);       // mS
    m_tableView->setColumnWidth(FonksiyonTestleriModel::ColKakrYok, 40);     // Yok

    connect(m_model, &FonksiyonTestleriModel::modelDataChanged,
            this, &FonksiyonTestleriTable::dataChanged);

    layout->addWidget(m_tableView, 1);

    // Python: multi_pano_gui.py:680-681
    // Başlangıçta ilk satır ANA SİGORTA olarak eklenir
    m_model->insertAnaSigortaRow();
}

void FonksiyonTestleriTable::setupContextMenu() {
    m_tableView->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_tableView, &QTableView::customContextMenuRequested, [this](const QPoint& pos) {
        QMenu menu;
        menu.addAction(tr("Satır Ekle"), this, &FonksiyonTestleriTable::addRow);
        menu.addAction(tr("Seçilenleri Sil"), this, &FonksiyonTestleriTable::removeSelectedRows);
        menu.addSeparator();
        menu.addAction(tr("Linye Grubu Ekle"), this, &FonksiyonTestleriTable::addLinieGroup);
        menu.exec(m_tableView->viewport()->mapToGlobal(pos));
    });
}

void FonksiyonTestleriTable::addRow() {
    m_model->addRow();
}

void FonksiyonTestleriTable::addMultipleRows() {
    bool ok;
    int count = QInputDialog::getInt(this, tr("Satır Ekle"),
        tr("Eklenecek satır sayısı:"), 5, 1, 100, 1, &ok);

    if (ok) {
        QModelIndexList selected = m_tableView->selectionModel()->selectedRows();
        if (selected.isEmpty()) {
            m_model->addRows(count);
        } else {
            // Seçilen son satırın altına ekle
            int maxRow = -1;
            for (const QModelIndex& idx : selected) {
                if (idx.row() > maxRow) maxRow = idx.row();
            }
            m_model->insertRows(maxRow + 1, count);
        }
    }
}

void FonksiyonTestleriTable::removeSelectedRows() {
    QModelIndexList selected = m_tableView->selectionModel()->selectedRows();
    if (selected.isEmpty()) return;

    QVector<int> rows;
    for (const QModelIndex& idx : selected) {
        rows.append(idx.row());
    }

    m_model->removeRows(rows);
}

void FonksiyonTestleriTable::autoFillSelectedRcd() {
    QModelIndexList selected = m_tableView->selectionModel()->selectedRows();
    if (selected.isEmpty()) {
        QMessageBox::warning(this, tr("Uyarı"), tr("Önce satır seçmelisiniz."));
        return;
    }

    for (const QModelIndex& idx : selected) {
        int row = idx.row();
        QString rcdType = m_model->data(m_model->index(row, FonksiyonTestleriModel::ColRcdAcma), Qt::EditRole).toString();

        if (rcdType == "30mA") {
            int ma = QRandomGenerator::global()->bounded(20, 30); // 20-29
            double ms = QRandomGenerator::global()->bounded(200, 401) / 10.0; // 20.0 - 40.0
            m_model->setData(m_model->index(row, FonksiyonTestleriModel::ColRcdMa), ma);
            m_model->setData(m_model->index(row, FonksiyonTestleriModel::ColRcdMs), QString::number(ms, 'f', 1));
        } else if (rcdType == "300mA") {
            int ma = (QRandomGenerator::global()->bounded(250, 291) / 10) * 10; // 250, 260, 270, 280, 290
            double ms = QRandomGenerator::global()->bounded(200, 401) / 10.0; // 20.0 - 40.0
            m_model->setData(m_model->index(row, FonksiyonTestleriModel::ColRcdMa), ma);
            m_model->setData(m_model->index(row, FonksiyonTestleriModel::ColRcdMs), QString::number(ms, 'f', 1));
        }
    }
}

void FonksiyonTestleriTable::copySelectedRows() {
    QModelIndexList selected = m_tableView->selectionModel()->selectedRows();
    if (selected.isEmpty()) {
        QMessageBox::warning(this, tr("Uyarı"), tr("Önce satır seçmelisiniz."));
        return;
    }

    m_clipboard.clear();
    // Satır indexlerine göre sırala
    std::sort(selected.begin(), selected.end(), [](const QModelIndex& a, const QModelIndex& b) {
        return a.row() < b.row();
    });

    QVector<FonksiyonTesti> currentData = m_model->getData();
    for (const QModelIndex& idx : selected) {
        if (idx.row() < currentData.size()) {
            m_clipboard.append(currentData[idx.row()]);
        }
    }

    QMessageBox::information(this, tr("Kopyalandı"), tr("%1 satır kopyalandı.").arg(m_clipboard.size()));
}

void FonksiyonTestleriTable::pasteRows() {
    if (m_clipboard.isEmpty()) {
        QMessageBox::warning(this, tr("Uyarı"), tr("Yapıştırılacak veri yok. Önce kopyalayın."));
        return;
    }

    int startRow = m_model->rowCount();
    m_model->addRows(m_clipboard.size());

    QVector<FonksiyonTesti> currentData = m_model->getData();
    for (int i = 0; i < m_clipboard.size(); ++i) {
        int row = startRow + i;
        FonksiyonTesti pasted = m_clipboard[i];
        pasted.siraNo = row + 1; // Sıra numarasını güncelle
        currentData[row] = pasted;
    }
    m_model->setData(currentData);
}

void FonksiyonTestleriTable::addLinieGroup() {
    LinieGroupDialog dialog(this);
    if (dialog.exec() == QDialog::Accepted) {
        QVector<FonksiyonTesti> tests = dialog.getTests();
        if (tests.isEmpty()) return;

        int startRow = m_model->rowCount();
        m_model->addRows(tests.size());

        QVector<FonksiyonTesti> currentData = m_model->getData();
        for (int i = 0; i < tests.size(); ++i) {
            int row = startRow + i;
            FonksiyonTesti test = tests[i];
            test.siraNo = row + 1;
            currentData[row] = test;
        }
        m_model->setData(currentData);

        emit dataChanged();
    }
}

QVector<FonksiyonTesti> FonksiyonTestleriTable::getData() const {
    return m_model->getData();
}

void FonksiyonTestleriTable::setData(const QVector<FonksiyonTesti>& data) {
    m_model->setData(data);
}

void FonksiyonTestleriTable::clear() {
    m_model->clear();
}

} // namespace RaporSistemi

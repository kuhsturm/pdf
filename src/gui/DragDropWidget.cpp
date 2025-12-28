/**
 * DragDropWidget.cpp
 */

#include "DragDropWidget.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QFileInfo>
#include <QFileDialog>
#include <QPainter>
#include <QPushButton>
#include <QUrl>

namespace RaporSistemi {

DragDropWidget::DragDropWidget(QWidget* parent)
    : QWidget(parent)
{
    setupUi();
    setAcceptDrops(true);
}

void DragDropWidget::setupUi() {
    QVBoxLayout* layout = new QVBoxLayout(this);

    m_placeholderLabel = new QLabel(tr("Dosyaları buraya sürükleyin\nveya çift tıklayarak seçin"));
    m_placeholderLabel->setAlignment(Qt::AlignCenter);
    m_placeholderLabel->setStyleSheet("color: #888; font-size: 12px;");
    layout->addWidget(m_placeholderLabel);

    m_fileList = new QListWidget();
    m_fileList->setVisible(false);
    m_fileList->setMaximumHeight(100);
    layout->addWidget(m_fileList);

    setMinimumHeight(80);
    setStyleSheet("DragDropWidget { "
                  "border: 2px dashed #555; "
                  "border-radius: 8px; "
                  "background-color: rgba(50, 50, 50, 0.5); "
                  "}");
}

void DragDropWidget::setAcceptedExtensions(const QStringList& extensions) {
    m_acceptedExtensions = extensions;
}

void DragDropWidget::setPlaceholderText(const QString& text) {
    m_placeholderLabel->setText(text);
}

bool DragDropWidget::isAcceptedFile(const QString& path) const {
    if (m_acceptedExtensions.isEmpty()) return true;

    QFileInfo info(path);
    QString ext = info.suffix().toLower();

    for (const QString& accepted : m_acceptedExtensions) {
        if (ext == accepted.toLower()) {
            return true;
        }
    }
    return false;
}

void DragDropWidget::dragEnterEvent(QDragEnterEvent* event) {
    if (event->mimeData()->hasUrls()) {
        m_isDragging = true;
        event->acceptProposedAction();
        update();
    }
}

void DragDropWidget::dragLeaveEvent(QDragLeaveEvent* event) {
    Q_UNUSED(event);
    m_isDragging = false;
    update();
}

void DragDropWidget::dropEvent(QDropEvent* event) {
    m_isDragging = false;

    QStringList newFiles;
    for (const QUrl& url : event->mimeData()->urls()) {
        QString path = url.toLocalFile();
        if (!path.isEmpty() && isAcceptedFile(path)) {
            newFiles.append(path);
        }
    }

    if (!newFiles.isEmpty()) {
        addFiles(newFiles);
        event->acceptProposedAction();
    }

    update();
}

void DragDropWidget::paintEvent(QPaintEvent* event) {
    QWidget::paintEvent(event);

    if (m_isDragging) {
        QPainter painter(this);
        painter.fillRect(rect(), QColor(42, 130, 218, 50));
        painter.setPen(QPen(QColor(42, 130, 218), 3));
        painter.drawRect(rect().adjusted(2, 2, -2, -2));
    }
}

void DragDropWidget::mouseDoubleClickEvent(QMouseEvent* event) {
    Q_UNUSED(event);

    QString filter;
    if (!m_acceptedExtensions.isEmpty()) {
        QStringList patterns;
        for (const QString& ext : m_acceptedExtensions) {
            patterns.append("*." + ext);
        }
        filter = tr("Dosyalar (%1)").arg(patterns.join(" "));
    }

    QStringList files = QFileDialog::getOpenFileNames(this,
        tr("Dosya Seç"), QString(), filter);

    if (!files.isEmpty()) {
        addFiles(files);
    }
}

void DragDropWidget::setFiles(const QStringList& files) {
    m_files = files;
    updateDisplay();
    emit filesChanged();
}

void DragDropWidget::addFiles(const QStringList& files) {
    for (const QString& file : files) {
        if (!m_files.contains(file)) {
            m_files.append(file);
            emit fileAdded(file);
        }
    }
    updateDisplay();
    emit filesChanged();
}

void DragDropWidget::clear() {
    m_files.clear();
    updateDisplay();
    emit filesChanged();
}

void DragDropWidget::updateDisplay() {
    m_fileList->clear();

    if (m_files.isEmpty()) {
        m_placeholderLabel->setVisible(true);
        m_fileList->setVisible(false);
    } else {
        m_placeholderLabel->setVisible(false);
        m_fileList->setVisible(true);

        for (int i = 0; i < m_files.size(); ++i) {
            QFileInfo info(m_files[i]);

            // Her dosya için özel widget oluştur
            QWidget* itemWidget = new QWidget();
            QHBoxLayout* layout = new QHBoxLayout(itemWidget);
            layout->setContentsMargins(4, 2, 4, 2);
            layout->setSpacing(4);

            // Dosya adı
            QLabel* nameLabel = new QLabel(info.fileName());
            nameLabel->setStyleSheet("color: #4CAF50;");
            layout->addWidget(nameLabel, 1);

            // X (Sil) butonu
            QPushButton* deleteBtn = new QPushButton("✖");
            deleteBtn->setFixedSize(20, 20);
            deleteBtn->setStyleSheet(
                "QPushButton { background-color: #8B0000; color: white; "
                "border: none; border-radius: 3px; font-weight: bold; }"
                "QPushButton:hover { background-color: #B22222; }");
            deleteBtn->setToolTip(tr("Dosyayı listeden kaldır"));

            // Closure ile index'i yakala
            int fileIndex = i;
            connect(deleteBtn, &QPushButton::clicked, this, [this, fileIndex]() {
                removeFile(fileIndex);
            });
            layout->addWidget(deleteBtn);

            // Liste öğesi ekle
            QListWidgetItem* item = new QListWidgetItem();
            item->setSizeHint(itemWidget->sizeHint());
            m_fileList->addItem(item);
            m_fileList->setItemWidget(item, itemWidget);
        }
    }
}

void DragDropWidget::removeFile(int index) {
    if (index >= 0 && index < m_files.size()) {
        m_files.removeAt(index);
        updateDisplay();
        emit filesChanged();
    }
}

} // namespace RaporSistemi

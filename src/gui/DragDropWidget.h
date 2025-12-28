/**
 * DragDropWidget.h
 *
 * Dosya sürükle-bırak widget'ı.
 */

#ifndef DRAGDROPWIDGET_H
#define DRAGDROPWIDGET_H

#include <QWidget>
#include <QLabel>
#include <QListWidget>
#include <QStringList>

namespace RaporSistemi {

class DragDropWidget : public QWidget {
    Q_OBJECT

public:
    explicit DragDropWidget(QWidget* parent = nullptr);

    void setAcceptedExtensions(const QStringList& extensions);
    void setPlaceholderText(const QString& text);

    QStringList getFiles() const { return m_files; }
    void setFiles(const QStringList& files);
    void addFiles(const QStringList& files);
    void removeFile(int index);  // Tek dosya silme
    void clear();

signals:
    void filesChanged();
    void fileAdded(const QString& path);

protected:
    void dragEnterEvent(QDragEnterEvent* event) override;
    void dragLeaveEvent(QDragLeaveEvent* event) override;
    void dropEvent(QDropEvent* event) override;
    void paintEvent(QPaintEvent* event) override;
    void mouseDoubleClickEvent(QMouseEvent* event) override;

private:
    void setupUi();
    void updateDisplay();
    bool isAcceptedFile(const QString& path) const;

    QStringList m_files;
    QStringList m_acceptedExtensions;
    QLabel* m_placeholderLabel;
    QListWidget* m_fileList;
    bool m_isDragging = false;
};

} // namespace RaporSistemi

#endif // DRAGDROPWIDGET_H

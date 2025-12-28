/**
 * Elektrik Rapor Sistemi - C++ Edition
 *
 * Ana giriş noktası.
 * Qt6 ve modern C++17 ile yazılmıştır.
 */

#include <QApplication>
#include <QTranslator>
#include <QLocale>
#include <QStyleFactory>
#include <QFont>
#include <QFile>
#include "gui/MainWindow.h"

int main(int argc, char *argv[])
{
    // High DPI desteği
    QApplication::setHighDpiScaleFactorRoundingPolicy(
        Qt::HighDpiScaleFactorRoundingPolicy::PassThrough);

    QApplication app(argc, argv);

    // Uygulama bilgileri
    app.setApplicationName("Elektrik Rapor Sistemi");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("Rapor Sistemi");

    // Türkçe locale ayarla
    QLocale::setDefault(QLocale(QLocale::Turkish, QLocale::Turkey));

    // Qt Türkçe çevirileri yükle
    QTranslator translator;
    if (translator.load(QLocale(), "qt", "_", ":/translations")) {
        app.installTranslator(&translator);
    }

    // Fusion style (modern görünüm)
    app.setStyle(QStyleFactory::create("Fusion"));

    // Açık tema (göz yormayan kirli beyaz)
    QPalette lightPalette;
    lightPalette.setColor(QPalette::Window, QColor(245, 245, 245));        // Kirli beyaz arka plan
    lightPalette.setColor(QPalette::WindowText, QColor(50, 50, 50));       // Koyu gri metin
    lightPalette.setColor(QPalette::Base, QColor(255, 255, 255));          // Beyaz input alanları
    lightPalette.setColor(QPalette::AlternateBase, QColor(240, 240, 240)); // Alternatif satır rengi
    lightPalette.setColor(QPalette::ToolTipBase, QColor(255, 255, 225));   // Sarımsı tooltip
    lightPalette.setColor(QPalette::ToolTipText, QColor(50, 50, 50));
    lightPalette.setColor(QPalette::Text, QColor(50, 50, 50));             // Koyu gri metin
    lightPalette.setColor(QPalette::Button, QColor(240, 240, 240));        // Buton arka planı
    lightPalette.setColor(QPalette::ButtonText, QColor(50, 50, 50));
    lightPalette.setColor(QPalette::BrightText, Qt::red);
    lightPalette.setColor(QPalette::Link, QColor(0, 100, 180));            // Mavi link
    lightPalette.setColor(QPalette::Highlight, QColor(42, 130, 218));      // Seçim rengi
    lightPalette.setColor(QPalette::HighlightedText, Qt::white);
    lightPalette.setColor(QPalette::Disabled, QPalette::Text, QColor(160, 160, 160));
    lightPalette.setColor(QPalette::Disabled, QPalette::ButtonText, QColor(160, 160, 160));
    app.setPalette(lightPalette);

    // Varsayılan font
    QFont defaultFont("Segoe UI", 10);
    app.setFont(defaultFont);

    // Ana pencereyi oluştur ve göster
    RaporSistemi::MainWindow mainWindow;
    mainWindow.show();

    return app.exec();
}

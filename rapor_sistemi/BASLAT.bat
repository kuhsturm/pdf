@echo off
chcp 65001 >nul
title Elektrik Rapor Sistemi v2.0

echo ═══════════════════════════════════════════════════════════════════════
echo          Elektrik Tesisatı Rapor Sistemi v2.0
echo          Şablon Dosyası EXE İçine Gömülüdür
echo ═══════════════════════════════════════════════════════════════════════
echo.

cd /d "%~dp0"

:: Önce dist_new kontrol et (yeni derleme)
if exist "dist_new\ElektrikRaporSistemi.exe" (
    echo [*] Derlenmiş uygulama başlatılıyor (dist_new)...
    start "" "dist_new\ElektrikRaporSistemi.exe"
    exit
)

:: Sonra dist kontrol et
if exist "dist\ElektrikRaporSistemi.exe" (
    echo [*] Derlenmiş uygulama başlatılıyor (dist)...
    start "" "dist\ElektrikRaporSistemi.exe"
    exit
)

:: Python ile başlat
echo [*] Python ile başlatılıyor...

:: Önce sanal ortamı dene
if exist "..\.venv\Scripts\python.exe" (
    "..\.venv\Scripts\python.exe" multi_pano_gui.py
) else if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" multi_pano_gui.py
) else (
    python multi_pano_gui.py
)

if errorlevel 1 (
    echo.
    echo [!] Hata oluştu. Gerekli paketler yüklü olmayabilir.
    echo     Paketleri yüklemek için: pip install customtkinter python-docx openpyxl pypdf pillow
    pause
)

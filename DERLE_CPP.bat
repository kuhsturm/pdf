@echo off
chcp 65001 >nul
echo ═══════════════════════════════════════════════════════════════════════
echo         Elektrik Tesisatı Rapor Sistemi - Derleme Scripti (C++)
echo ═══════════════════════════════════════════════════════════════════════
echo.

cd /d "%~dp0"

:: PATH ayarları
set PATH=C:\Program Files\CMake\bin;C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.6.0\mingw_64\bin;%PATH%

:: Build dizini oluştur
if not exist "build" mkdir build

:: 1. CMake Konfigürasyonu
echo [*] CMake konfigürasyonu yapılıyor...
cmake -S . -B build -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Release
if errorlevel 1 goto :Error

:: 2. Derleme
echo.
echo [*] Derleniyor...
cmake --build build --config Release --parallel 8
if errorlevel 1 goto :Error

:: 3. Kontrol
if not exist "build\ElektrikRaporSistemi.exe" goto :Error

echo.
echo ═══════════════════════════════════════════════════════════════════════
echo [✓] DERLEME BAŞARILI!
echo ═══════════════════════════════════════════════════════════════════════

:: 4. Deployment (Qt DLL'lerini kopyala)
echo.
echo [*] Qt bağımlılıkları kopyalanıyor (windeployqt)...
windeployqt --no-translations --compiler-runtime "build\ElektrikRaporSistemi.exe"
if errorlevel 1 (
    echo [!] Windeployqt hatası! DLL'ler kopyalanamamış olabilir.
)

echo [*] Sablon dosyalari kopyalaniyor...
if not exist "build\sablon" mkdir "build\sablon"
copy /Y "src\..\sablon\rapor_sablonu.docx" "build\sablon\rapor_sablonu.docx"

:: 5. Ekstra Dosyalar
if exist "kisi_bilgileri.xlsx" (
    copy "kisi_bilgileri.xlsx" "build\" >nul
    echo [i] Kişi bilgileri dosyası kopyalandı.
)

:: 6. Başlat
echo.
echo [i] Uygulama başlatılıyor...
start "" "build\ElektrikRaporSistemi.exe"

goto :EOF

:Error
echo.
echo ═══════════════════════════════════════════════════════════════════════
echo [✗] DERLEME HATASI!
echo    CMake, derleyici hatası veya dosya bulunamadı.
echo ═══════════════════════════════════════════════════════════════════════
exit /b 1

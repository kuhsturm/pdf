@echo off
chcp 65001 >nul
echo ═══════════════════════════════════════════════════════════════════════
echo         Elektrik Tesisatı Rapor Sistemi - Derleme Scripti
echo ═══════════════════════════════════════════════════════════════════════
echo.

cd /d "%~dp0"

:: Python yolunu kontrol et
set VENV_PATH=..\.venv\Scripts
if exist "%VENV_PATH%\python.exe" (
    echo [✓] Sanal ortam bulundu: %VENV_PATH%
) else (
    echo [!] Sanal ortam bulunamadı, sistem Python kullanılacak...
    set VENV_PATH=
)

:: Önceki build'leri temizle
echo.
echo [*] Önceki derleme dosyaları temizleniyor...
if exist "dist" rd /s /q "dist"
if exist "build" rd /s /q "build"

:: PyInstaller ile derle
echo.
echo [*] PyInstaller ile derleniyor...
echo.

if defined VENV_PATH (
    "%VENV_PATH%\pyinstaller.exe" RaporSistemi.spec --clean
) else (
    pyinstaller RaporSistemi.spec --clean
)

echo.
if exist "dist\ElektrikRaporSistemi.exe" (
    echo ═══════════════════════════════════════════════════════════════════════
    echo [✓] DERLEME BAŞARILI!
    echo.
    echo    Çalıştırılabilir dosya: dist\ElektrikRaporSistemi.exe
    echo ═══════════════════════════════════════════════════════════════════════

    :: İsteğe bağlı: config dosyasını kopyala
    if not exist "dist\config" mkdir "dist\config"
    if exist "..\config\system_config.json" (
        copy "..\config\system_config.json" "dist\config\" >nul
        echo [i] Config dosyası kopyalandı.
    )

    :: Kişi bilgileri dosyasını kopyala (Kullanıcı düzenleyebilsin diye)
    if exist "kisi_bilgileri.xlsx" (
        copy "kisi_bilgileri.xlsx" "dist\" >nul
        echo [i] Kişi bilgileri dosyası kopyalandı.
    )

) else (
    echo ═══════════════════════════════════════════════════════════════════════
    echo [✗] DERLEME BAŞARISIZ!
    echo    Lütfen hata mesajlarını kontrol edin.
    echo ═══════════════════════════════════════════════════════════════════════
)

echo.

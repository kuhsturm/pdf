@echo off
setlocal
chcp 65001 >nul
echo ═══════════════════════════════════════════════════════════════════════
echo         SAFE BUILD SCRIPT (SUBST Fix)
echo ═══════════════════════════════════════════════════════════════════════
echo.

set "BUILD_DIR=C:\Users\cmshe\.gemini\antigravity\brain\build_cpp"
set "SOURCE_DRIVE=Z:"

:: Get absolute path without trailing backslash issues
pushd "%~dp0"
set "ABS_PATH=%CD%"
popd

:: Clean previous drive mapping if exists
subst %SOURCE_DRIVE% /d >nul 2>&1

:: Map source to drive
echo Mapping %SOURCE_DRIVE% to %ABS_PATH%
subst %SOURCE_DRIVE% "%ABS_PATH%"
if errorlevel 1 (
    echo [!] Surucu haritalama hatasi. Z surucusu kullanimda olabilir.
    exit /b 1
)

:: Clean build dir
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
mkdir "%BUILD_DIR%"

:: Switch to mapped drive for building
%SOURCE_DRIVE%
cd \

:: PATH settings
set PATH=C:\Program Files\CMake\bin;C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.6.0\mingw_64\bin;%PATH%

echo [*] CMake Config...
cmake -S . -B "%BUILD_DIR%" -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Release
if errorlevel 1 goto :Error

echo.
echo [*] Building...
cmake --build "%BUILD_DIR%" --config Release --parallel 8
if errorlevel 1 goto :Error

echo.
echo [✓] SUCCESS!

:: Deploy
echo [*] Deploying Qt files...
windeployqt --no-translations --compiler-runtime "%BUILD_DIR%\ElektrikRaporSistemi.exe"

echo [*] Copying templates...
if not exist "%BUILD_DIR%\sablon" mkdir "%BUILD_DIR%\sablon"
copy /Y "sablon\rapor_sablonu.docx" "%BUILD_DIR%\sablon\rapor_sablonu.docx"

echo [*] Copying Python scripts...
copy /Y "src\resources\sozlesme_parser.py" "%BUILD_DIR%\sozlesme_parser.py"

if exist "kisi_bilgileri.xlsx" copy "kisi_bilgileri.xlsx" "%BUILD_DIR%\" >nul

:: Unmount drive
cd /d C:\
subst %SOURCE_DRIVE% /d

echo.
echo [i] Launching...
start "" "%BUILD_DIR%\ElektrikRaporSistemi.exe"
goto :EOF

:Error
cd /d C:\
subst %SOURCE_DRIVE% /d
echo [✗] BUILD FAILED
exit /b 1

@echo off
chcp 65001 >nul
title ConvertFlow - Kurulum Sihirbazi
cd /d "%~dp0"

echo ======================================================
echo    ConvertFlow - Tek Tikla Kurulum (Windows)
echo ======================================================
echo.

where node >nul 2>&1
if errorlevel 1 (
  echo [!] Node.js bulunamadi. Winget ile otomatik kuruluyor...
  winget install -e --id OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
  echo [*] Lutfen kurulumdan sonra bu dosyayi tekrar calistirin.
  pause
  exit /b 1
)

where python >nul 2>&1
if errorlevel 1 (
  echo [!] Python bulunamadi. Winget ile otomatik kuruluyor...
  winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
  echo [*] Lutfen kurulumdan sonra bu dosyayi tekrar calistirin.
  pause
  exit /b 1
)

echo [1/2] Node.js paketleri kuruluyor...
call npm install --omit=dev
if errorlevel 1 (
  echo [HATA] npm install basarisiz oldu.
  pause
  exit /b 1
)

echo.
echo [2/2] Python PDF motoru ve Sag Tik menusu kuruluyor...
node context-menu/setup.js
if errorlevel 1 (
  echo [HATA] Kurulum tamamlanamadi.
  pause
  exit /b 1
)

echo.
set /p WIN11_CHOICE="Windows 11'de Shift'e basmadan ilk sag tikta gorunsun mu? (E/H): "
if /I "%WIN11_CHOICE%"=="E" (
  node context-menu/win11-classic-menu.js enable
)

echo.
echo ======================================================
echo    KURULUM BASARIYLA TAMAMLANDI!
echo ======================================================
pause


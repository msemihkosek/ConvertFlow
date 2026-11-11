@echo off
chcp 65001 >nul
title ConvertFlow - Kaldirma
cd /d "%~dp0"

echo ConvertFlow sag tik menusu kaldiriliyor...
node context-menu/uninstall-menu.js

echo.
set /p WIN11_RESET="Windows 11 modern sag tik menusune geri donulsun mu? (E/H): "
if /I "%WIN11_RESET%"=="E" (
  node context-menu/win11-classic-menu.js disable
)

echo.
echo Islem tamamlandi.
pause


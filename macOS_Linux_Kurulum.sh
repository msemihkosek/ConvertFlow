#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

echo "======================================================"
echo "   ConvertFlow - Tek Komutla Kurulum (macOS / Linux)"
echo "======================================================"
echo ""

OS="$(uname -s)"

if ! command -v node >/dev/null 2>&1; then
  echo "[!] Node.js bulunamadı."
  if [ "$OS" = "Darwin" ] && command -v brew >/dev/null 2>&1; then
    echo "Homebrew ile Node.js kuruluyor..."
    brew install node
  elif command -v apt-get >/dev/null 2>&1; then
    echo "Kurmak için: sudo apt update && sudo apt install -y nodejs npm"
    exit 1
  else
    echo "Lütfen Node.js (v18+) kurun: https://nodejs.org/"
    exit 1
  fi
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "[!] Python 3 bulunamadı."
  if [ "$OS" = "Darwin" ] && command -v brew >/dev/null 2>&1; then
    echo "Homebrew ile Python kuruluyor..."
    brew install python
  elif command -v apt-get >/dev/null 2>&1; then
    echo "Kurmak için: sudo apt update && sudo apt install -y python3 python3-venv python3-pip"
    exit 1
  else
    echo "Lütfen Python 3.10+ kurun: https://www.python.org/"
    exit 1
  fi
fi

if [ "$OS" = "Linux" ] && ! command -v zenity >/dev/null 2>&1 && ! command -v kdialog >/dev/null 2>&1; then
  echo "💡 İpucu: GNOME/Nemo sağ tık menüsü için 'zenity' önerilir (sudo apt install -y zenity libnotify-bin)"
fi

echo "[1/2] Node.js paketleri kuruluyor..."
npm install --omit=dev

echo ""
echo "[2/2] Python PDF motoru ve Sağ Tık menüsü kuruluyor..."
node context-menu/setup.js

echo ""
echo "✅ Kurulum tamamlandı!"


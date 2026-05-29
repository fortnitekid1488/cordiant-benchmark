#!/bin/zsh
set -e
cd "$(dirname "$0")"
source scripts/unix_python.sh
PACKAGE_PATH=$("$PYTHON_EXE" scripts/prepare_aistudio_sources.py --mode quarterly --download-mode full)
rm -f OPEN_THIS_QUARTERLY_PACKAGE
ln -s "$PACKAGE_PATH" OPEN_THIS_QUARTERLY_PACKAGE
echo "$PACKAGE_PATH"
echo ""
echo "Готово. Это квартальный пакет для AI Studio. Открывай папку OPEN_THIS_QUARTERLY_PACKAGE."
read -k 1 "?Нажми любую клавишу для выхода..."

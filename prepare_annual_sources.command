#!/bin/zsh
set -e
cd "$(dirname "$0")"
source scripts/unix_python.sh
PACKAGE_PATH=$("$PYTHON_EXE" scripts/prepare_aistudio_sources.py --mode annual --download-mode full)
rm -f OPEN_THIS_ANNUAL_PACKAGE OPEN_THIS_ANNUAL_TRAINING_PACKAGE
ln -s "$PACKAGE_PATH" OPEN_THIS_ANNUAL_PACKAGE
echo "$PACKAGE_PATH"
echo ""
echo "Готово. Это годовой пакет для последнего полного года. Открывай папку OPEN_THIS_ANNUAL_PACKAGE."
read -k 1 "?Нажми любую клавишу для выхода..."

#!/bin/zsh
set -e
cd "$(dirname "$0")/../../.."
source scripts/unix_python.sh
PACKAGE_PATH=$("$PYTHON_EXE" scripts/prepare_aistudio_sources.py --mode quarterly --download-mode full)
echo "$PACKAGE_PATH"
echo ""
echo "Готово. Это квартальный пакет для AI Studio. Путь к папке напечатан выше."
read -k 1 "?Нажми любую клавишу для выхода..."

#!/bin/zsh
set -e
cd "$(dirname "$0")/../../.."
source scripts/unix_python.sh
"$PYTHON_EXE" scripts/apply_aistudio_json.py --mode quarterly
echo ""
echo "Готово. Новый квартальный Excel сохранен в outputs/aistudio_quarterly_excel_update_*."
read -k 1 "?Нажми любую клавишу для выхода..."

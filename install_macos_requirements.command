#!/bin/zsh
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 was not found."
  echo "Install Python 3.11 or 3.12 from https://www.python.org/downloads/macos/"
  echo "Then run this file again."
  echo ""
  read -k 1 "?Press any key to exit..."
  exit 1
fi

python3 -m venv .venv
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements.txt

echo ""
echo "Installed. Use start_dashboard.command to open the dashboard."
read -k 1 "?Press any key to exit..."

#!/bin/zsh
set -e
cd "$(dirname "$0")"
source scripts/unix_python.sh
"$PYTHON_EXE" scripts/dashboard_server.py --open

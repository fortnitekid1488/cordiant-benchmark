#!/bin/zsh

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VENV_PYTHON="$ROOT_DIR/.venv/bin/python"

if [[ -x "$VENV_PYTHON" ]]; then
  export PYTHON_EXE="$VENV_PYTHON"
  return 0 2>/dev/null || exit 0
fi

if command -v python3 >/dev/null 2>&1; then
  export PYTHON_EXE="$(command -v python3)"
  return 0 2>/dev/null || exit 0
fi

if command -v python >/dev/null 2>&1; then
  export PYTHON_EXE="$(command -v python)"
  return 0 2>/dev/null || exit 0
fi

echo "Python 3 was not found. Install Python 3.11 or 3.12 and run install_macos_requirements.command."
return 1 2>/dev/null || exit 1

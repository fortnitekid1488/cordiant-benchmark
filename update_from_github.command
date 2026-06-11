#!/bin/zsh
set -e
cd "$(dirname "$0")"

REPO_URL="https://github.com/fortnitekid1488/cordiant-benchmark.git"
ZIP_URL="https://github.com/fortnitekid1488/cordiant-benchmark/archive/refs/heads/main.zip"

echo "Updating Cordiant dashboard from GitHub main..."
echo "Folder: $PWD"
echo ""

if [ -d ".git" ] && command -v git >/dev/null 2>&1; then
  if git remote get-url origin >/dev/null 2>&1; then
    git remote set-url origin "$REPO_URL"
  else
    git remote add origin "$REPO_URL"
  fi
  git fetch origin main
  git merge --ff-only origin/main
else
  if ! command -v curl >/dev/null 2>&1 || ! command -v unzip >/dev/null 2>&1 || ! command -v rsync >/dev/null 2>&1; then
    echo "curl, unzip, or rsync was not found. Install Git and run: git pull --ff-only origin main"
    read -k 1 "?Press any key to exit..."
    exit 1
  fi

  tmp_dir="$(mktemp -d)"
  zip_path="$tmp_dir/cordiant-benchmark-main.zip"
  curl -L "$ZIP_URL" -o "$zip_path"
  unzip -q "$zip_path" -d "$tmp_dir"
  src_dir="$(find "$tmp_dir" -maxdepth 1 -type d -name 'cordiant-benchmark-*' | head -n 1)"
  if [ -z "$src_dir" ]; then
    echo "Downloaded ZIP did not contain a repository folder."
    rm -rf "$tmp_dir"
    read -k 1 "?Press any key to exit..."
    exit 1
  fi
  rsync -a \
    --exclude ".git" \
    --exclude ".venv" \
    --exclude "outputs" \
    --exclude "__pycache__" \
    --exclude ".pytest_cache" \
    --exclude ".env" \
    "$src_dir/" "$PWD/"
  rm -rf "$tmp_dir"
fi

if [ -x ".venv/bin/python" ] && [ -f "requirements.txt" ]; then
  ".venv/bin/python" -m pip install -r requirements.txt
else
  echo ""
  echo "No local .venv found. If this is the first run on this Mac, run install_macos_requirements.command once."
fi

echo ""
echo "Update complete. Start the dashboard with start_dashboard.command."
read -k 1 "?Press any key to exit..."

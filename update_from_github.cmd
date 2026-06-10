@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

set "REPO_URL=https://github.com/fortnitekid1488/cordiant-benchmark.git"
set "ZIP_URL=https://github.com/fortnitekid1488/cordiant-benchmark/archive/refs/heads/main.zip"

echo Updating Cordiant dashboard from GitHub main...
echo Folder: %CD%
echo.

if exist ".git\" (
  where git >nul 2>nul
  if errorlevel 1 (
    echo This folder is a Git checkout, but git.exe was not found.
    echo Falling back to ZIP update.
    goto zip_update
  )

  git remote get-url origin >nul 2>nul
  if errorlevel 1 (
    git remote add origin "%REPO_URL%"
  ) else (
    git remote set-url origin "%REPO_URL%"
  )
  if errorlevel 1 goto git_error

  git fetch origin main
  if errorlevel 1 goto git_error

  git merge --ff-only origin/main
  if errorlevel 1 goto git_error

  goto update_deps
)

:zip_update
where powershell >nul 2>nul
if errorlevel 1 (
  echo PowerShell was not found. Install Git and run:
  echo git pull --ff-only origin main
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference = 'Stop';" ^
  "$root = (Get-Location).Path;" ^
  "$zip = Join-Path $env:TEMP ('cordiant-benchmark-main-' + [guid]::NewGuid() + '.zip');" ^
  "$tmp = Join-Path $env:TEMP ('cordiant-benchmark-main-' + [guid]::NewGuid());" ^
  "Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile $zip;" ^
  "Expand-Archive -Path $zip -DestinationPath $tmp -Force;" ^
  "$src = Get-ChildItem -Path $tmp -Directory | Select-Object -First 1;" ^
  "if (-not $src) { throw 'Downloaded ZIP did not contain a repository folder.' }" ^
  "robocopy $src.FullName $root /E /XD .git .venv outputs __pycache__ .pytest_cache /XF .env /NFL /NDL /NJH /NJS /NP;" ^
  "$code = $LASTEXITCODE;" ^
  "Remove-Item $zip -Force -ErrorAction SilentlyContinue;" ^
  "Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue;" ^
  "if ($code -ge 8) { exit $code } else { exit 0 }"
if errorlevel 1 goto zip_error

:update_deps
if exist ".venv\Scripts\python.exe" (
  if exist "requirements.txt" (
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto deps_error
  )
) else (
  echo.
  echo No local .venv found. If this is the first run on this PC, run install_windows_requirements.cmd once.
)

echo.
echo Update complete. Start the dashboard with start_dashboard.cmd.
pause
exit /b 0

:git_error
echo.
echo Git update failed. If you changed tracked project files locally, save them first or use a clean project folder.
pause
exit /b 1

:zip_error
echo.
echo ZIP update failed. Check internet access and that the GitHub repository is public.
pause
exit /b 1

:deps_error
echo.
echo Project files were updated, but dependency refresh failed. Try install_windows_requirements.cmd.
pause
exit /b 1

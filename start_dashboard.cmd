@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
call scripts\windows_python.cmd
if errorlevel 1 (
  pause
  exit /b 1
)

%PYTHON_EXE% scripts\dashboard_server.py --open
if errorlevel 1 (
  echo.
  echo Dashboard failed.
  pause
  exit /b 1
)

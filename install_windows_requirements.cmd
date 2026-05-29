@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
  set "BOOTSTRAP_PYTHON=py -3"
) else (
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    set "BOOTSTRAP_PYTHON=python"
  ) else (
    echo Python 3 was not found. Install Python 3.11 or 3.12 from python.org and enable Add python.exe to PATH.
    pause
    exit /b 1
  )
)

%BOOTSTRAP_PYTHON% -m venv .venv
if errorlevel 1 goto error

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto error

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto error

echo.
echo Installed. Use start_dashboard.cmd or the prepare/update .cmd launchers.
pause
exit /b 0

:error
echo.
echo Setup failed.
pause
exit /b 1

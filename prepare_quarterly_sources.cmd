@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
call scripts\windows_python.cmd
if errorlevel 1 (
  pause
  exit /b 1
)

%PYTHON_EXE% scripts\prepare_aistudio_sources.py --mode quarterly --download-mode full
if errorlevel 1 goto error

set "PACKAGE_FILE=outputs\latest_aistudio_quarterly_package_path.txt"
if exist "%PACKAGE_FILE%" (
  for /f "usebackq delims=" %%P in ("%PACKAGE_FILE%") do set "PACKAGE_PATH=%%P"
  if defined PACKAGE_PATH explorer "%PACKAGE_PATH%"
)

echo.
echo Done. Quarterly package is ready.
pause
exit /b 0

:error
echo.
echo Quarterly source preparation failed.
pause
exit /b 1

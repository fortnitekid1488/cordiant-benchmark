@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
call scripts\windows_python.cmd
if errorlevel 1 (
  pause
  exit /b 1
)

%PYTHON_EXE% scripts\apply_aistudio_json.py --mode quarterly
if errorlevel 1 goto error

echo.
echo Done. New quarterly Excel was saved under outputs\aistudio_quarterly_excel_update_*.
pause
exit /b 0

:error
echo.
echo Quarterly Excel update failed.
pause
exit /b 1

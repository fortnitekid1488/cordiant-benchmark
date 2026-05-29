@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
call scripts\windows_python.cmd
if errorlevel 1 (
  pause
  exit /b 1
)

%PYTHON_EXE% scripts\apply_aistudio_json.py --mode annual
if errorlevel 1 goto error

echo.
echo Done. New annual Excel was saved under outputs\aistudio_annual_excel_update_*.
pause
exit /b 0

:error
echo.
echo Annual Excel update failed.
pause
exit /b 1

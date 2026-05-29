@echo off
set "VENV_PYTHON=%CD%\.venv\Scripts\python.exe"
if exist "%VENV_PYTHON%" (
  set "PYTHON_EXE="%VENV_PYTHON%""
  exit /b 0
)

where py >nul 2>nul
if %ERRORLEVEL%==0 (
  set "PYTHON_EXE=py -3"
  exit /b 0
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
  set "PYTHON_EXE=python"
  exit /b 0
)

echo Python 3 was not found. Install Python 3.11 or 3.12 and run install_windows_requirements.cmd.
exit /b 1

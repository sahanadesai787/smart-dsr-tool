@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.11 or newer from https://www.python.org/downloads/ and select Add Python to PATH.
  pause
  exit /b 1
)
python -m pip install -r requirements.txt
python -m streamlit run app.py
pause

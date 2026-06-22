@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo   Conversion Studio Windows Agent
echo ================================================
echo.

if exist ".venv\Scripts\activate.bat" call ".venv\Scripts\activate.bat"
if exist "myenv\Scripts\activate.bat" call "myenv\Scripts\activate.bat"

python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
  echo ERROR: pywin32 is not installed in the active Python environment.
  echo Run: pip install pywin32
  pause
  exit /b 1
)

echo Starting local agent at http://127.0.0.1:8000
python -m uvicorn main:app --host 127.0.0.1 --port 8000
pause

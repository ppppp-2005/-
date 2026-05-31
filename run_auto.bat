@echo off
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"

if not exist "%PY%" (
    echo Creating virtual environment...
    python -m venv "%~dp0.venv"
)

if not exist "%PY%" (
    echo ERROR: Cannot find %PY%
    echo Try: python -m venv .venv
    pause
    exit /b 1
)

echo Installing dependencies...
"%PY%" -m pip install -r "%~dp0requirements.txt" -q

echo Running formatter...
"%PY%" "%~dp0run_auto.py"
set ERR=%ERRORLEVEL%

echo.
echo === run_result.txt ===
if exist "%~dp0output\run_result.txt" (
    type "%~dp0output\run_result.txt"
) else (
    echo output\run_result.txt not found
)

echo.
if %ERR% neq 0 (
    echo FAILED exit code %ERR%
) else (
    echo DONE. Open output\from_template*.docx in Word
)
pause
exit /b %ERR%

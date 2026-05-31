@echo off
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"

if not exist "%PY%" (
    echo Creating virtual environment...
    python -m venv "%~dp0.venv"
)

if not exist "%PY%" (
    echo ERROR: Cannot find Python venv
    pause
    exit /b 1
)

"%PY%" -m pip install -r "%~dp0requirements.txt" -q
"%PY%" "%~dp0template_main.py" "%~dp0examples\123.docx" "%~dp0templates\111.docx" -o "%~dp0output\from_template.docx"

pause

@echo off
setlocal
cd /d "%~dp0"

REM Double-click uses a minimal PATH; "python" often fails with "Python was not found" (9009).
REM Try the Windows Python Launcher first (python.org installer), then "python".

py -3 -c "import sys" 1>nul 2>nul
if %errorlevel% equ 0 (
    py -3 -m streamlit run pricing_ui.py
    goto :check_err
)

python -c "import sys" 1>nul 2>nul
if %errorlevel% equ 0 (
    python -m streamlit run pricing_ui.py
    goto :check_err
)

echo.
echo [ERROR] Neither "py -3" nor "python" worked from this shortcut.
echo.
echo Fix options:
echo   1. Install Python from https://www.python.org/downloads/
echo      - On the first screen, enable "Add python.exe to PATH"
echo   2. Or open "Anaconda Prompt" / a terminal where Python works, then run:
echo      cd /d "%~dp0"
echo      streamlit run pricing_ui.py
echo   3. Or activate your venv first, then run streamlit from this folder.
echo.
pause
exit /b 1

:check_err
if errorlevel 1 (
    echo.
    echo Streamlit exited with an error (see messages above).
    pause
)

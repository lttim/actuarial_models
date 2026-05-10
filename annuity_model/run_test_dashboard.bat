@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
set "PYTHONPATH=%CD%\src;%PYTHONPATH%"

set MIN_PY_MAJOR=3
set MIN_PY_MINOR=11
set SELF_CHECK=0
if "%~1"=="--self-check" set SELF_CHECK=1

set PY=
if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
    goto :got_py
)
if defined VIRTUAL_ENV if exist "%VIRTUAL_ENV%\Scripts\python.exe" (
    set "PY=%VIRTUAL_ENV%\Scripts\python.exe"
    goto :got_py
)
py -3 -c "import sys" 1>nul 2>nul
if %errorlevel% equ 0 (
    set "PY=py -3"
    goto :got_py
)
python -c "import sys" 1>nul 2>nul
if %errorlevel% equ 0 (
    set "PY=python"
    goto :got_py
)

echo [ERROR] No usable Python interpreter.
pause
exit /b 1

:got_py
%PY% -c "import sys; sys.exit(0 if sys.version_info[:2] >= (%MIN_PY_MAJOR%, %MIN_PY_MINOR%) else 1)"
if errorlevel 1 (
    echo [ERROR] Python is too old for this project.
    pause
    exit /b 1
)

%PY% -c "import streamlit" 1>nul 2>nul
if errorlevel 1 (
    if exist ".venv\Scripts\python.exe" (
        echo [INFO] Installing pinned dependencies into .venv ...
        %PY% -m pip install -r requirements.txt
    ) else if defined VIRTUAL_ENV (
        echo [INFO] Installing pinned dependencies into the active venv ...
        %PY% -m pip install -r requirements.txt
    ) else (
        echo [ERROR] streamlit not importable; create .venv first.
        pause
        exit /b 1
    )
)

%PY% -c "import annuity_model.test_dashboard" 1>nul 2>nul
if errorlevel 1 (
    echo [ERROR] Failed to import annuity_model.test_dashboard.
    %PY% -c "import annuity_model.test_dashboard"
    pause
    exit /b 1
)

if "%SELF_CHECK%"=="1" (
    echo [OK] test_dashboard launcher self-check passed.
    exit /b 0
)

%PY% -m streamlit run src\annuity_model\test_dashboard.py
if errorlevel 1 (
    echo Streamlit exited with an error.
    pause
)
endlocal

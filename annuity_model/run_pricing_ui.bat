@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Multi-policy portfolio UI: default ON for this local launcher. Opt out: create .disable-portfolio-v1
if exist ".disable-portfolio-v1" (
    set "ANNUITY_MODEL_PORTFOLIO_V1=0"
) else if not defined ANNUITY_MODEL_PORTFOLIO_V1 (
    set "ANNUITY_MODEL_PORTFOLIO_V1=1"
)

REM Hardening rules (mirrored by tests/test_launcher_invariants.py):
REM   1. PROJECT-VENV-FIRST: prefer .venv\Scripts\python.exe.
REM   2. MIN-PYTHON: require Python >= 3.11 (kept in sync with pyproject.toml).
REM   3. IMPORT-SMOKE: confirm `pricing_ui` itself imports before launching streamlit.
REM   4. SELF-CHECK: `--self-check` runs (1)+(2)+(3) and exits without starting streamlit.

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

echo.
echo [ERROR] No usable Python interpreter (need ^>= %MIN_PY_MAJOR%.%MIN_PY_MINOR%).
echo.
echo Fix options:
echo   1. Install Python from https://www.python.org/downloads/
echo      - On the first screen, enable "Add python.exe to PATH"
echo   2. Create the project venv:
echo        py -3 -m venv .venv
echo        .venv\Scripts\python.exe -m pip install -r requirements.txt
echo.
pause
exit /b 1

:got_py
%PY% -c "import sys; sys.exit(0 if sys.version_info[:2] >= (%MIN_PY_MAJOR%, %MIN_PY_MINOR%) else 1)"
if errorlevel 1 (
    echo.
    echo [ERROR] %PY% is too old for this project ^(requires ^>= %MIN_PY_MAJOR%.%MIN_PY_MINOR%^).
    echo Install a newer Python or recreate .venv with Python 3.12.
    echo.
    pause
    exit /b 1
)

%PY% -c "import streamlit" 1>nul 2>nul
if errorlevel 1 (
    if exist ".venv\Scripts\python.exe" (
        echo [INFO] Installing pinned dependencies into .venv ...
        %PY% -m pip install -r requirements.txt
        if errorlevel 1 (
            echo [ERROR] Dependency install failed.
            pause
            exit /b 1
        )
    ) else if defined VIRTUAL_ENV (
        echo [INFO] Installing pinned dependencies into the active venv ...
        %PY% -m pip install -r requirements.txt
        if errorlevel 1 (
            echo [ERROR] Dependency install failed.
            pause
            exit /b 1
        )
    ) else (
        echo [ERROR] streamlit not importable; refusing to pip install into a non-venv interpreter.
        echo Run: py -3 -m venv .venv ^&^& .venv\Scripts\python.exe -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

%PY% -c "import pricing_ui" 1>nul 2>nul
if errorlevel 1 (
    echo [ERROR] Failed to import pricing_ui. Re-running with traceback:
    %PY% -c "import pricing_ui"
    pause
    exit /b 1
)

if "%SELF_CHECK%"=="1" (
    echo [OK] Launcher self-check passed.
    exit /b 0
)

%PY% -m streamlit run pricing_ui.py
if errorlevel 1 (
    echo.
    echo Streamlit exited with an error (see messages above).
    pause
)
endlocal

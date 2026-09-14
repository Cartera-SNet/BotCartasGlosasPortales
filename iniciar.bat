@echo off
:: ============================================================
::  Activa IT -- Descargador de Cartas Glosa (panel unificado)
::  Estado + Sura + Bolivar + Previsora + Mundial
:: ============================================================

setlocal
echo.
echo ============================================================
echo   Activa IT -- Panel unificado (5 aseguradoras)
echo ============================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%venv"

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado.
    pause
    exit /b 1
)

if not exist "%VENV_DIR%" (
    echo Creando entorno virtual Python...
    python -m venv "%VENV_DIR%"
)

call "%VENV_DIR%\Scripts\activate.bat"

echo Instalando dependencias...
pip install --quiet -r "%SCRIPT_DIR%requirements.txt"

echo Instalando Playwright Chromium...
python -m playwright install chromium

for %%B in (estado sura bolivar previsora mundial) do (
    if not exist "%SCRIPT_DIR%downloads\%%B" mkdir "%SCRIPT_DIR%downloads\%%B"
)

start "" "http://localhost:8080"
python "%SCRIPT_DIR%app.py"

pause

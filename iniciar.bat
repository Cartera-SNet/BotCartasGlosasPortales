@echo off
:: ============================================================
::  Panel Bot Glosas -- version RAILWAY, corrida local para
::  PROBAR antes de subir a produccion.
:: ============================================================

setlocal
echo.
echo ============================================================
echo   Panel Bot Glosas (version Railway) -- prueba local
echo ============================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%venv"

:: ============================================================
::  Conexion a base de datos
::  Por defecto, DATABASE_URL queda VACIA -- eso hace que el
::  programa use una base SQLite local aparte, SIN tocar tu Neon
::  real de produccion. Esto es a proposito, para que puedas
::  probar cambios sin riesgo de mezclar datos de prueba con los
::  reales.
::
::  Si alguna vez SI quieres probar contra la base real de Neon,
::  descomenta la siguiente linea (quitale el "::" de adelante) y
::  pon tu cadena de conexion real:
:: set "DATABASE_URL=postgresql://usuario:contrasena@host/neondb?sslmode=require"
:: ============================================================

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

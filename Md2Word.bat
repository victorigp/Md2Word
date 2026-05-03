@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM Md2Word.bat - Lanzador principal para convertir Markdown a Word
REM ============================================================================

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Leer InterpreterPath de Settings.json
for /f "usebackq delims=" %%A in (`powershell -NoProfile -Command "(Get-Content '%SCRIPT_DIR%Settings.json' | ConvertFrom-Json).Python.InterpreterPath"`) do set "PYTHON=%%A"

if not exist "%PYTHON%" (
    echo [ERROR] Python no encontrado en: %PYTHON%
    echo Actualice InterpreterPath en Settings.json
    pause
    exit /b 1
)

REM Buscar el primer .md en docs\ (por fecha más antigua)
set "MD_FILE="
for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command "Get-ChildItem '%SCRIPT_DIR%docs\*.md' | Sort-Object LastWriteTime | Select-Object -First 1 -ExpandProperty FullName"`) do set "MD_FILE=%%F"

if not defined MD_FILE (
    echo [ERROR] No se encontró ningún archivo .md en docs\
    pause
    exit /b 1
)

REM Buscar la primera plantilla .docx en docs\ (excluir ficheros generados)
set "TEMPLATE="
for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command "Get-ChildItem '%SCRIPT_DIR%docs\*.docx' | Where-Object { $_.Name -notmatch '_\d+\.docx$' } | Sort-Object LastWriteTime | Select-Object -First 1 -ExpandProperty FullName"`) do set "TEMPLATE=%%F"

if not defined TEMPLATE (
    echo [ERROR] No se encontró ninguna plantilla .docx en docs\
    pause
    exit /b 1
)

REM Obtener título del .md via GetTitle.ps1
for /f "usebackq delims=" %%T in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%GetTitle.ps1" "%MD_FILE%"`) do set "TITLE=%%T"

if not defined TITLE set "TITLE=Documento"

REM Determinar nombre de salida con sufijo si ya existe
set "OUTPUT=%SCRIPT_DIR%docs\%TITLE%.docx"
set "SUFFIX=0"

:check_exists
if exist "!OUTPUT!" (
    set /a SUFFIX+=1
    set "OUTPUT=%SCRIPT_DIR%docs\%TITLE%_!SUFFIX!.docx"
    goto check_exists
)

echo ============================================================
echo  Md2Word - Conversión Markdown a Word
echo ============================================================
echo  Markdown : %MD_FILE%
echo  Plantilla: %TEMPLATE%
echo  Salida   : %OUTPUT%
echo ============================================================

"%PYTHON%" "%SCRIPT_DIR%Md2Word.py" "%MD_FILE%" "%TEMPLATE%" "%OUTPUT%"

if %ERRORLEVEL% neq 0 (
    echo [ERROR] La conversión falló.
    pause
    exit /b 1
)

echo.
echo [OK] Proceso completado.
pause

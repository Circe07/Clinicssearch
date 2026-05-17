@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title Buscador de Clinicas - Detector de Oportunidades Web
color 0A

echo.
echo ============================================================
echo   BUSCADOR DE CLINICAS - Detector de Oportunidades Web
echo ============================================================
echo.

set "CONFIG_FILE=%~dp0.api_key"

if exist "!CONFIG_FILE!" (
    set /p API_KEY=<"!CONFIG_FILE!"
    echo   API key encontrada. Usando la guardada.
    echo   Para cambiarla, borra el archivo .api_key
    echo.
    goto MENU
)

echo   Primera vez? Necesitas tu API key de Google Places.
echo   Solo se pide una vez, se guarda para futuras ejecuciones.
echo.
set /p "API_KEY=   Introduce tu API key de Google Places: "

if "!API_KEY!"=="" (
    echo.
    echo   [ERROR] No has introducido ninguna API key.
    pause
    exit /b 1
)

echo !API_KEY!>"!CONFIG_FILE!"
echo.
echo   API key guardada correctamente.
echo.

:MENU
echo.
echo ============================================================
echo   Que quieres buscar?
echo ============================================================
echo.
echo   Ejemplos:
echo     - clinicas dentales en Madrid
echo     - clinicas esteticas en Barcelona
echo     - fisioterapia en Valencia
echo     - veterinarias en Sevilla
echo     - podologia en Bilbao
echo.

set "QUERY="
set /p "QUERY=   Tu busqueda: "

if "!QUERY!"=="" (
    echo.
    echo   [ERROR] No has escrito ninguna busqueda.
    goto MENU
)

set "MIN_REVIEWS="
echo.
set /p "MIN_REVIEWS=   Minimo de resenas para incluir (Enter = 50): "
if "!MIN_REVIEWS!"=="" set "MIN_REVIEWS=50"

set "OUTPUT="
echo.
set /p "OUTPUT=   Nombre del archivo Excel (Enter = clinicas_prospecto.xlsx): "
if "!OUTPUT!"=="" set "OUTPUT=clinicas_prospecto.xlsx"

echo.
echo ============================================================
echo   Lanzando busqueda...
echo   Busqueda:      !QUERY!
echo   Min. resenas:  !MIN_REVIEWS!
echo   Archivo:       !OUTPUT!
echo ============================================================
echo.

python buscar_clinicas.py -k "!API_KEY!" -q "!QUERY!" -o "!OUTPUT!" -n !MIN_REVIEWS!

if !ERRORLEVEL! neq 0 (
    echo.
    echo   [ERROR] Algo fallo. Revisa los mensajes de arriba.
    echo   Si el error es de API key, borra el archivo .api_key y ejecuta de nuevo.
)

echo.
echo ============================================================

set "BUSCAR_EMAIL="
set /p "BUSCAR_EMAIL=   Quieres buscar emails que faltan en el Excel? (s/n): "
if /i "!BUSCAR_EMAIL!"=="s" goto BUSCAR_EMAILS
if /i "!BUSCAR_EMAIL!"=="si" goto BUSCAR_EMAILS
goto PREGUNTAR_OTRA

:BUSCAR_EMAILS
echo.
echo   Buscando emails en las webs de las clinicas...
echo.
python buscar_emails.py -i "!OUTPUT!"
echo.
echo ============================================================

:PREGUNTAR_OTRA
set "OTRA="
set /p "OTRA=   Quieres hacer otra busqueda? (s/n): "
if /i "!OTRA!"=="s" goto MENU
if /i "!OTRA!"=="si" goto MENU

echo.
echo   Hasta luego!
echo.
pause
endlocal

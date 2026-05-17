@echo off
chcp 65001 >nul 2>&1
title Buscador de Clínicas - Detector de Oportunidades Web
color 0A

echo.
echo ============================================================
echo   BUSCADOR DE CLINICAS - Detector de Oportunidades Web
echo ============================================================
echo.

:: Archivo donde se guarda la API key para no pedirla cada vez
set "CONFIG_FILE=%~dp0.api_key"

:: Comprobar si ya existe una API key guardada
if exist "%CONFIG_FILE%" (
    set /p API_KEY=<"%CONFIG_FILE%"
    echo   API key encontrada. Usando la guardada.
    echo   ^(Para cambiarla, borra el archivo .api_key^)
    echo.
) else (
    echo   Primera vez? Necesitas tu API key de Google Places.
    echo   ^(Solo se pide una vez, se guarda para futuras ejecuciones^)
    echo.
    set /p API_KEY="   Introduce tu API key de Google Places: "
    if "!API_KEY!"=="" (
        echo.
        echo   [ERROR] No has introducido ninguna API key.
        pause
        exit /b 1
    )
    echo !API_KEY!>"%CONFIG_FILE%"
    echo.
    echo   API key guardada correctamente.
    echo.
)

:: Habilitar expansión retardada después de leer la key
setlocal enabledelayedexpansion

:: Re-leer la API key con expansión retardada habilitada
if exist "%CONFIG_FILE%" (
    set /p API_KEY=<"%CONFIG_FILE%"
)

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
set /p QUERY="   Tu busqueda: "

if "!QUERY!"=="" (
    echo.
    echo   [ERROR] No has escrito ninguna busqueda.
    goto MENU
)

echo.
set /p OUTPUT="   Nombre del archivo Excel (Enter = clinicas_prospecto.xlsx): "
if "!OUTPUT!"=="" set "OUTPUT=clinicas_prospecto.xlsx"

echo.
echo ============================================================
echo   Lanzando busqueda...
echo   Busqueda: !QUERY!
echo   Archivo:  !OUTPUT!
echo ============================================================
echo.

python buscar_clinicas.py -k "!API_KEY!" -q "!QUERY!" -o "!OUTPUT!"

if %ERRORLEVEL% neq 0 (
    echo.
    echo   [ERROR] Algo fallo. Revisa los mensajes de arriba.
    echo   Si el error es de API key, borra el archivo .api_key y ejecuta de nuevo.
)

echo.
echo ============================================================
set /p OTRA="   Quieres hacer otra busqueda? (s/n): "
if /i "!OTRA!"=="s" goto MENU

echo.
echo   Hasta luego!
echo.
pause
endlocal

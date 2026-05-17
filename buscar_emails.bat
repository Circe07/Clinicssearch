@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title Buscador de Emails - Completar datos de clinicas
color 0B

echo.
echo ============================================================
echo   BUSCADOR DE EMAILS - Completar datos de clinicas
echo ============================================================
echo.

set "INPUT="
set /p "INPUT=   Archivo Excel (Enter = clinicas_prospecto.xlsx): "
if "!INPUT!"=="" set "INPUT=clinicas_prospecto.xlsx"

echo.
python buscar_emails.py -i "!INPUT!"

echo.
pause
endlocal

@echo off
REM ============================================================================
REM Script para instalar los protocolos personalizados de todas las aplicaciones
REM ============================================================================

echo.
echo ========================================
echo   Instalador de Protocolos - Apps
echo ========================================
echo.
echo Este script registrara 13 protocolos personalizados en Windows
echo para ejecutar cada aplicacion directamente desde el navegador.
echo.
echo Presione cualquier tecla para continuar o Ctrl+C para cancelar..
pause >nul

echo.
echo Instalando protocolos...

REM Importar el archivo .reg
reg import "%~dp0..\03_OTROS\instalar_protocolos_apps.reg"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] Protocolos instalados correctamente!
    echo.
    echo Protocolos registrados:
    echo   - sueldos-calculadora:  Calculadora de Horas
    echo   - sueldos-recibos:      Generar Recibos
    echo   - sueldos-procesar:     Procesar Sueldos
    echo   - sueldos-buscar:       Buscar Recibos PDF
    echo   - sueldos-deposito:     Horas Deposito
    echo   - sueldos-sobres:       Imprimir Sobres
    echo   - sueldos-aguinaldo:    Aguinaldo Unificado
    echo   - sueldos-email:        Enviar Documentacion
    echo   - sueldos-firmas:       Extractor de Firmas PDF
    echo   - sueldos-epp:          Formulario EPP
    echo   - sueldos-lector:       Lector Inteligente IA
    echo   - sueldos-promedio:     Promedios de Sueldos
    echo   - sueldos-pago:         Guia Pago Bancario
    echo   - sueldos-conceptos:    Buscador de Conceptos
    echo   - sueldos-fechas:       Fechas de Ingreso
    echo   - sueldos-planilla:     Planilla x Índice
    echo   - sueldos-acomodar:     Acomodar PDF
    echo   - sueldos-firmador:     Firmador Masivo PDF
    echo.
    echo Ahora puede usar el archivo Launcher_Web.html desde el navegador.
    echo Cada aplicacion se ejecutara directamente al hacer clic.
    echo.
) else (
    echo.
    echo [ERROR] No se pudieron instalar los protocolos.
    echo Por favor, ejecute este script como Administrador.
    echo.
)

echo.
echo Presione cualquier tecla para salir...
pause >nul

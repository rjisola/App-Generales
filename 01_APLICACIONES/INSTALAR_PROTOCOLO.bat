@echo off
REM ============================================================================
REM Script para instalar el protocolo personalizado sueldoslauncher:
REM ============================================================================

echo.
echo ========================================
echo   Instalador de Protocolo Launcher
echo ========================================
echo.
echo Este script registrara el protocolo "sueldoslauncher:" en Windows
echo para que pueda abrir el launcher desde el navegador.
echo.
echo Presione cualquier tecla para continuar o Ctrl+C para cancelar...
pause >nul

echo.
echo Instalando protocolo...

REM Importar el archivo .reg
reg import "%~dp0instalar_protocolo.reg"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] Protocolo instalado correctamente!
    echo.
    echo Ahora puede usar el archivo AbrirLauncher.html desde el navegador.
    echo.
) else (
    echo.
    echo [ERROR] No se pudo instalar el protocolo.
    echo Por favor, ejecute este script como Administrador.
    echo.
)

echo.
echo Presione cualquier tecla para salir...
pause >nul

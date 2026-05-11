@echo off
:: ============================================================
::  ABRIR SISTEMA DE SUELDOS EN MICROSOFT EDGE
::  Abre en modo Aplicacion (sin barras de navegacion)
::  maximizado, forzando recarga sin cache
:: ============================================================

setlocal

:: Ruta al archivo HTML del launcher
set "HTML=%~dp0Launcher_Web.html"

:: ---- Intentar encontrar Edge ----
set "EDGE="
for %%P in (
    "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
    "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
    "%LocalAppData%\Microsoft\Edge\Application\msedge.exe"
) do (
    if exist %%P (
        set "EDGE=%%~P"
        goto :found
    )
)

:: Edge no encontrado → abrir con navegador por defecto
echo [AVISO] Microsoft Edge no encontrado. Abriendo con navegador predeterminado...
start "" "%HTML%"
goto :eof

:found
:: Abrir en modo App (sin barras) + maximizado + sin cache
start "" "%EDGE%" ^
    --app="file:///%HTML:\=/%"  ^
    --start-maximized            ^
    --disable-cache              ^
    --disk-cache-size=0          ^
    --no-first-run               ^
    --disable-extensions

endlocal

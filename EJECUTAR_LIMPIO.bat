@echo off
echo Limpiando cache de Python...
if exist __pycache__ rmdir /s /q __pycache__
if exist modern_gui_components.pyc del /f modern_gui_components.pyc

echo Ejecutando aplicacion...
python A-GENERAR_RECIBOS_CONTROL.pyw

echo.
echo Presiona cualquier tecla para cerrar...
pause >nul

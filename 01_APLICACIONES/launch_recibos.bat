@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "A-GENERAR_RECIBOS_CONTROL.pyw"
exit

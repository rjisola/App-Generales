@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "actualizar_quincena_gui.pyw"
exit

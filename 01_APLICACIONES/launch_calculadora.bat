@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "10-CALCULAR_HORAS.pyw"
exit

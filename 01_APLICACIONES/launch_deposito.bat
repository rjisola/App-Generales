@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "15-PASAR_HORAS_DEPOSITO.pyw"
exit

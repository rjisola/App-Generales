@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "Asistente_Aguinaldo_UNIFICADO.pyw"
exit

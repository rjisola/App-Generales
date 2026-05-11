@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "Enviar_Documentacion_Email.pyw"
exit

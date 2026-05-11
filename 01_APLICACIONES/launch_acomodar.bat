@echo off
set PYTHONPATH=%~dp0..\03_OTROS;%PYTHONPATH%
cd /d "%~dp0"
start "" pythonw "gui_acomodar_pdf.pyw"
exit

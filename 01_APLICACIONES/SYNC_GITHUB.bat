@echo off
TITLE Sincronizando App-Generales con GitHub
cd /d "%~dp0"
echo Iniciando sincronizacion...
python sync_github.py
if %ERRORLEVEL% NEQ 0 pause
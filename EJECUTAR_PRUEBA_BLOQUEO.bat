@echo off
echo Ejecutando prueba de bloqueo...
python test_lock.py
if %errorlevel% neq 0 pause
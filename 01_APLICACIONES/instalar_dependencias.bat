@echo off
echo ========================================
echo Instalando dependencias del sistema
echo ========================================
echo.

python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ========================================
echo Instalacion completada
echo ========================================
pause

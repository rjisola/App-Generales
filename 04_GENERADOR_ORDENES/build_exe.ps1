# Script de compilación para generar el ejecutable
# Requiere: pip install pyinstaller

pyinstaller --noconsole --onefile --clean `
    --name "Generador_Ordenes" `
    --add-data "schema.sql;." `
    main.py

Write-Host "`nEXITO: El ejecutable se encuentra en la carpeta 'dist'." -ForegroundColor Green

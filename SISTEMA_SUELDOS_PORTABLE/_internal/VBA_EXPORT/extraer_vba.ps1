$excelPath = "C:\Users\rjiso\OneDrive\Escritorio\1ERA ABRIL 2026\PROGRAMA DEPOSITO 1ERA ABRIL2026.xlsm"
$outputDir = "C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\VBA_EXPORT"

if (-not (Test-Path $excelPath)) {
    Write-Host "ARCHIVO NO ENCONTRADO: $excelPath"
    exit 1
}

Write-Host "Abriendo Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($excelPath)
    Write-Host "Archivo abierto OK"
    
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

    # Listar hojas
    Write-Host "=== HOJAS DEL LIBRO ==="
    foreach ($sheet in $wb.Sheets) {
        Write-Host ("  Hoja: " + $sheet.Name)
    }

    # Extraer componentes VBA
    Write-Host "=== MODULOS VBA ==="
    $vbaProject = $wb.VBProject
    foreach ($component in $vbaProject.VBComponents) {
        $name = $component.Name
        $lines = $component.CodeModule.CountOfLines
        Write-Host ("  Modulo: " + $name + " | Lineas: " + $lines)
        if ($lines -gt 0) {
            $code = $component.CodeModule.Lines(1, $lines)
            $outFile = Join-Path $outputDir ($name + ".vba")
            [System.IO.File]::WriteAllText($outFile, $code, [System.Text.Encoding]::UTF8)
        }
    }

    $wb.Close($false)
    Write-Host "EXPORTACION COMPLETA"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

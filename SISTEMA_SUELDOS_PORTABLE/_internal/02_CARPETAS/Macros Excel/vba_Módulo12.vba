Sub LimpiarValores()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    
    Sheets("CALCULAR HORAS").Range("AM9:AM500" & Range("A" & Rows.Count).End(xlUp).Row).ClearContents
    Sheets("CALCULAR HORAS").Range("C500:R1000").ClearContents
    Sheets("CALCULAR HORAS").Range("C500:R1000").Borders.LineStyle = xlNone
    Sheets("CALCULAR HORAS").Range("C500:R1000").Interior.ColorIndex = xlNone
    Sheets("SUELDO_ALQ_GASTOS").Range("AM9:AM500" & Range("A" & Rows.Count).End(xlUp).Row).ClearContents
    
    
    
    'Borrar el rango de la hoja RECUENTO PAPELERA desde B1 hasta K500
    Sheets("RECUENTO TOTAL").Range("a2:K500").ClearContents
    Sheets("RECUENTO TOTAL").Range("A2:K500").Interior.color = RGB(211, 235, 247)
    
    'Agregar color a las celdas en la hoja IMPRIMIR TOTALES
    Sheets("IMPRIMIR TOTALES").Range("A1:F5000").ClearContents
    Sheets("IMPRIMIR TOTALES").Range("A1:F5000").Interior.color = RGB(255, 255, 255)
    
    
    
    
    
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

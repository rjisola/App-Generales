Sub CompararSUELDOCONTADOR()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Especifica el nombre de la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("ENVIO CONTADOR")
    
    ' Encuentra la última fila con datos en la columna C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Inicializar variables para las columnas que se van a comparar
    Dim col1 As String, col2 As String, resultCol As String
    col1 = "B"
    col2 = "BC"
    resultCol = "BF"
    
    ' Recorre las filas desde la 9 hasta la última fila con datos en columna C
    For i = 9 To lastRow
        ' Compara el valor de las columnas H y BC en cada fila
        If ws.Cells(i, col1).Value = ws.Cells(i, col2).Value Then
            ' Si son iguales, coloca "VERDADERO" en la columna BM
            ws.Cells(i, resultCol).Value = "VERDADERO"
        Else
            ' Si no son iguales, coloca "FALSO" en la columna BM
            ws.Cells(i, resultCol).Value = "FALSO"
        End If
    Next i
    
    ' Repite el mismo proceso para las otras comparaciones
    
    ' Columnas I y BD, con resultado en columna BN
    col1 = "AT"
    col2 = "BE"
    resultCol = "BG"
    
    For i = 9 To lastRow
        If ws.Cells(i, col1).Value = ws.Cells(i, col2).Value Then
            ws.Cells(i, resultCol).Value = "VERDADERO"
        Else
            ws.Cells(i, resultCol).Value = "FALSO"
        End If
    Next i
    
    
End Sub


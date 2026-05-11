Sub LimpiarValoresPL()
    Dim ws As Worksheet
    Dim Rango As Range
    
    For Each ws In ThisWorkbook.Worksheets(Array("CALCULAR HORAS", "SUELDO_ALQ_GASTOS", "ENVIO CONTADOR", "RECUENTO TOTAL", "IMPRIMIR TOTALES"))
        Select Case ws.Name
            Case "CALCULAR HORAS"
                ws.Range("S9:AJ" & ws.Range("A" & ws.Rows.Count).End(xlUp).Row).ClearContents
            Case "SUELDO_ALQ_GASTOS"
                
                Set Rango = ActiveSheet.Range("C500:R500")
                If WorksheetFunction.CountA(Rango) = 0 Then ' Verificar si hay datos en el rango
                    ws.Range("C500:R1000").Clear
                Else
                    ActiveCell.Offset(1, 0).Select ' Si hay datos, seleccionar la celda debajo de la activa
                    MsgBox "El rango " & Rango.Address & " ya contiene datos.", vbInformation, "Rango con datos"
                End If
            Case "RECUENTO TOTAL"
                ws.Range("B1:K500").ClearContents
            Case "IMPRIMIR TOTALES"
                ws.Range("A1:F5000").ClearContents
        End Select
    Next ws
End Sub

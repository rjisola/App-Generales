Sub BorrarContenidoCeldas3()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets(Array("CALCULAR HORAS", "SUELDO_ALQ_GASTOS", "ENVIO CONTADOR", "RECUENTO TOTAL", "IMPRIMIR TOTALES"))
    
    
        Select Case ws.Name
            Case "CALCULAR HORAS"
                ws.Range("C9:R" & ws.Range("A" & ws.Rows.Count).End(xlUp).Row).ClearContents
                ws.Range("AM9:AM" & ws.Range("A" & ws.Rows.Count).End(xlUp).Row).ClearContents
                ws.Range("S9:AJ" & ws.Range("A" & ws.Rows.Count).End(xlUp).Row).ClearContents
                ws.Range("C500:R1000").Clear
             Case "SUELDO_ALQ_GASTOS"
                ws.Range("AM9:AM" & ws.Range("K" & ws.Rows.Count).End(xlUp).Row).ClearContents
                
            Case "RECUENTO TOTAL"
                ws.Range("A1:K500").ClearContents
                ws.Range("A1:K500" & ws.Range("K" & ws.Rows.Count).End(xlUp).Row).Interior.color = RGB(211, 235, 247)
            Case "IMPRIMIR TOTALES"
                ws.Range("A1:F5000").ClearContents
                ws.Range("A1:f5000" & ws.Range("K" & ws.Rows.Count).End(xlUp).Row).Interior.color = RGB(255, 255, 255)
        End Select
    Next ws
End Sub

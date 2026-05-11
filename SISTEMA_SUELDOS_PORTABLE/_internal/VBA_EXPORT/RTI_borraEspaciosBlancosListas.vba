Sub borrarEspaciosBlancosListas()
    
    Dim filaVacia As Integer
    Dim filaLlena As Integer
    Dim vacio As Boolean
    Dim i As Integer
    
    For fila = 2 To Hoja2.Cells(4, 21).Value
        
        vacio = False
    
        If ActiveSheet.Cells(fila, 1 + 2).Value = 0 Then
            vacio = True
            filaVacia = fila
            i = 0
            Do While vacio And i < 2000
                If ActiveSheet.Cells(fila + i, 1 + 2).Value <> 0 Then
                    vacio = False
                    filaLlena = fila + i
                    Call copiarFila(filaLlena, filaVacia)
                Else
                    i = i + 1
                End If
            Loop
        End If
    
    Next fila
    
    ActiveWindow.ScrollRow = 1
    
End Sub

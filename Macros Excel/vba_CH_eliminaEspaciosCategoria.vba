Sub eliminarEspaciosEnCategoria(fila)
    
    
    If InStr(1, ActiveSheet.Cells(fila, 2), "MEDIO OFICIAL") Then
        ActiveSheet.Cells(fila, 2).Value = "MEDIO OFICIAL"
    Else
        If InStr(1, ActiveSheet.Cells(fila, 2), "OFICIAL") Then
            ActiveSheet.Cells(fila, 2).Value = "OFICIAL"
        Else
            If InStr(1, ActiveSheet.Cells(fila, 2), "ESPECIALIZADO") Then
                ActiveSheet.Cells(fila, 2).Value = "ESPECIALIZADO"
            Else
                If InStr(1, ActiveSheet.Cells(fila, 2), "AYUDANTE") Then
                    ActiveSheet.Cells(fila, 2).Value = "AYUDANTE"
                End If
            End If
        End If
    End If
            
    
End Sub

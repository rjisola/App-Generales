Sub copiarSueldosAcordados(ByRef fila)

    ActiveSheet.Cells(fila, 19).Value = Hoja4.Cells(fila, 12).Value
    
End Sub

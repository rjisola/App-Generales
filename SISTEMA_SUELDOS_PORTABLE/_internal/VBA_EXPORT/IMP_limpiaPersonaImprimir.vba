Sub limpiarPersonaImprimir(contador, columna)

    Dim contadorFor As Integer

    For contadorFor = contador To contador + 19
    
        Hoja1.Cells(contadorFor, columna).Value = ""
        Hoja1.Cells(contadorFor, columna + 1).Value = ""
        Hoja1.Cells(contadorFor, columna + 2).Value = ""
    
    Next contadorFor
    
    

End Sub

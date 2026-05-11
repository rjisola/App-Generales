Sub colorearImprimir(contador, columna, color, desplazamiento)

    For contadorFor = contador To contador + desplazamiento
        
        If color = "26265" Or color = "13260" Then
            Hoja3.Cells(contadorFor, columna).Interior.color = color
            Hoja3.Cells(contadorFor, columna + 1).Interior.color = color
            Hoja3.Cells(contadorFor, columna + 2).Interior.color = color
        Else
            Hoja1.Cells(contadorFor, columna).Interior.color = color
            Hoja1.Cells(contadorFor, columna + 1).Interior.color = color
            Hoja1.Cells(contadorFor, columna + 2).Interior.color = color
        End If
        
    Next contadorFor

End Sub

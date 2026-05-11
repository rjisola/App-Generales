Function CalcularPosicionConFila(ByVal fila, ByVal desplazamiento) As Integer
    
    '9 es el numero de filas que tiene "inutiles" la planilla.
    
    desplazamiento = desplazamiento + 1
    
    Dim posicion As Integer
    
    If (fila - 9) Mod 2 = 0 Then
        posicion = (fila - 9) * desplazamiento + 1 - (((fila - 9) / 2) * desplazamiento)
    Else
        posicion = (fila - 9) * desplazamiento + 1 - (desplazamiento * ((fila - 9) \ 2) + desplazamiento)
    End If
    
    CalcularPosicionConFila = posicion

End Function

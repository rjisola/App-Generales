Sub eliminarEspaciosEnBlanco()
    
    Dim desplazamiento As Integer
    Dim maximoPersonas As Integer
    Dim posicionVacio As Integer
    Dim contador As Integer
    Dim posicionActual As Integer
    Dim posicionLleno As Integer
    Dim columnaLlena As Integer
    Dim columnaVacia As Integer
    Dim maxPosicion As Integer
    Dim i As Integer
    
    desplazamiento = 18
    maximoPersonas = Hoja2.Cells(4, 21).Value + 9
    maxPosicion = (maximoPersonas * desplazamiento / 2)
    For fila = 9 To maximoPersonas
        
        posicionActual = CalcularPosicionConFila(fila, desplazamiento)
        
        Call formatearQuincena(posicionActual + 11, 1, ActiveSheet)
        Call formatearQuincena(posicionActual + 11, 4, ActiveSheet)

        If ActiveSheet.Cells(6, 10) = "SI" Then
            Call acomodarAltoFilas(posicionActual)
        End If
        vacio = False
        'Si es par
        If fila Mod 2 = 0 Then
            If ActiveSheet.Cells(posicionActual, 5).Value = 0 Then
                vacio = True
                posicionVacio = posicionActual
                columnaVacia = 4
                i = 0
                Do While vacio And posicionActual <= maxPosicion
                    If ActiveSheet.Cells(posicionActual, 2).Value <> 0 And posicionVacio <> posicionActual Then
                        posicionLleno = posicionActual
                        vacio = False
                        columnaLlena = 1
                    Else
                        If ActiveSheet.Cells(posicionActual, 5).Value <> 0 Then
                            posicionLleno = posicionActual
                            vacio = False
                            columnaLlena = 4
                        Else
                            i = i + 2
                            posicionActual = CalcularPosicionConFila(fila + i, desplazamiento)
                        End If
                    End If
                Loop
                Call cortarYPegarCupon(posicionLleno, posicionVacio, columnaLlena, desplazamiento, columnaVacia)
            End If

        Else
            If ActiveSheet.Cells(posicionActual, 2).Value = 0 Then
                vacio = True
                posicionVacio = posicionActual
                columnaVacia = 1
                i = 0
                Do While vacio And posicionActual <= maxPosicion
                    If ActiveSheet.Cells(posicionActual, 2).Value <> 0 Then
                        posicionLleno = posicionActual
                        vacio = False
                        columnaLlena = 1
                    Else
                        If ActiveSheet.Cells(posicionActual, 5).Value <> 0 Then
                            posicionLleno = posicionActual
                            vacio = False
                            columnaLlena = 4
                        Else
                            i = i + 2
                            posicionActual = CalcularPosicionConFila(fila + i, desplazamiento)
                        End If
                    End If
                Loop
                Call cortarYPegarCupon(posicionLleno, posicionVacio, columnaLlena, desplazamiento, columnaVacia)
                
            End If
        End If
        
        Call formatearQuincena(posicionActual + 11, 1, ActiveSheet)
        Call formatearQuincena(posicionActual + 11, 4, ActiveSheet)
        
    Next fila
    
    ActiveWindow.ScrollRow = 1
    
End Sub

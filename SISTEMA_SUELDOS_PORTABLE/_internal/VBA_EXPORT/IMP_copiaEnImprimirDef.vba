Sub copiarEnImprimirDef(fila, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
    
    'COLORES:
    'ROJO = RGB(255,0,0)
    'VERDE = RGB(112,173,71)
    'NARANJA = RGB (255,192,0)
    'BLANCO (vacío) = xlNone / RGB(255,255,255)
    'GRIS = RGB(165,165,165)
    'AZUL = RGB(68,114,196)
    'SALMON (rosita) = RGB(252,228,214)
    'CELESTE = RGB(91,155,213)
    'TEJA  = RGB(204,51,0)
    
    Dim contador    As Integer
    Dim colorEmpleado As String
    Dim par         As Boolean
    Dim desplazamiento As Integer
    
    desplazamiento = 18
    colorEmpleado = Hoja2.Cells(fila, 1).Interior.color
    
    For contador = CalcularPosicionConFila(fila, desplazamiento) To CalcularPosicionConFila(fila, desplazamiento) + desplazamiento
        
        If fila Mod 2 = 0 Then
            par = True
        Else
            par = False
        End If
        
        'Si es VERDE
        If colorEmpleado = RGB(112, 173, 71) Then
            If par Then
                Call completarImprimirVerde(fila, contador, 4, desplazamiento, colorEmpleado)
            Else
                Call completarImprimirVerde(fila, contador, 1, desplazamiento, colorEmpleado)
            End If
        Else
            'Si es NARANJA
            If colorEmpleado = RGB(255, 192, 0) Then
                If par Then
                    Call completarImprimirNaranja(fila, contador, 4, desplazamiento, colorEmpleado)
                Else
                    Call completarImprimirNaranja(fila, contador, 1, desplazamiento, colorEmpleado)
                End If
            Else
                'Si es GRIS
                If colorEmpleado = RGB(165, 165, 165) Then
                    If par Then
                        Call completarImprimirGris(fila, contador, 4, desplazamiento, colorEmpleado)
                    Else
                        Call completarImprimirGris(fila, contador, 1, desplazamiento, colorEmpleado)
                    End If
                Else
                    'Si es AZUL
                    If colorEmpleado = RGB(68, 114, 196) Then
                        If par Then
                            Call completarImprimirAzul(fila, contador, 4, desplazamiento, colorEmpleado)
                        Else
                            Call completarImprimirAzul(fila, contador, 1, desplazamiento, colorEmpleado)
                        End If
                    Else
                    'Si es TEJA
                    If colorEmpleado = RGB(204, 51, 0) Then
                        If par Then
                            Call completarImprimirTeja(fila, contador, 4, desplazamiento, colorEmpleado)
                        Else
                            Call completarImprimirTeja(fila, contador, 1, desplazamiento, colorEmpleado)
                        End If
                    
                    Else
                        'Si es SALMON
                        If colorEmpleado = RGB(252, 228, 214) Then
                            If par Then
                                Call completarImprimirSalmon(fila, contador, 4, desplazamiento, colorEmpleado)
                            Else
                                Call completarImprimirSalmon(fila, contador, 1, desplazamiento, colorEmpleado)
                            End If
                        Else
                            'Si es BLANCO
                            If colorEmpleado = RGB(255, 255, 255) Then
                                If par Then
                                    Call completarImprimirBlanco(fila, contador, 4, desplazamiento, colorEmpleado)
                                Else
                                    Call completarImprimirBlanco(fila, contador, 1, desplazamiento, colorEmpleado)
                                End If
                            Else
                                'Si es AMARILLO
                                If colorEmpleado = RGB(255, 255, 0) Then
                                    If par Then
                                        Call completarImprimirAmarillo(fila, contador, 4, desplazamiento, colorEmpleado, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
                                    Else
                                        Call completarImprimirAmarillo(fila, contador, 1, desplazamiento, colorEmpleado, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
                                    End If
                                Else
                                    'Si es CELESTE
                                    If colorEmpleado = RGB(91, 155, 213) Then
                                        If par Then
                                            Call completarImprimirCeleste(fila, contador, 4, desplazamiento, colorEmpleado)
                                        Else
                                            Call completarImprimirCeleste(fila, contador, 1, desplazamiento, colorEmpleado)
                                        End If
                                    Else
                                        'Si es MARRON
                                        If colorEmpleado = RGB(153, 102, 0) Then
                                            If par Then
                                                Call completarImprimirMarron(fila, contador, 4, desplazamiento, colorEmpleado)
                                            Else
                                                Call completarImprimirMarron(fila, contador, 1, desplazamiento, colorEmpleado)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
        contador = contador + desplazamiento
        
    Next contador
    
End Sub
Sub MacroDefinitivaCalcularHoras()
    
    'COLORES:
    'ROJO = RGB(255,0,0)
    'VERDE = RGB(112,173,71)
    'NARANJA = RGB (255,192,0)
    'BLANCO (vacío) = xlNone
    'GRIS = RGB(165,165,165)
    'AZUL = RGB(68,114,196)
    'SALMON (rosita) = RGB(252,228,214)
    'AMARILLO = RGB(255,255,0)
    'CELESTE = RGB(91,155,213)
    'TEJA = RGB(204,51,0)
    
    Dim horas       As Single
    Dim horasAlCien As Single
    Dim horasAlCincuenta As Single
    Dim horasNormales As Single
    Dim Dia         As String
    Dim columna     As Long
    Dim fila        As Long
    Dim maximoDias  As Integer
    Dim inicioPersonas As Integer
    Dim maximoPersonas As Integer
    Dim presentismo As Boolean
    Dim colorCategoria As String
    Dim categoria   As String
    Dim feriado     As Boolean
    Dim horasQuilmesCincuenta As Single
    Dim horasPapeleraCincuenta As Single
    Dim horasQuilmesCien As Single
    Dim horasPapeleraCien As Double
    
    If Hoja2.Cells(7, 23).Value = "SI" Then
        Hoja1.Columns("A:F").ClearContents
        Hoja1.Columns("A:F").Interior.color = RGB(255, 255, 255)
        Hoja3.Columns("A:F").ClearContents
        Hoja3.Columns("A:F").Interior.color = RGB(255, 255, 255)
    End If
    
    horasAlCien = 0
    horasAlCincuenta = 0
    horasNormales = 0
    horas = 0
    
    inicioDias = ActiveSheet.Range("u1").Value
    
    maximoDias = ActiveSheet.Range("u2").Value
    
    inicioPersonas = ActiveSheet.Range("u3").Value
    
    maximoPersonas = ActiveSheet.Range("u4").Value
    
    Call eliminarError
    
    For fila = 8 + inicioPersonas To maximoPersonas + 8
        
        horasQuilmesCincuenta = 0
        horasPapeleraCincuenta = 0
        horasQuilmesCien = 0
        horasPapeleraCien = 0
        
        If ActiveSheet.Cells(fila, 1).Value <> vbNullString Then
            
            ActiveSheet.Cells(fila, 20).Value = 0
            ActiveSheet.Cells(fila, 21).Value = 0
            ActiveSheet.Cells(fila, 22).Value = 0
            ActiveSheet.Cells(fila, 23).Value = 0
            ActiveSheet.Cells(fila, 24).Value = 0
            ActiveSheet.Cells(fila, 25).Value = 0
            ActiveSheet.Cells(fila, 26).Value = 0
            ActiveSheet.Cells(fila, 27).Value = 0
            ActiveSheet.Cells(fila, 28).Value = 0
            ActiveSheet.Cells(fila, 29).Value = 0
            ActiveSheet.Cells(fila, 30).Value = 0
            
            presentismo = True
        End If
        
        For columna = 2 + inicioDias To maximoDias + 2
            
            If ActiveSheet.Cells(fila, 1).Value <> vbNullString Then
                
                If ActiveSheet.Cells(7, columna) <> vbNullString Then
                    feriado = True
                Else
                    feriado = False
                End If
                
                colorCategoria = ActiveSheet.Cells(fila, 1).Interior.color
                Dia = ActiveSheet.Cells(8, columna)
                
                'Para los que tienen algo en "NO CONSIDERAR" no hace nada.
                If Not IsEmpty(Hoja2.Cells(fila, 33)) Then
                Else
                    'Para los VERDES
                    If colorCategoria = RGB(112, 173, 71) Then
                        
                        Call generarHorasVerde(fila, columna, Dia, feriado, unificarDatosVerde(fila, columna, Dia))
                    Else
                        'Para los NARANJAS
                        If colorCategoria = RGB(255, 192, 0) Then
                            Call generarHorasNaranja(fila, columna, Dia, presentismo, feriado, unificarDatosNaranja(fila, columna, Dia))
                        Else
                            'Para los VACIOS o BLANCOS
                            If colorCategoria = RGB(255, 255, 255) Then
                                Call generarHorasBlanco(fila, columna, Dia, presentismo, feriado, unificarDatosBlanco(fila, columna, Dia))
                            Else
                                'Para los GRISES
                                If colorCategoria = RGB(165, 165, 165) Then
                                    Call copiarSueldosAcordados(fila)
                                    Call generarHorasGris(fila, columna, Dia, presentismo, feriado, unificarDatosGris(fila, columna, Dia))
                                Else
                                    'Para los AZULES
                                    If colorCategoria = RGB(68, 114, 196) Then
                                    Call copiarSueldosAcordados(fila)
                                        Call generarHorasAzul(fila, columna, Dia, feriado, unificarDatosCeleste(fila, columna, Dia))
                                    Else
                                        'Para los TEJA
                                        If colorCategoria = RGB(204, 51, 0) Then
                                        Call copiarSueldosAcordados(fila)
                                            Call generarHorasTeja(fila, columna, Dia, feriado, unificarDatosTeja(fila, columna, Dia))
                                        Else
                                            'Para los SALMON
                                            If colorCategoria = RGB(252, 228, 214) Then
                                                Call generarHorasSalmon(fila, columna, Dia, presentismo, feriado, unificarDatosSalmon(fila, columna, Dia))
                                            Else
                                                'Para los AMARILLO
                                                If colorCategoria = RGB(255, 255, 0) Then
                                                    Call generarHorasAmarillo(fila, columna, Dia, presentismo, feriado, unificarDatosAmarillo(fila, columna, Dia), horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
                                                Else
                                                    'Para los CELESTE
                                                    If colorCategoria = RGB(91, 155, 213) Then
                                                    Call copiarSueldosAcordados(fila)
                                                        Call generarHorasCeleste(fila, columna, Dia, feriado, unificarDatosCeleste(fila, columna, Dia))
                                                    Else
                                                        'Para los MARRONES
                                                        If colorCategoria = RGB(153, 102, 0) Then
                                                            Call generarHorasMarron(fila, columna, Dia, presentismo, feriado, unificarDatosMarron(fila, columna, Dia))
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
            End If
        Next columna
        
        If ActiveSheet.Cells(fila, 1).Value <> vbNullString Then
            
            eliminarEspaciosEnCategoria (fila)
            categoria = ActiveSheet.Cells(fila, 2).Value
            
            If colorCategoria = RGB(112, 173, 71) Then
                Call calcularImporteVerde(fila)
            End If
            If colorCategoria = RGB(255, 192, 0) Then
                Call calcularImporteNaranja(fila, maximoDias, categoria)
            End If
            If colorCategoria = RGB(255, 255, 255) Then
                Call calcularImporteBlanco(fila, presentismo, categoria)
            End If
            If colorCategoria = RGB(165, 165, 165) Then
                Call calcularImporteGris(fila, categoria)
            End If
            If colorCategoria = RGB(68, 114, 196) Then
                Call calcularImporteAzul(fila)
            End If
            If colorCategoria = RGB(204, 51, 0) Then
                Call calcularImporteTeja(fila)
            End If
            If colorCategoria = RGB(252, 228, 214) Then
                Call calcularImporteSalmon(fila, presentismo, categoria)
            End If
            If colorCategoria = RGB(255, 255, 0) Then
                Call calcularImporteAmarillo(fila, presentismo, categoria, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
            End If
            If colorCategoria = RGB(91, 155, 213) Then
                Call calcularImporteCeleste(fila)
            End If
            If colorCategoria = RGB(153, 102, 0) Then
                Call calcularImporteMarron(fila, columna, categoria, Dia, presentismo, feriado)
            End If
            If Hoja2.Cells(7, 23).Value = "SI" Then
                Call copiarEnImprimirDef(fila, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
            End If
        End If
    Next fila
    
End Sub

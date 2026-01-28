'COLORES:
'VERDE = RGB(112,173,71)
'NARANJA = RGB =(255,192,0)
'AZUL = RGB(68,114,196)

Sub generarHorasContador(fila, columna, horas, colorHoras, Dia)

    If IsNumeric(horas) Then

        If Dia = "domingo" Then
        Else
            If Dia = "sábado" Then
            Else
                If Dia = "viernes" Then
                'BLANCO
                    If colorHoras = RGB(255, 255, 255) Then
                        If horas <= 8 Then
                            Hoja9.Cells(fila, 6).Value = Hoja9.Cells(fila, 6).Value + horas
                        Else
                            Hoja9.Cells(fila, 6).Value = Hoja9.Cells(fila, 6).Value + 8
                        End If
                    Else
                        'VERDE(papelera)
                        If colorHoras = RGB(112, 173, 71) Or colorHoras = RGB(153, 102, 0) Then
                            If horas <= 8 Then
                                Hoja9.Cells(fila, 8).Value = Hoja9.Cells(fila, 8).Value + horas
                            Else
                                Hoja9.Cells(fila, 8).Value = Hoja9.Cells(fila, 8).Value + 8
                            End If
                        Else
                            'NARANJA(quilmes)
                            If colorHoras = RGB(255, 192, 0) Then
                                If horas <= 8 Then
                                    Hoja9.Cells(fila, 7).Value = Hoja9.Cells(fila, 7).Value + horas
                                Else
                                    Hoja9.Cells(fila, 7).Value = Hoja9.Cells(fila, 7).Value + 8
                                End If
                            End If
                        End If
                    End If
                
                Else
                    'BLANCO
                    If colorHoras = RGB(255, 255, 255) Then
                        If horas <= 9 Then
                            Hoja9.Cells(fila, 6).Value = Hoja9.Cells(fila, 6).Value + horas
                        Else
                            Hoja9.Cells(fila, 6).Value = Hoja9.Cells(fila, 6).Value + 9
                        End If
                    Else
                        'MARRON(papelera)
                        If colorHoras = RGB(112, 173, 71) Or colorHoras = RGB(153, 102, 0) Then
                            If horas <= 9 Then
                                Hoja9.Cells(fila, 8).Value = Hoja9.Cells(fila, 8).Value + horas
                            Else
                                Hoja9.Cells(fila, 8).Value = Hoja9.Cells(fila, 8).Value + 9
                            End If
                        Else
                            'NARANJA(quilmes)
                            If colorHoras = RGB(255, 192, 0) Then
                                If horas <= 9 Then
                                    Hoja9.Cells(fila, 7).Value = Hoja9.Cells(fila, 7).Value + horas
                                Else
                                    Hoja9.Cells(fila, 7).Value = Hoja9.Cells(fila, 7).Value + 9
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    
End Sub

Function unificarDatosNaranja(fila, columna, Dia) As Single
    
    Dim horas As String
    
    horas = ActiveSheet.Cells(fila, columna).Value
    
    'Si CORTARON
    If horas = "CORTARON" Or horas = "cortaron" Then
        unificarDatosNaranja = 0
    Else
        'Si lluvia
        If horas = "LLUVIA" Or horas = "lluvia" Then
            unificarDatosNaranja = 2.5
        Else
            
            'Si VACACIONES o NO o AVISO
            If horas = "NO" Or horas = "vacaciones" Or horas = "VACACIONES" Or horas = "c/aviso" Or horas = "C/AVISO" Or horas = "C/A" Or horas = "c/a" Or horas = "ENFERMO" Or horas = "enfermo" Or horas = "ART" Then
                unificarDatosNaranja = 0
            Else
            'Si FALLECIMIENTO
            If horas = "FALLEC" Or horas = "fallec " Then
                unificarDatosNaranja = 0
            
               Else
                'Si FALTO
                If horas = "falto" Or horas = "FALTO" Then
                    unificarDatosNaranja = -1
                Else
                
                    'Si ENFERMO
                    If horas = "ENFERMO" Then
                        If Dia = "sábado" Then
                            unificarDatosNaranja = -1
                        Else
                            unificarDatosNaranja = -1
                        End If
                    Else
                            
                        'Si presentan CERTIFICADO
                        If horas = "certif" Or horas = "CERTIF" Or horas = "CERT" Or horas = "cert" Then
                            unificarDatosNaranja = -1
                            'Si la celda está VACÍA
                        Else
                            If ActiveSheet.Cells(fila, columna) = vbNullString Then
                                ActiveSheet.Cells(fila, columna) = 0
                                unificarDatosNaranja = 0
                    
                            Else
                                If ActiveSheet.Cells(fila, columna) >= 0 Or ActiveSheet.Cells(fila, columna) <= 24 Then
                                    unificarDatosNaranja = horas
                                Else
                                    Call informarError
                                End If
                    
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    End If
End Function

Function unificarDatosVerde(fila, columna, Dia) As Single
    
    Dim horas As String
    
    horas = ActiveSheet.Cells(fila, columna).Value
    
    'Si CORTARON
    If horas = "LLUVIA" Or horas = "lluvia" Then
        unificarDatosVerde = 2.5
    Else
        If horas = "CORTARON" Or horas = "cortaron" Then
            unificarDatosVerde = 0
        Else
            'Si VACACIONES o NO o AVISO
            If horas = "NO" Or horas = "vacaciones" Or horas = "VACACIONES" Or horas = "c/aviso" Or horas = "C/AVISO" Or horas = "C/A" Or horas = "c/a" Or horas = "ART" Then
                unificarDatosVerde = 0
            Else
                'Si FALTO
                If horas = "falto" Or horas = "FALTO" Then
                     If Dia = "sábado" Then
                     unificarDatosVerde = 0
                     End If
                           
                    unificarDatosVerde = 0
                Else
                    'Si presentan CERTIFICADO
                    If horas = "certif" Or horas = "CERTIF" Or horas = "CERT" Or horas = "cert" Or horas = "ENFERMO" Or horas = "enfermo" Then
                        If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
                            unificarDatosVerde = 0
                            'Para el dia SABADO:
                        Else
                            If Dia = "sábado" Then
                                unificarDatosVerde = 0
                            End If
                        End If
                        'Si la celda está VACÍA
                    Else
                        If ActiveSheet.Cells(fila, columna) = vbNullString Then
                            ActiveSheet.Cells(fila, columna).Value = 0
                            unificarDatosVerde = 0
                        Else
                            If ActiveSheet.Cells(fila, columna) >= 0 Or ActiveSheet.Cells(fila, columna) <= 24 Then
                                unificarDatosVerde = horas
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

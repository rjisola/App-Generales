Function unificarDatosMarron(fila, columna, Dia) As Single
    
    Dim horas As String
    
    
    horas = ActiveSheet.Cells(fila, columna).Value
    
    'Si LLUVIA
    If horas = "LLUVIA" Or horas = "lluvia" Then
        unificarDatosMarron = 2.5
    Else
        'Si CORTARON
        If horas = "CORTARON" Or horas = "cortaron" Then
            unificarDatosMarron = 0
        Else
            'Si VACACIONES  o AVISO
            If horas = "vacaciones" Or horas = "VACACIONES" Or horas = "c/aviso" Or horas = "C/AVISO" Or horas = "C/A" Or horas = "c/a" Or horas = "ART" Then
                unificarDatosMarron = 0
            Else
                
                'Si FALTO
                If horas = "falto" Or horas = "FALTO" Then
                    unificarDatosMarron = -1
                Else
                    'Si ENFERMO
                    If horas = "ENFERMO" Or horas = "enfermo" Or horas = "certif" Or horas = "CERTIF" Or horas = "cert" Or horas = "CERT" Or horas = "ENFERMO" Or horas = "enfermo" Then
                        unificarDatosMarron = -1
                    Else
             
                        'Si la celda está VACÍA
           
                        If ActiveSheet.Cells(fila, columna) = vbNullString Then
                            ActiveSheet.Cells(fila, columna).Value = 0
                            unificarDatosMarron = 0
                    
                        Else
                    
                            If ActiveSheet.Cells(fila, columna) >= 0 Or ActiveSheet.Cells(fila, columna) <= 24 Then
                                unificarDatosMarron = horas
                            Else
                                Call informarError
                            End If
                    
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function


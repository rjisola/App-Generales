Function unificarDatosGris(fila, columna, Dia) As Single

    
    Dim horas As String
    
    horas = ActiveSheet.Cells(fila, columna).Value
    
    
    
    'Si CORTARON
    If horas = "LLUVIA" Or horas = "lluvia" Then
        unificarDatosGris = 2.5
    Else
        If horas = "CORTARON" Or horas = "cortaron" Then
        
            unificarDatosGris = 0
        Else
            'Si VACACIONES o NO o AVISO
            If horas = "NO" Or horas = "VAC" Or horas = "VACACIONES" Or horas = "c/aviso" Or horas = "C/AVISO" Or horas = "C/A" Or horas = "c/a" Or horas = "ENFERMO" Or horas = "enfermo" Or horas = "ART" Then
                unificarDatosGris = 0
            Else
                
                'Si FALTO
                If horas = "falto" Or horas = "FALTO" Then
                    unificarDatosGris = 0
                Else
                            
                    'Si presentan CERTIFICADO
                    If horas = "certif" Or horas = "CERTIF" Or horas = "cert" Or horas = "CERT" Then
                        unificarDatosGris = 0
                        'Si la celda está VACÍA
                    Else
                        If ActiveSheet.Cells(fila, columna) = vbNullString Then
                            ActiveSheet.Cells(fila, columna) = 0
                            unificarDatosGris = 0
                    
                        Else
                    
                            If ActiveSheet.Cells(fila, columna) >= 0 Or ActiveSheet.Cells(fila, columna) <= 24 Then
                                unificarDatosGris = horas
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

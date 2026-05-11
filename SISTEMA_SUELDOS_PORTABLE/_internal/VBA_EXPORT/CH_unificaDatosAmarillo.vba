
Function unificarDatosAmarillo(fila, columna, Dia) As Single
    
    Dim horas As String
    
    
    horas = ActiveSheet.Cells(fila, columna).Value
    
    'Si LLUVIA
    If horas = "LLUVIA" Or horas = "lluvia" Then
        unificarDatosAmarillo = 2.5
    Else
        'Si CORTARON
        If horas = "CORTARON" Or horas = "cortaron" Then
            unificarDatosAmarillo = 0
        Else
            'Si VACACIONES  o AVISO
            If horas = "vacaciones" Or horas = "VACACIONES" Or horas = "c/aviso" Or horas = "C/AVISO" Or horas = "C/A" Or horas = "c/a" Or horas = "ART" Or horas = "cortado" Or horas = "SIN HORAS" Then
                unificarDatosAmarillo = 0
            Else
                
                'Si FALTO
                If horas = "falto" Or horas = "FALTO" Then
                                                       
                    unificarDatosAmarillo = -1
                Else
                    'Si ENFERMO
                    If horas = "ENFERMO" Or horas = "enfermo" Or horas = "certif" Or horas = "CERTIF" Or horas = "cert" Or horas = "CERT" Or horas = "ENFERMO" Or horas = "enfermo" Then
                        unificarDatosAmarillo = -1
                    Else
             
                        'Si la celda está VACÍA
           
                        If ActiveSheet.Cells(fila, columna) = vbNullString Then
                            ActiveSheet.Cells(fila, columna).Value = 0
                            unificarDatosAmarillo = 0
                    
                        Else
                    
                            If ActiveSheet.Cells(fila, columna) >= 0 Or ActiveSheet.Cells(fila, columna) <= 24 Then
                                unificarDatosAmarillo = horas
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

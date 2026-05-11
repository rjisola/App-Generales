Sub generarHorasTeja(fila, columna, Dia, feriado, ByRef horas)
    
    'TEJA:
    'de lunes a viernes: 12 hs normales, >12 al 50%
    'sabados: 4hs al cincuenta, >5 al 100%
    'domingos: al 100%
    'PRESENTISMO: Siempre presente
    'CERTIF: No suma
    
    Dim horasAlCien As Single
    Dim horasAlCincuenta As Single
    Dim horasNormales As Single
    Dim horasFeriado As Single
    
    horasAlCien = 0
    horasAlCincuenta = 0
    horasNormales = 0
     horasFeriado = 0
    
    If feriado Then
        If horas <= -1 Or horas > 24 Then
            If horas = -1 Then
                If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
                    horasNormales = 12
                Else
                    If Dia = "sábado" Then
                        horasNormales = 5
                    End If
                End If
            Else
                
                Call informarError
            End If
        Else
            horasFeriado = horas
        End If
        
    Else
        If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
            If horas <= -1 Or horas > 24 Then
                If horas = -1 Then
                    horasNormales = 0
                Else
                    
                    Call informarError
                End If
            Else
                If horas <= 12 Then
                    horasNormales = horas
                Else
                    If horas > 12 Then
                        horasNormales = 12
                        horasAlCincuenta = horas - horasNormales
                    End If
                End If
            End If
        End If
        
        If Dia = "sábado" Then
            If horas < 0 Or horas > 24 Then
                If horas = -1 Then
                Else
                    Call informarError
                End If
            Else
                If horas <= 4 Then
                    
                    horasAlCincuenta = horas + horasAlCincuenta
                    
                End If
                
                If horas > 4 Then
                horasAlCincuenta = 4
                    horasAlCien = horas - horasAlCincuenta + horasAlCien
                    
                End If
            End If
        Else
            If Dia = "domingo" Then
                If horas >= 0 And horas <= 24 Then
                    horasAlCien = horas
                End If
            End If
        End If
    End If
    
    ActiveSheet.Cells(fila, 20).Value = ActiveSheet.Cells(fila, 20).Value + horasNormales
    ActiveSheet.Cells(fila, 21).Value = ActiveSheet.Cells(fila, 21).Value + horasAlCincuenta
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
    
    ActiveSheet.Cells(fila, 24).Value = "-"
    
End Sub
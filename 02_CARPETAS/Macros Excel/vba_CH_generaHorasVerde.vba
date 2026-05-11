Sub generarHorasVerde(fila, columna, Dia, feriado, ByRef horas)

    'VERDE:
    'de lunes a viernes: 12 hs normales, >12 al 100%
    'sabados: 5hs normales, >5 al 100%
    'domingos: No trabaja
    'PRESENTISMO: Siempre presente
    'CERTIF: No suma
    

    Dim horasAlCien As Single
    Dim horasNormales As Single
    Dim horasFeriado As Single
    
    horasAlCien = 0
    horasNormales = 0
    horasFeriado = 0
    
    presentismo = True
     
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
                    presentismo = False
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
                        horasAlCien = horas - horasNormales
                    End If
                End If
            End If
        Else
            If Dia = "sábado" Then
                If horas < 0 Or horas > 24 Then
                    If horas = -1 Then
                        horasNormales = 5
                        presentismo = False
                    Else
                       
                        Call informarError
                    End If
                Else
                    If horas <= 5 Then
                        horasNormales = horas
                    End If
                    If horas > 5 Then
                        horasNormales = 5
                        horasAlCien = horas - horasNormales
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
    End If
    
    ActiveSheet.Cells(fila, 20).Value = ActiveSheet.Cells(fila, 20).Value + horasNormales
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
    
    If presentismo Then
        ActiveSheet.Cells(fila, 24).Value = "PRESENTISMO"
    Else
        ActiveSheet.Cells(fila, 24).Value = "Pierde PRES."
    End If

End Sub

Sub generarHorasSalmon(fila, columna, Dia, ByRef presentismo, feriado, ByRef horas)

    'SALMON:
    'de lunes a viernes: 8 hs normales, >8 al 50%
    'sabados: 4hs normales, >4 1 al 50% y el resto al 100%
    'domingos: No trabaja, si trabaja son al 100%
    'feriados: 100%
    'PRESENTISMO: Si falta lo pierde.
    'con certificado cobra horas normales y pierde pres.
    'con aviso no se le paga el dia pero no pierde el pres.
    'Le cuentan las horas normales

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
            If horas = -1 Or horas = -8 Or horas = -9 Or horas = -4 Then
                If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Then
                    horasNormales = 9
                End If
                If Dia = "viernes" Then
                    horasNormales = 8
                End If
                If Dia = "sábado" Then
                    horasNormales = 4
                End If
            End If
        Else
            horasFeriado = horas
        End If
    
    Else
        If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Then
            If horas < 0 Or horas > 24 Then
                If horas = -1 Then
                    presentismo = False
                Else
                    If horas = -9 Then
                        presentismo = False
                        horasNormales = 9
                    Else
                        Call informarError
                    End If
                End If
            End If
        
            If horas <= 9 And horas > 0 Then
                horasNormales = horas
            End If
        
            If horas > 9 And horas <= 24 Then
                horasNormales = 9
                horasAlCincuenta = horas - horasNormales
            End If
            
            
        End If
        
        If Dia = "viernes" Then
        
        If horas < 0 Or horas > 24 Then
                If horas = -1 Then
                    presentismo = False
                Else
                    If horas = -8 Then
                        presentismo = False
                        horasNormales = 8
                    Else
                        Call informarError
                    End If
                End If
            End If

            If horas <= 8 And horas > 0 Then
                horasNormales = horas
            End If
        
            If horas > 8 Then
                horasNormales = 8
                horasAlCincuenta = horas - horasNormales
            End If
            
            
            
        End If
    
        If Dia = "sábado" Then
            
            If horas = -1 Then
                presentismo = False
                horasNormales = 0
                horasAlCincuenta = 0
                horasAlCien = 0
            Else
            If horas <= 4 Then
                horasNormales = 0
                horasAlCincuenta = horas
            Else
                If horas > 4 Then
                    horasNormales = 0
                    horasAlCincuenta = 4
                    horasAlCien = horas - horasAlCincuenta
                Else
                    Call informarError
                End If
            End If
        End If
        End If
        
        If Dia = "domingo" Then
            If horas = -8 Or horas = -1 Then
                horasNormales = 0
                horasAlCincuenta = 0
                horasAlCien = 0
            Else
                horasAlCien = horas
            End If
        End If
    End If
        
    ActiveSheet.Cells(fila, 20).Value = ActiveSheet.Cells(fila, 20).Value + horasNormales
    ActiveSheet.Cells(fila, 21).Value = ActiveSheet.Cells(fila, 21).Value + horasAlCincuenta
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
        
    If presentismo Then
        ActiveSheet.Cells(fila, 24).Value = "PRESENTISMO"
    Else
        ActiveSheet.Cells(fila, 24).Value = "pierde PRES."
    End If
End Sub

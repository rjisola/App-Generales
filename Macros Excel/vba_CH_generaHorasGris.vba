Sub generarHorasGris(fila, columna, Dia, ByRef presentismo, feriado, ByRef horas)


    'GRIS:
    'de lunes a jueves: 9 hs normales, 9><12 al 50%, >12 100%
    'vienes: 8hs normales, 8><12 al 50%, >12 100%
    'sabados: 4hs 50%, >4 al 100%
    'domingos: No trabaja
    'feriado: 100%
    'PRESENTISMO: Si falta lo pierde.
    'Con certificado cobra horas normalmente y pierde pres.

    Dim horasAlCien As Single
    Dim horasAlCincuenta As Single
    Dim horasNormales As Single
    Dim horasFeriado As Single
    Dim vac As Single

    horasAlCien = 0
    horasAlCincuenta = 0
    horasNormales = 0
    horasFeriado = 0
    
    
    
    If feriado Then
        
        If horas <= -1 Or horas > 24 Then
            If horas = -1 Then
                If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Then
                    horasNormales = 9
                End If
                If Dia = "viernes" Then
                    horasNormales = 8
                End If
                If Dia = "sábado" Then
                    horasAlCincuenta = 4
                End If
            Else
                Call informarError
               
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
                    If horas = -8 Then
                        horasNormales = 8
                        presentismo = False
                    Else
                       
                        Call informarError
                    End If
                End If
            End If
        
            If horas <= 9 And horas > 0 Then
            
                horasNormales = horas
            
            End If
        
            If horas > 9 Then
            
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
                        horasNormales = 8
                        presentismo = False
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
    
            If horas < 0 Or horas > 24 Then
        
                If horas = -1 Then
                Else
                    If horas = -8 Then
                        horasNormales = 8
                        presentismo = False
                    Else
        
                       
                        Call informarError
                    End If
                End If
            End If
        
            If horas <= 4 And horas > 0 Then
            
                horasAlCincuenta = horas
            
            End If
        
            If horas > 4 Then
            
                horasAlCincuenta = 4
                horasAlCien = horas - horasAlCincuenta
            
            End If
        End If
    End If
    
    If Dia = "domingo" Or Dia = "feriado" Then
        horasAlCien = horas
    End If
    If Dia = "feriado" Then
        horasFeriado = horas
    End If
    
    ActiveSheet.Cells(fila, 20).Value = ActiveSheet.Cells(fila, 20).Value + horasNormales
    ActiveSheet.Cells(fila, 21).Value = ActiveSheet.Cells(fila, 21).Value + horasAlCincuenta
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
        
    If presentismo Then
        ActiveSheet.Cells(fila, 24).Value = "-"
    Else
        ActiveSheet.Cells(fila, 24).Value = "-"
    End If
End Sub

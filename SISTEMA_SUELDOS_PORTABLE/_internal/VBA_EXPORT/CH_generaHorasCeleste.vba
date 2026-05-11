Sub generarHorasCeleste(fila, columna, Dia, feriado, ByRef horas)

    'CELESTE:
    'de lunes a viernes: 12 hs normales, >12 al 50%
    'sabados: 5hs normales, >5 al 100%
    'domingos: No trabaja
    'PRESENTISMO: Siempre presente
    'CERTIF: No suma
    
    Dim horasAlCien As Single
    Dim horasAlCincuenta As Single
    Dim horasNormales As Single
    Dim horasFeriado As Single
    Dim apellido As String
    
    horasAlCien = 0
    horasAlCincuenta = 0
    horasNormales = 0
    horasFeriado = 0
    apellido = Hoja2.Cells(fila, 1)
          
    If feriado Then
        If horas <= -1 Or horas > 24 Then
            If horas = -1 Then
                If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
                    horasNormales = 10
                Else
                    If Dia = "sábado" Then
                        horasAlCincuenta = 5
                    End If
                End If
            Else
                Call informarError
            End If
        Else
            horasFeriado = horas
        End If
        
    Else ' NO ES FERIADO
        If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
            If horas <= -1 Or horas > 24 Then
                If horas = -1 Then
                    horasNormales = 0
                Else
                    Call informarError
                End If
            Else
                If horas <= 10 Then
                    horasNormales = horas
                Else
                    ' --- CORRECCIÓN AQUÍ ---
                    If apellido = "Ferreyra David Ismael" Then
                        ' Si es Ferreyra y trabaja más de 10 hs, se clava en 10 normales y 0 extras
                        horasNormales = 10
                        horasAlCincuenta = 0
                    Else
                        ' Para el resto de los empleados: 10 normales y el resto al 50%
                        horasNormales = 10
                        horasAlCincuenta = horas - 10
                    End If
                    ' -------------------------
                End If
            End If
        Else
            If Dia = "sábado" Then
                If horas < 0 Or horas > 24 Then
                    If horas = -1 Then
                        ' No hace nada
                    Else
                        Call informarError
                    End If
                Else
                    If horas <= 5 Then
                        horasAlCincuenta = horas
                    Else
                        horasAlCincuenta = 5
                        horasAlCien = horas - 5
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
    ActiveSheet.Cells(fila, 21).Value = ActiveSheet.Cells(fila, 21).Value + horasAlCincuenta
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
        
    ActiveSheet.Cells(fila, 24).Value = "-"

End Sub
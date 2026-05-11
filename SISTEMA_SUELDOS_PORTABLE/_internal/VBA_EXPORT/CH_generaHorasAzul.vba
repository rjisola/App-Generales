

Sub generarHorasAzul(fila, columna, Dia, feriado, ByRef horas)

    'AZUL:
    'de lunes a viernes: 12 hs normales, >12 al 50%
    'sábados: 5hs normales, >5 al 100%
    'domingos: No trabaja
    'PRESENTISMO: Siempre presente
    'CERTIF: No suma

    Dim horasAlCien As Single
    Dim horasAlCincuenta As Single
    Dim horasNormales As Single
    Dim horasFeriado As Single
    Dim apellido As String
    
    apellido = Hoja2.Cells(fila, 1)
    
    horasAlCien = 0
    horasAlCincuenta = 0
    horasNormales = 0
    horasFeriado = 0
    
    If feriado Then
        If horas = -1 Then
            If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
                horasNormales = 12
            ElseIf Dia = "sábado" Then
                horasNormales = 5
            End If
        ElseIf horas >= 0 And horas <= 24 Then
            horasFeriado = horas
        Else
            Call informarError
        End If
    Else
        If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
            If horas = -1 Then
                horasNormales = 0
            ElseIf horas >= 0 And horas <= 24 Then
                If horas <= 10 Then
                    horasNormales = horas
                ElseIf apellido = "Albornoz Claudio Gera" Then
                    horasNormales = horas
                If horas > 12 Then
                    horasNormales = 12
                    horasAlCincuenta = horas - horasNormales
                End If
                ElseIf horas > 10 Then
                    horasNormales = 10
                    horasAlCincuenta = horas - horasNormales
                End If
            Else
                Call informarError
            End If
        ElseIf Dia = "sábado" Then
            If horas = -1 Then
                ' No hacer nada, horasNormales ya es 0
             
            ElseIf horas >= 0 And horas <= 24 Then
                If horas > 5 Then
                    horasAlCincuenta = 5
                    horasAlCien = horas - horasAlCincuenta
                Else
                    horasAlCincuenta = horas
                End If
            Else
                Call informarError
            End If
        ElseIf Dia = "domingo" Then
            If horas >= 0 And horas <= 24 Then
                horasAlCien = horas
            Else
                Call informarError
            End If
        End If
    End If

    ActiveSheet.Cells(fila, 20).Value = ActiveSheet.Cells(fila, 20).Value + horasNormales
    ActiveSheet.Cells(fila, 21).Value = ActiveSheet.Cells(fila, 21).Value + horasAlCincuenta
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
    ActiveSheet.Cells(fila, 24).Value = "-"

End Sub



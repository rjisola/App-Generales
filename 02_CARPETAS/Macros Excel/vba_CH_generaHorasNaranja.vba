Sub generarHorasNaranja(fila, columna, Dia, ByRef presentismo, feriado, ByRef horas)

    'NARANJA:
    '150 $ por hora, 180 $ si tienen presentismo.
    'Feriado se paga según categoría
    
    Dim horasNormales As Single
    Dim horasFeriado As Single
    Dim horasAlCien As Single
    
    
    horasNormales = 0
    horasAlCien = 0
    horasFeriado = 0
    'Si es feriado
    If Not IsEmpty(Hoja2.Cells(7, columna)) Then
        If horas <> 0 Then
        horasAlCien = horas - horasNormales
        End If
    horasFeriado = 8
    End If
    
   If IsEmpty(Hoja2.Cells(7, columna)) Then
    If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Or Dia = "viernes" Then
        If horas <= -1 Or horas > 24 Then
            If horas = -12 Then
                presentismo = False
            Else
                If horas = -1 Then
                    presentismo = False
                Else
                    Call informarError
                End If
            End If
        Else
            horasNormales = horas
        End If
    End If
    End If
    If IsEmpty(Hoja2.Cells(7, columna)) Then
    If Dia = "sábado" Then
        If horas <= -1 Or horas > 24 Then
            If horas = -12 Then
                presentismo = False
            Else
                If horas = -1 Then
                    presentismo = False
                Else
                    Call informarError
                End If
            End If
        Else
            If horas <= 12 Then
                horasNormales = horas
            Else
                horasNormales = horas
                
            End If
            
        End If
    End If
    
    
    If Dia = "domingo" Then
        If horas <= -1 Or horas > 24 Then
            If horas = -12 Then
                presentismo = False
            Else
                If horas = -1 Then
                    presentismo = False
                Else
                    Call informarError
                End If
            End If
        Else
            horasNormales = horas
        End If
    End If
    End If

    ActiveSheet.Cells(fila, 20).Value = ActiveSheet.Cells(fila, 20).Value + horasNormales
    
    ActiveSheet.Cells(fila, 22).Value = ActiveSheet.Cells(fila, 22).Value + horasAlCien
    
    ActiveSheet.Cells(fila, 23).Value = ActiveSheet.Cells(fila, 23).Value + horasFeriado
    
    If presentismo Then
        ActiveSheet.Cells(fila, 24).Value = "PRESENTISMO"
    Else
        ActiveSheet.Cells(fila, 24).Value = "Pierde PRES"
    End If
    
   
   
End Sub


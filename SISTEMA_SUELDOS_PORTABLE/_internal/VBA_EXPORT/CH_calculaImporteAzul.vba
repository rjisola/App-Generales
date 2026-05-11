Sub calcularImporteAzul(fila As Long)
    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim importeHoraFeriado As Double
    Dim total As Double
    
    Dim horasAlCincuenta As Double
    Dim horasAlCien As Double
    Dim horasFeriado As Double
    Dim apellido As String
    
    ' Obtener valores de las celdas
    horasAlCincuenta = Hoja2.Cells(fila, 21).Value
    horasAlCien = Hoja2.Cells(fila, 22).Value
    horasFeriado = Hoja2.Cells(fila, 23).Value
    apellido = Hoja2.Cells(fila, 1).Value
    
    ' Calcular valores por hora estándar
    valorHoraAlCincuenta = Hoja4.Cells(fila, 12).Value / 100
    valorHoraFeriado = Hoja4.Cells(fila, 12).Value / 110 * 2
    valorHoraAlCien = Hoja4.Cells(fila, 12).Value / 110 * 2
    
    ' Ajustes especiales para ciertos empleados
    Select Case apellido
        Case "Holgado Pedro Atilio", "Souza Edgardo Andres", "Albornoz Claudio Gera"
            valorHoraAlCincuenta = Hoja4.Cells(fila, 12).Value / 120 * 1.5
            valorHoraAlCien = Hoja4.Cells(fila, 12).Value / 120 * 1.5
            
            ' Ajuste adicional solo para Albornoz Claudio Gera
            If apellido = "Albornoz Claudio Gera" Then
                valorHoraAlCien = Hoja4.Cells(fila, 12).Value / 120 * 2
            End If
    End Select
    
    ' Calcular importes
    importeHoraAlCincuenta = horasAlCincuenta * valorHoraAlCincuenta
    importeHoraAlCien = horasAlCien * valorHoraAlCien
    importeHoraFeriado = horasFeriado * valorHoraAlCien
    
    ' Calcular totales
    total = ActiveSheet.Cells(fila, 19).Value + importeHoraAlCien + importeHoraFeriado + importeHoraAlCincuenta
    valorHoraNormal = total - importeHoraFeriado
    
    ' Escribir resultados
    With ActiveSheet
        .Cells(fila, 25).Value = importeHoraFeriado
        .Cells(fila, 26).Value = importeHoraNormal
        .Cells(fila, 27).Value = importeHoraAlCincuenta
        .Cells(fila, 28).Value = importeHoraAlCien
        .Cells(fila, 29).Value = total
        .Cells(fila, 30).Value = total
    End With
End Sub
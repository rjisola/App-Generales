Sub calcularImporteGris(fila, categoria)

    Dim horasAlCincuenta  As Single
    Dim horasFeriado As Single
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim horasAlCien As Single
    Dim sueldoAcordado As Double
    Dim importeHorasAlCincuenta As Double
    Dim importeHorasAlCien As Double
    Dim importeHorasFeriado As Double
    Dim total As Double

    valorHoraAlCincuenta = Hoja2.Range("C1").Value
    valorHoraAlCien = Hoja2.Range("D1").Value

    horasAlCincuenta = Hoja2.Cells(fila, 21).Value
    importeHorasAlCincuenta = horasAlCincuenta * valorHoraAlCincuenta
    horasAlCien = Hoja2.Cells(fila, 22).Value
    importeHorasAlCien = horasAlCien * valorHoraAlCien
    horasFeriado = Hoja2.Cells(fila, 23).Value
    importeFeriado = horasFeriado * valorHoraAlCien
    
    sueldoAcordado = Hoja2.Cells(fila, 19).Value
        
    total = sueldoAcordado + importeHorasAlCincuenta + importeHorasAlCien + importeFeriado
    
    ActiveSheet.Cells(fila, 25).Value = importeFeriado
    ActiveSheet.Cells(fila, 27).Value = importeHorasAlCincuenta
    ActiveSheet.Cells(fila, 28).Value = importeHorasAlCien
    ActiveSheet.Cells(fila, 29).Value = total
    ActiveSheet.Cells(fila, 30).Value = total
    
    

End Sub

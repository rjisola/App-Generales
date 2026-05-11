

Sub calcularImporteCeleste(fila)


    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim importeHoraFeriado As Double
    Dim total As Double
    
    horasAlCincuenta = Hoja2.Cells(fila, 21).Value
    horasAlCien = Hoja2.Cells(fila, 22).Value
    horasFeriado = Hoja2.Cells(fila, 23).Value
    
    
   
    valorHoraAlCincuenta = Hoja4.Cells(fila, 12).Value / 120 * 1.5
    valorHoraFeriado = Hoja4.Cells(fila, 12).Value / 120 * 2
    valorHoraAlCien = Hoja4.Cells(fila, 12).Value / 120 * 2
        
    importeHoraAlCincuenta = horasAlCincuenta * valorHoraAlCincuenta
    importeHoraAlCien = horasAlCien * valorHoraAlCien
    importeHoraFeriado = horasFeriado * valorHoraAlCien
    
    total = ActiveSheet.Cells(fila, 19).Value + importeHoraAlCien + importeHoraFeriado + importeHoraAlCincuenta

    valorHoraNormal = total - importeHoraFeriado
    
    
    importeHoraNormal = total - importeHoraFeriado
    
    

    ActiveSheet.Cells(fila, 25).Value = importeHoraFeriado
    ActiveSheet.Cells(fila, 26).Value = importeHoraNormal
    ActiveSheet.Cells(fila, 27).Value = importeHoraAlCincuenta
    ActiveSheet.Cells(fila, 28).Value = importeHoraAlCien
    
    

    ActiveSheet.Cells(fila, 29).Value = total
    ActiveSheet.Cells(fila, 30).Value = total

End Sub
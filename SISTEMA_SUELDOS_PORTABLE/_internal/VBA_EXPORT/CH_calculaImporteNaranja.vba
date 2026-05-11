Sub calcularImporteNaranja(fila, maximoDias, categoria)

    Dim importeHoras As Double
    Dim valorHoraConPresentismo As Double
    Dim valorHoraSinPresentismo As Double
    Dim valorHoraConPresentismoF As Double
    Dim valorHoraSinPresentismoF As Double
    Dim valorHoraFeriado As Double
    Dim importeHoraFeriado As Double
    Dim horasFeriado As Double
    Dim valorHoraAlCien As Double
    Dim importeHoraAlCien As Double
    Dim total As Double
    Dim horas As Double
    Dim plusNasa As Double
    
    
    importeHoras = 0
    importeHoraFeriado = 0
    importeHoraAlCien = 0
    
    horas = ActiveSheet.Cells(fila, 20).Value + ActiveSheet.Cells(fila, 21).Value
    horasAlCien = ActiveSheet.Cells(fila, 22).Value
    valorHoraSinPresentismo = ActiveSheet.Range("E1").Value
    valorHoraConPresentismo = ActiveSheet.Range("F1").Value
    valorHoraSinPresentismoF = Hoja2.Cells(1, 11).Value
    valorHoraConPresentismoF = Hoja2.Cells(1, 12).Value
    
    valorHoraFeriado = Hoja2.Cells(1, 2).Value
    valorHoraAlCien = Hoja2.Cells(1, 4).Value

    If ActiveSheet.Cells(fila, 24).Value <> "PRESENTISMO" Then
        importeHoras = horas * valorHoraSinPresentismo
        If ActiveSheet.Cells(fila, 1).Value = "Ferreyra Diego Gaston" Then
            importeHoras = horas * valorHoraSinPresentismoF
        End If
    Else
        importeHoras = horas * valorHoraConPresentismo
        If ActiveSheet.Cells(fila, 1).Value = "Ferreyra Diego Gaston" Then
            importeHoras = horas * valorHoraConPresentismoF
        End If
    End If
        
   
        
    ActiveSheet.Cells(fila, 26).Value = importeHoras
        
    horasFeriado = Hoja2.Cells(fila, 23).Value
        
    importeHoraFeriado = valorHoraFeriado * horasFeriado
    
    importeHoraAlCien = valorHoraAlCien * horasAlCien

    total = importeHoras + importeHoraFeriado + importeHoraAlCien
    
    ActiveSheet.Cells(fila, 28).Value = importeHoraAlCien
    
    ActiveSheet.Cells(fila, 25).Value = importeHoraFeriado

    ActiveSheet.Cells(fila, 29).Value = total

    ActiveSheet.Cells(fila, 30).Value = total
    
    If ActiveSheet.Cells(fila, 35).Value = "SI" Then
        ActiveSheet.Cells(fila, 36).Value = horas * ActiveSheet.Range("N2").Value
    End If
    
    

 End Sub
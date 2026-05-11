Sub calcularImporteVerde(fila)


    Dim valorHoraNormal As Double
    Dim importeHoraNormal As Double
    Dim horasAlCien As Single
    Dim horasNormales As Single
    Dim presentismo As Boolean
    Dim total As Double
    
    horasNormales = Hoja2.Cells(fila, 20).Value
    horasAlCien = Hoja2.Cells(fila, 22).Value

    If Hoja2.Cells(fila, 24).Value = "PRESENTISMO" Then
        presentismo = True
    Else
        presentismo = False
    End If
    
    valorHoraNormal = Hoja2.Cells(1, 2).Value
    valorHoraAlCien = Hoja2.Cells(1, 2).Value
    
    importeHoraAlCien = horasAlCien * valorHoraAlCien
    
    If presentismo Then
        importeHoraNormal = horasNormales * valorHoraNormal * 1.2
    Else
        importeHoraNormal = horasNormales * valorHoraNormal
    End If
    
    ActiveSheet.Cells(fila, 26).Value = importeHoraNormal
    ActiveSheet.Cells(fila, 28).Value = importeHoraAlCien
    
    total = ActiveSheet.Cells(fila, 19).Value + importeHoraAlCien

    ActiveSheet.Cells(fila, 29).Value = total
    ActiveSheet.Cells(fila, 30).Value = total

End Sub

Sub calcularImporteSalmon(fila, ByRef presentismo, categoria)

    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim total As Double
    
    valorHoraNormal = 0
    valorHoraAlCincuenta = 0
    valorHoraAlCien = 0
    valorHoraFeriado = 0

    If categoria <> vbNullString Then
        
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(189, 215, 238)
    
        If categoria = "ESPECIALIZADO" Or categoria = "MAQUINISTA" Then
            valorHoraNormal = ActiveSheet.Range("B1").Value * 1.2
        Else
            If categoria = "OFICIAL" Then
                valorHoraNormal = ActiveSheet.Range("B2").Value * 1.2
            Else
                If categoria = "MEDIO OFICIAL" Then
                    valorHoraNormal = ActiveSheet.Range("B3").Value * 1.2
                Else
                    If categoria = "AYUDANTE" Then
                        valorHoraNormal = ActiveSheet.Range("B4").Value * 1.2
                    End If
                End If
            End If
        End If
        
    Else
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(255, 0, 0)
    End If

    valorHoraAlCincuenta = valorHoraNormal * 1.5
    valorHoraAlCien = valorHoraNormal * 2
    valorHoraFeriado = valorHoraAlCien
    
    If presentismo Then
        
        'importeHoraNormal = ActiveSheet.Cells(fila, 20).Value * valorHoraNormal * 1.2
        importeHoraAlCincuenta = ActiveSheet.Cells(fila, 21).Value * valorHoraAlCincuenta
        importeHoraAlCien = ActiveSheet.Cells(fila, 22).Value * valorHoraAlCien
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
    
    Else
    
        'importeHoraNormal = ActiveSheet.Cells(fila, 20).Value * valorHoraNormal
        importeHoraAlCincuenta = ActiveSheet.Cells(fila, 21).Value * valorHoraAlCincuenta
        importeHoraAlCien = ActiveSheet.Cells(fila, 22).Value * valorHoraAlCien
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
    
    End If
    
    ActiveSheet.Cells(fila, 25).Value = importeHoraFeriado
    ActiveSheet.Cells(fila, 26).Value = importeHoraNormal
    ActiveSheet.Cells(fila, 27).Value = importeHoraAlCincuenta
    ActiveSheet.Cells(fila, 28).Value = importeHoraAlCien

    total = importeHoraAlCincuenta + importeHoraAlCien + importeHoraFeriado

    ActiveSheet.Cells(fila, 29).Value = total
    ActiveSheet.Cells(fila, 30).Value = total

End Sub

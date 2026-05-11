Sub calcularImporteBlanco(fila, ByRef presentismo, categoria)

    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim valorHoraAltura As Double
    Dim importeHoraAltura As Double
    Dim total As Double
    
    valorHoraNormal = 0
    valorHoraAlCincuenta = 0
    valorHoraAlCien = 0
    valorHoraFeriado = 0

    If categoria <> vbNullString Then
        
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(189, 215, 238)
    
        If categoria = "ESPECIALIZADO" Or categoria = "MAQUINISTA" Then
            valorHoraNormal = ActiveSheet.Range("B1").Value * 1.2
            If presentismo Then
                valorHoraAltura = ActiveSheet.Range("AE5").Value
            Else
                valorHoraAltura = ActiveSheet.Range("AD5").Value
            End If
        Else
            If categoria = "OFICIAL" Then
                valorHoraNormal = ActiveSheet.Range("B2").Value * 1.2
                If presentismo Then
                    valorHoraAltura = ActiveSheet.Range("AE4").Value
                Else
                    valorHoraAltura = ActiveSheet.Range("AD4").Value
                End If
            Else
                If categoria = "MEDIO OFICIAL" Then
                    valorHoraNormal = ActiveSheet.Range("B3").Value * 1.2
                    If presentismo Then
                        valorHoraAltura = ActiveSheet.Range("AE3").Value
                    Else
                        valorHoraAltura = ActiveSheet.Range("AD3").Value
                    End If
                Else
                    If categoria = "AYUDANTE" Then
                        valorHoraNormal = ActiveSheet.Range("B4").Value * 1.2
                        If presentismo Then
                            valorHoraAltura = ActiveSheet.Range("AE2").Value
                        Else
                            valorHoraAltura = ActiveSheet.Range("AD2").Value
                        End If
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
    
    If ActiveSheet.Cells(fila, 1).Interior.color = RGB(255, 255, 255) Then
        horasAltura = ActiveSheet.Cells(fila, 31).Value
    End If
    
    If presentismo Then
        
        importeHoraAlCincuenta = ActiveSheet.Cells(fila, 21).Value * valorHoraAlCincuenta
        importeHoraAlCien = ActiveSheet.Cells(fila, 22).Value * valorHoraAlCien
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
    
    Else
    
        importeHoraAlCincuenta = ActiveSheet.Cells(fila, 21).Value * valorHoraAlCincuenta
        importeHoraAlCien = ActiveSheet.Cells(fila, 22).Value * valorHoraAlCien
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
    
    End If
    
    importeHoraAltura = ActiveSheet.Cells(fila, 31).Value * valorHoraAltura
    
    ActiveSheet.Cells(fila, 25).Value = importeHoraFeriado
    ActiveSheet.Cells(fila, 27).Value = importeHoraAlCincuenta
    ActiveSheet.Cells(fila, 28).Value = importeHoraAlCien
    ActiveSheet.Cells(fila, 32).Value = importeHoraAltura

    total = importeHoraAlCincuenta + importeHoraAlCien + importeHoraFeriado + importeHoraAltura

    ActiveSheet.Cells(fila, 29).Value = total
    ActiveSheet.Cells(fila, 30).Value = total

End Sub

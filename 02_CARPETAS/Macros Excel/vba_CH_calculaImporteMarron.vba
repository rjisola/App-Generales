Sub calcularImporteMarron(fila, columna, categoria, Dia, presentismo, feriado)
    
    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim total As Double
    Dim valorHoraAltura As Double
    Dim importeHoraAltura As Double
    
    valorHoraNormal = 0
    valorHoraAlCincuenta = 0
    valorHoraAlCien = 0
    valorHoraFeriado = 0
    
    If categoria <> vbNullString Then
        
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(189, 215, 238)
        
        If categoria = "ANDAMISTA ESP" Then
            valorHoraNormal = ActiveSheet.Range("AP1").Value
            valorHoraAltura = ActiveSheet.Range("AI1").Value
            If presentismo Then
                valorHoraNormal = ActiveSheet.Range("AS1").Value
            Else
                valorHoraNormal = ActiveSheet.Range("AP1").Value
            End If
        Else
            If categoria = "ESPECIALIZADO" Or categoria = "MAQUINISTA" Then
                valorHoraNormal = ActiveSheet.Range("AJ1").Value
                valorHoraAltura = ActiveSheet.Range("AH1").Value
                If presentismo Then
                    valorHoraNormal = ActiveSheet.Range("AM1").Value
                Else
                    valorHoraNormal = ActiveSheet.Range("AJ1").Value
                End If
            Else
                If categoria = "ANDAMISTA OFIC" Then
                    valorHoraNormal = ActiveSheet.Range("AP2").Value
                    valorHoraAltura = ActiveSheet.Range("AI2").Value
                    If presentismo Then
                        valorHoraNormal = ActiveSheet.Range("AS2").Value
                    Else
                        valorHoraNormal = ActiveSheet.Range("AP2").Value
                    End If
                Else
                    If categoria = "OFICIAL" Then
                        valorHoraNormal = ActiveSheet.Range("AJ2").Value
                        valorHoraAltura = ActiveSheet.Range("AH2").Value
                        If presentismo Then
                            valorHoraNormal = ActiveSheet.Range("AM2").Value
                        Else
                            valorHoraNormal = ActiveSheet.Range("AJ2").Value
                        End If
                    Else
                        If categoria = "MEDIO OFICIAL" Then
                            valorHoraNormal = ActiveSheet.Range("AJ3").Value
                            valorHoraAltura = ActiveSheet.Range("AH3").Value
                            If presentismo Then
                                valorHoraNormal = ActiveSheet.Range("AM3").Value
                            Else
                                valorHoraNormal = ActiveSheet.Range("AJ3").Value
                            End If
                        Else
                            If categoria = "AYUDANTE" Then
                                valorHoraNormal = ActiveSheet.Range("AJ4").Value
                                valorHoraAltura = ActiveSheet.Range("AH4").Value
                                If presentismo Then
                                    valorHoraNormal = ActiveSheet.Range("AM4").Value
                                Else
                                    valorHoraNormal = ActiveSheet.Range("AJ4").Value
                                End If
                                
                            End If
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
    
    If presentismo Then
        
        importeHoraAlCincuenta = ActiveSheet.Cells(fila, 21).Value * valorHoraAlCincuenta
        importeHoraAlCien = ActiveSheet.Cells(fila, 22).Value * valorHoraAlCien
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
        
    Else
        
        importeHoraAlCincuenta = ActiveSheet.Cells(fila, 21).Value * valorHoraAlCincuenta
        importeHoraAlCien = ActiveSheet.Cells(fila, 22).Value * valorHoraAlCien
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
        
    End If
    
    ActiveSheet.Cells(fila, 25).Value = importeHoraFeriado
    ActiveSheet.Cells(fila, 27).Value = importeHoraAlCincuenta
    ActiveSheet.Cells(fila, 28).Value = importeHoraAlCien
    importeHoraAltura = ActiveSheet.Cells(fila, 31).Value * valorHoraAltura
    ActiveSheet.Cells(fila, 32).Value = importeHoraAltura
    total = importeHoraAlCincuenta + importeHoraAlCien + importeHoraFeriado
    
    ActiveSheet.Cells(fila, 29).Value = total
    ActiveSheet.Cells(fila, 30).Value = total
    
End Sub
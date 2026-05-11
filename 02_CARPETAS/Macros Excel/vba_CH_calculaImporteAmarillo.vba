Sub calcularImporteAmarillo(fila, ByRef presentismo, categoria, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)

    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim importeHorasQuilmesCien As Double
    Dim importeHorasPapeleraCien As Double
    Dim importeHorasQuilmesCincuenta As Double
    Dim importeHorasPapeleraCincuenta As Double
    Dim importeHorasAlCincuentaBlancas As Double
    Dim importeHorasAlCienBlancas As Double
    Dim total As Double
    
    valorHoraNormal = 0
    valorHoraAlCincuenta = 0
    valorHoraAlCien = 0
    valorHoraFeriado = 0

    If categoria <> vbNullString Then
        
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(189, 215, 238)
    
        If categoria = "ESPECIALIZADO" Or categoria = "MAQUINISTA" Then
            valorHoraNormal = ActiveSheet.Range("B1").Value
        Else
            If categoria = "OFICIAL" Then
                valorHoraNormal = ActiveSheet.Range("B2").Value
            Else
                If categoria = "MEDIO OFICIAL" Then
                    valorHoraNormal = ActiveSheet.Range("B3").Value
                Else
                    If categoria = "AYUDANTE" Then
                        valorHoraNormal = ActiveSheet.Range("B4").Value
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
        
        'importeHoraNormal = ActiveSheet.Cells(fila, 20).Value * valorHoraNormal
        horasBlancasCincuenta = ActiveSheet.Cells(fila, 21).Value - horasPapeleraCincuenta - horasQuilmesCincuenta
        
        importeHorasAlCincuentaBlancas = ((horasBlancasCincuenta) * valorHoraAlCincuenta)
        importeHorasQuilmesCincuenta = ((horasQuilmesCincuenta) * valorHoraAlCincuenta * 1.2)
        importeHorasPapeleraCincuenta = ((horasPapeleraCincuenta) * valorHoraAlCincuenta * 1.2 * 1.12)
        importeHoraAlCincuenta = importeHorasQuilmesCincuenta + importeHorasPapeleraCincuenta + importeHorasAlCincuentaBlancas
        
        horasBlancasCien = ActiveSheet.Cells(fila, 22).Value - horasPapeleraCien - horasQuilmesCien
        
        importeHorasAlCienBlancas = ((horasBlancasCien) * valorHoraAlCien)
        importeHorasQuilmesCien = ((horasQuilmesCien) * valorHoraAlCien * 1.2)
        importeHorasPapeleraCien = ((horasPapeleraCien) * valorHoraAlCien * 1.2 * 1.12)
        importeHoraAlCien = importeHorasQuilmesCien + importeHorasPapeleraCien + importeHorasAlCienBlancas
       
        importeHoraFeriado = ActiveSheet.Cells(fila, 23).Value * valorHoraFeriado
    Else
    
        'importeHoraNormal = ActiveSheet.Cells(fila, 20).Value * valorHoraNormal
        horasBlancasCincuenta = ActiveSheet.Cells(fila, 21).Value - horasPapeleraCincuenta - horasQuilmesCincuenta
        
        importeHorasAlCincuentaBlancas = ((horasBlancasCincuenta) * valorHoraAlCincuenta)
        importeHorasQuilmesCincuenta = ((horasQuilmesCincuenta) * valorHoraAlCincuenta * 1.2)
        importeHorasPapeleraCincuenta = ((horasPapeleraCincuenta) * valorHoraAlCincuenta * 1.2 * 1.12)
        importeHoraAlCincuenta = importeHorasQuilmesCincuenta + importeHorasPapeleraCincuenta + importeHorasAlCincuentaBlancas
        
        horasBlancasCien = ActiveSheet.Cells(fila, 22).Value - horasPapeleraCien - horasQuilmesCien
        
        importeHorasAlCienBlancas = ((horasBlancasCien) * valorHoraAlCien)
        importeHorasQuilmesCien = ((horasQuilmesCien) * valorHoraAlCien * 1.2)
        importeHorasPapeleraCien = ((horasPapeleraCien) * valorHoraAlCien * 1.2 * 1.12)
        importeHoraAlCien = importeHorasQuilmesCien + importeHorasPapeleraCien + importeHorasAlCienBlancas
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

Sub completarImprimirBlanco(fila, contador, columna, desplazamiento, color)
    
    Dim nombre As String
    Dim quincena As String
    Dim categoria As String
    Dim horasAlCincuenta As Double
    Dim importeHorasAlCincuenta As Double
    Dim horasAlCien As Single
    Dim importeHorasAlCien As Double
    Dim totalExtras As Double
    Dim horasAltura As Single
    Dim importeHorasAltura As Double
    Dim premio As Double
    Dim presentismo As String
    Dim sueldoSobre As Double
    Dim totalQuincena As Double
    Dim adelanto As Double
    Dim obraSocial As Double
    Dim banco As Double
    Dim cajaDeAhorro As Double
    Dim gastoPersonal As Double
    Dim reintegro As Double
    Dim fondoDesempleo As Double
    Dim horasFeriado As Single
    Dim importeHorasFeriado As Double

    Call colorearImprimir(contador, columna, color, desplazamiento)
    
    '*****************
    'ASIGNO VARIABLES*
    '*****************
    
    'Nombre
    nombre = Hoja2.Cells(fila, 1).Value
    'Quincena
    quincena = Hoja2.Cells(6, 20).Value
    'Categoria
    categoria = Hoja2.Cells(fila, 2).Value
    'horas al cincuenta
    horasAlCincuenta = Hoja2.Cells(fila, 21).Value
    'Importe horas al cincuenta
    importeHorasAlCincuenta = Hoja2.Cells(fila, 27).Value
    'horas feriado
    horasFeriado = Hoja2.Cells(fila, 23).Value
    'Importe horas feriado
    importeHorasFeriado = Hoja2.Cells(fila, 25).Value
    'Reintegro
    reintegro = Hoja4.Cells(fila, 14).Value
    'Horas al cien
    horasAlCien = Hoja2.Cells(fila, 22).Value + horasFeriado
    'Importe horas al cien
    importeHorasAlCien = Hoja2.Cells(fila, 28).Value + importeHorasFeriado
    'Fondo de desempleo
    fondoDesempleo = (importeHorasAlCincuenta + importeHorasAlCien) * 0.12
    'Horas en altura
    horasAltura = Hoja2.Cells(fila, 31).Value
    importeHorasAltura = Hoja2.Cells(fila, 32).Value
    'Total extras
    totalExtras = importeHorasAlCincuenta + importeHorasAlCien + fondoDesempleo
    'Ajuste premio equivalente a VARIOS del verde
    If IsNumeric(Hoja4.Cells(fila, 19)) Then
        premio = Hoja4.Cells(fila, 19).Value
    End If
    'Sueldo sobre
    sueldoSobre = Hoja4.Cells(fila, 10).Value
    'Presentismo
    If Hoja2.Cells(fila, 24).Value = "PRESENTISMO" Then
        presentismo = "SI"
    Else
        presentismo = "NO"
    End If
    'Total quincena
    totalQuincena = totalExtras + premio + reintegro + importeHorasAltura + sueldoSobre
    totalQuincena = Redondear(totalQuincena)
    'Adelanto
    adelanto = Hoja4.Cells(fila, 13).Value
    'Gasto personal
    gastoPersonal = Hoja4.Cells(fila, 16).Value
    'Obra social
    obraSocial = Hoja4.Cells(fila, 17).Value
    'Banco
    banco = sueldoSobre
    'Caja de ahorro
    cajaDeAhorro = totalQuincena - adelanto - obraSocial - banco - gastoPersonal
    
    If cajaDeAhorro < 0 Then
        banco = banco + cajaDeAhorro
        cajaDeAhorro = 0
    End If
    
    '**************************************
    'Limpio el "cuadradito" que se imprime*
    '**************************************
    
    Call limpiarPersonaImprimir(contador, columna)
    
    '********************
    'ESCRIBO EN IMPRIMIR*
    '********************
    
    'Nombre
    Call mergearCeldas(contador, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador, columna).Value = "Apellido y Nombre"
    Hoja3.Cells(contador, columna + 1).Value = nombre
    Hoja3.Cells(contador, columna + 1).Font.Size = 10
    
    'Quincena
    Call mergearCeldas(contador + 1, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 1, columna).Value = "QUINCENA"
    Hoja3.Cells(contador + 1, columna + 1).Value = quincena
    
    'Categoria
    Call mergearCeldas(contador + 2, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 2, columna).Value = "Categoría"
    Hoja3.Cells(contador + 2, columna + 1).Value = categoria
    
    
    'Horas Al cincuenta
    Call unMergearCeldas(contador + 3, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 3, columna).Value = "HS.50%"
    Hoja3.Cells(contador + 3, columna + 1).Value = horasAlCincuenta
    Hoja3.Cells(contador + 3, columna + 2).Value = importeHorasAlCincuenta
    
    'Horas al cien
    Call unMergearCeldas(contador + 4, columna + 1, columna + 2, Hoja3)
    If horasFeriado <> 0 Then
        Hoja3.Cells(contador + 4, columna).Value = "HS.100% + FERIADO"
    Else
        Hoja3.Cells(contador + 4, columna).Value = "HS.100%"
    End If
    Hoja3.Cells(contador + 4, columna + 1).Value = horasAlCien
    Hoja3.Cells(contador + 4, columna + 2).Value = importeHorasAlCien
    
    'Fondo desempleo
    Call mergearCeldas(contador + 5, columna + 1, columna + 2, Hoja3)
    If fondoDesempleo <> 0 Then
        Hoja3.Cells(contador + 5, columna).Value = "Fondo des. 12%"
        Hoja3.Cells(contador + 5, columna + 1).Value = fondoDesempleo
    End If
    
    'Total extras
    Call mergearCeldas(contador + 6, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 6, columna).Value = "TOTAL EXTRAS"
    Hoja3.Cells(contador + 6, columna + 1).Value = totalExtras
    
    'Presentismo
    Call mergearCeldas(contador + 7, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 7, columna).Value = "PRESENTISMO"
    Hoja3.Cells(contador + 7, columna + 1).Value = presentismo
    
    'Horas en altura
    Call mergearCeldas(contador + 8, columna + 1, columna + 2, Hoja3)
    If horasAltura <> 0 Then
        Call unMergearCeldas(contador + 8, columna + 1, columna + 2, Hoja3)
        Hoja3.Cells(contador + 8, columna).Value = "Altura/Hormigón 15%"
        Hoja3.Cells(contador + 8, columna + 1).Value = horasAltura
        Hoja3.Cells(contador + 8, columna + 2).Value = importeHorasAltura
    End If
    
    'Reintegro
    Call mergearCeldas(contador + 9, columna + 1, columna + 2, Hoja3)
    If reintegro <> 0 Then
        Hoja3.Cells(contador + 9, columna).Value = "REINTEGRO"
        Hoja3.Cells(contador + 9, columna + 1).Value = reintegro
    End If

    
    'Sueldo Sobre
    Call mergearCeldas(contador + 10, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 10, columna).Value = "SUELDO SOBRE"
    Hoja3.Cells(contador + 10, columna + 1).Value = sueldoSobre
    
    'Total quincena
    Hoja3.Cells(contador + 11, columna).Value = "TOTAL QUINCENA"
    Hoja3.Cells(contador + 11, columna + 1).Value = totalQuincena
    Hoja3.Cells(contador + 11, columna + 1).NumberFormat = " $#,##0.00"
    Hoja3.Cells(contador + 11, columna + 1).HorizontalAlignment = xlCenter
    Hoja3.Cells(contador + 11, columna + 1).VerticalAlignment = xlCenter
    
    'Premio
    If premio <> 0 And reintegro = 0 And horasAltura = 0 Then
        Call mergearCeldas(contador + 9, columna + 1, columna + 2, Hoja3)
        Hoja3.Cells(contador + 9, columna).Value = "PREMIO"
        Hoja3.Cells(contador + 9, columna + 1).Value = premio
    End If
    
    'Adelanto
    Call mergearCeldas(contador + 14, columna + 1, columna + 2, Hoja3)
    Hoja3.Cells(contador + 14, columna).Value = "ADELANTO"
    Hoja3.Cells(contador + 14, columna + 1).Value = adelanto
    
    'Gastos
    Call mergearCeldas(contador + 15, columna + 1, columna + 2, Hoja3)
    If gastoPersonal <> 0 Then
        Hoja3.Cells(contador + 15, columna).Value = "GASTOS"
        Hoja3.Cells(contador + 15, columna + 1).Value = gastoPersonal
    End If
    
    'Obra social
    Call mergearCeldas(contador + 16, columna + 1, columna + 2, Hoja3)
    If obraSocial <> 0 And obraSocial > 0 Then
        Hoja3.Cells(contador + 16, columna).Value = "OBRA SOCIAL"
        Hoja3.Cells(contador + 16, columna + 1).Value = obraSocial
    End If
    
    If Hoja4.Cells(fila, 3).Value <> 0 And Hoja4.Cells(fila, 4).Value <> 0 Then
        'Banco
        Call mergearCeldas(contador + 17, columna + 1, columna + 2, Hoja3)
        Hoja3.Cells(contador + 17, columna).Value = "BANCO"
        Hoja3.Cells(contador + 17, columna + 1).Value = banco
     
        'Caja de Ahorro
        Call mergearCeldas(contador + 18, columna + 1, columna + 2, Hoja3)
        Hoja3.Cells(contador + 18, columna).Value = "Caja de Ahorro N°2"
        Hoja3.Cells(contador + 18, columna + 1).Value = cajaDeAhorro
        
    Else
        If Hoja4.Cells(fila, 3).Value <> 0 And Hoja4.Cells(fila, 4).Value = 0 Then
            'Banco
            Call mergearCeldas(contador + 17, columna + 1, columna + 2, Hoja3)
            Hoja3.Cells(contador + 17, columna).Value = "BANCO"
            Hoja3.Cells(contador + 17, columna + 1).Value = banco
     
            'Efectivo
            Call mergearCeldas(contador + 18, columna + 1, columna + 2, Hoja3)
            Hoja3.Cells(contador + 18, columna).Value = "EFECTIVO"
            Hoja3.Cells(contador + 18, columna + 1).Value = cajaDeAhorro
        Else
            'Efectivo
            Call mergearCeldas(contador + 18, columna + 1, columna + 2, Hoja3)
            Hoja3.Cells(contador + 18, columna).Value = "EFECTIVO"
            Hoja3.Cells(contador + 18, columna + 1).Value = banco + cajaDeAhorro
        End If
    End If

    
    'Pegar en recuento total
    Hoja10.Cells(fila - 7, 1).Interior.color = color
    Hoja10.Cells(fila - 7, 5).Value = totalQuincena
    Hoja10.Cells(fila - 7, 1).Value = nombre
    Hoja10.Cells(fila - 7, 1).Value = Hoja4.Cells(fila, 3).Value
    Hoja10.Cells(fila - 7, 2).Value = Hoja4.Cells(fila, 4).Value
    If IsEmpty(Hoja4.Cells(fila, 3)) And IsEmpty(Hoja4.Cells(fila, 4)) Then
        Hoja10.Cells(fila - 7, 4).Value = banco + cajaDeAhorro
    Else
        If IsEmpty(Hoja4.Cells(fila, 3)) And Not IsEmpty(Hoja4.Cells(fila, 4)) Then
            Hoja10.Cells(fila - 7, 4).Value = banco
        Else
            Hoja10.Cells(fila - 7, 2).Value = banco
        End If
        If IsEmpty(Hoja4.Cells(fila, 4)) And Not IsEmpty(Hoja4.Cells(fila, 3)) Then
            Hoja10.Cells(fila - 7, 4).Value = cajaDeAhorro
        Else
            Hoja10.Cells(fila - 7, 3).Value = cajaDeAhorro
        End If
    End If

End Sub

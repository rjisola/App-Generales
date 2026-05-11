Sub completarImprimirNaranja(fila, contador, columna, desplazamiento, color)
    
    Dim nombre As String
    Dim quincena As String
    Dim categoria As String
    Dim horasNormales As Double
    Dim importeHorasNormales As Double
    Dim horasAlCien As Single
    Dim importeHorasAlCien As Double
    Dim totalExtras As Double
    Dim premio As Double
    Dim presentismo As String
    Dim sueldoSobre As Double
    Dim totalQuincena As Double
    Dim adelanto As Double
    Dim descPatente As Double
    Dim obraSocial As Double
    Dim banco As Double
    Dim cajaDeAhorro As Double
    Dim gastoPersonal As Double
    Dim reintegro As Double
    Dim totalImporte As Double
    Dim importeARestarParaPremio As Double
    Dim ajusteAlquiler As Double
    Dim horasFeriado As Single
    Dim importeHorasFeriado As Double
    Dim plusNasa As Double
    Dim Legajo As Double
    
    
    Call colorearImprimir(contador, columna, color, desplazamiento)
    
    Hoja1.Cells(contador + 18, columna).VerticalAlignment = xlCenter
    Hoja1.Cells(contador + 18, columna).HorizontalAlignment = xlLeft
    
    '*****************
    'ASIGNO VARIABLES*
    '*****************
    
    'Nombre
    nombre = Hoja2.Cells(fila, 1).Value
    'Quincena
    quincena = Hoja2.Cells(6, 20).Value
    'Categoria
    categoria = Hoja2.Cells(fila, 2).Value
    'Horas normales
    horasNormales = Hoja2.Cells(fila, 20).Value
    'Importe horas normales
    importeHorasNormales = Hoja2.Cells(fila, 26).Value
    'Numero de Legajo
    Legajo = Hoja4.Cells(fila, 2).Value
    'Reintegro
    reintegro = Hoja4.Cells(fila, 14).Value
    'Horas al cien
    horasAlCien = Hoja2.Cells(fila, 22).Value
    'Importe horas al cien
    importeHorasAlCien = Hoja2.Cells(fila, 28).Value
    'Total extras
    totalExtras = importeHorasAlCien
    'Horas feriado
    horasFeriado = Hoja2.Cells(fila, 23).Value
    'Ajuste Alquiler
    ajusteAlquiler = Hoja4.Cells(fila, 15).Value
    'Importe horas feriado
    importeHorasFeriado = Hoja2.Cells(fila, 25).Value
    
    'Sueldo sobre
    sueldoSobre = Hoja4.Cells(fila, 10).Value
    'Presentismo
    If Hoja2.Cells(fila, 24).Value = "PRESENTISMO" Then
        presentismo = "SI"
    Else
        presentismo = "NO"
    End If
    'Ajuste premio
    totalImporte = Hoja2.Cells(fila, 28).Value + Hoja2.Cells(fila, 26).Value
    If Hoja2.Cells(fila, 24).Value = "PRESENTISMO" Then
        importeARestarParaPremio = (horasNormales + horasAlCien) * Hoja2.Cells(1, 6).Value
    Else
        importeARestarParaPremio = (horasNormales + horasAlCien) * Hoja2.Cells(1, 5).Value
    End If
    If IsNumeric(Hoja4.Cells(fila, 19)) Then
        premio = importeARestarParaPremio - totalImporte
    End If
    'Ajuste alquiler
    ajusteAlquiler = Hoja4.Cells(fila, 15).Value
    'Total quincena
    totalQuincena = importeHorasNormales + reintegro + ajusteAlquiler + importeHorasFeriado + importeHorasAlCien + Hoja2.Cells(fila, 36).Value + adelanto
    totalQuincena = Redondear(totalQuincena)
    'Adelanto
    adelanto = Hoja4.Cells(fila, 13).Value
    'Gasto personal
    gastoPersonal = Hoja4.Cells(fila, 16).Value
    'Decuento patente
    descPatente = Hoja4.Cells(fila, 18).Value
    'Obra social
    obraSocial = Hoja4.Cells(fila, 17).Value
    'Banco
    banco = sueldoSobre
    'Caja de ahorro
    cajaDeAhorro = totalQuincena - adelanto - descPatente - obraSocial - banco - gastoPersonal
    
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
    Call mergearCeldas(contador, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador, columna).Value = "Leg N° " & Legajo
    Hoja1.Cells(contador, columna + 1).Value = nombre
    Hoja1.Cells(contador, columna + 1).Font.Size = 10
    
    'Quincena
    Call mergearCeldas(contador + 1, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 1, columna).Value = "QUINCENA"
    Hoja1.Cells(contador + 1, columna + 1).Value = quincena
    
    'Categoria
    Call mergearCeldas(contador + 2, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 2, columna).Value = "Categoría"
    Hoja1.Cells(contador + 2, columna + 1).Value = categoria
    
    'Aclaracion de HORAS y $
    Call unMergearCeldas(contador + 3, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 3, columna + 1).Value = "HORAS"
    Hoja1.Cells(contador + 3, columna + 2).Value = "($)"
    
    'Horas normales
    Call unMergearCeldas(contador + 4, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 4, columna).Value = "HS. TOTALES"
    Hoja1.Cells(contador + 4, columna + 1).Value = horasNormales
    Hoja1.Cells(contador + 4, columna + 2).Value = importeHorasNormales
     'Horas al Cien
    If horasAlCien <> 0 Then
    Call unMergearCeldas(contador + 5, columna + 1, columna + 2, Hoja1)
        Call unMergearCeldas(contador + 5, columna + 1, columna + 2, Hoja1)
        Hoja1.Cells(contador + 5, columna).Value = "HS AL 100%"
        Hoja1.Cells(contador + 5, columna + 1).Value = horasAlCien
        Hoja1.Cells(contador + 5, columna + 2).Value = importeHorasAlCien
    End If
    'Horas feriado
    If horasFeriado <> 0 Then
        Call unMergearCeldas(contador + 6, columna + 1, columna + 2, Hoja1)
        Hoja1.Cells(contador + 6, columna).Value = "HS FERIADO"
        Hoja1.Cells(contador + 6, columna + 1).Value = horasFeriado
        Hoja1.Cells(contador + 6, columna + 2).Value = importeHorasFeriado
    End If
    'Ajuste Alquiler
    Call mergearCeldas(contador + 9, columna + 1, columna + 2, Hoja1)
    If ajusteAlquiler <> 0 Then
        Hoja1.Cells(contador + 9, columna).Value = "AJUSTE-ALQUILER"
        Hoja1.Cells(contador + 9, columna + 1).Value = ajusteAlquiler
    End If
    'Reintegro o alquiler
    Call mergearCeldas(contador + 7, columna + 1, columna + 2, Hoja1)
    If reintegro <> 0 Then
        Hoja1.Cells(contador + 7, columna).Value = "REINTEGRO"
        Hoja1.Cells(contador + 7, columna + 1).Value = reintegro
    Else
        If Hoja2.Cells(fila, 35).Value = "SI" Then
            Hoja1.Cells(contador + 7, columna).Value = "PLUS NASA"
            Hoja1.Cells(contador + 7, columna + 1).Value = Hoja2.Cells(fila, 36).Value
        Else
        If ajusteAlquiler <> 0 Then
            Hoja1.Cells(contador + 7, columna).Value = "ALQUILER"
            Hoja1.Cells(contador + 7, columna + 1).Value = ajusteAlquiler
        End If
        End If
    End If

    
    'Presentismo
    Call mergearCeldas(contador + 9, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 9, columna).Value = "PRESENTISMO"
    Hoja1.Cells(contador + 9, columna + 1).Value = presentismo
    
    'Sueldo Sobre
    Call mergearCeldas(contador + 10, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 10, columna).Value = "SUELDO SOBRE"
    Hoja1.Cells(contador + 10, columna + 1).Value = sueldoSobre
    
    'Total quincena
    Hoja1.Cells(contador + 11, columna).Value = "TOTAL QUINCENA"
    Hoja1.Cells(contador + 11, columna + 1).Value = totalQuincena
    Hoja1.Cells(contador + 11, columna + 1).NumberFormat = " $#,##0.00"
    Hoja1.Cells(contador + 11, columna + 1).HorizontalAlignment = xlCenter
    Hoja1.Cells(contador + 11, columna + 1).VerticalAlignment = xlCenter
    
    'Adelanto
    Call mergearCeldas(contador + 14, columna + 1, columna + 2, Hoja1)
    Hoja1.Cells(contador + 14, columna).Value = "ADELANTO"
    Hoja1.Cells(contador + 14, columna + 1).Value = adelanto
    
    'Descuento Patente
    Call mergearCeldas(contador + 15, columna + 1, columna + 2, Hoja1)
    If descPatente <> 0 Or gastoPersonal <> 0 Then
        Hoja1.Cells(contador + 15, columna).Value = "PATENTE - GASTOS"
        Hoja1.Cells(contador + 15, columna + 1).Value = descPatente + gastoPersonal
    End If
    
    'Obra social
    Call mergearCeldas(contador + 16, columna + 1, columna + 2, Hoja1)
    If obraSocial <> 0 And obraSocial > 0 Then
        Hoja1.Cells(contador + 16, columna).Value = "OBRA SOCIAL"
        Hoja1.Cells(contador + 16, columna + 1).Value = obraSocial
    End If
    
    
    
    If Hoja4.Cells(fila, 3).Value <> 0 And Hoja4.Cells(fila, 4).Value <> 0 Then
        'Banco
        Call mergearCeldas(contador + 17, columna + 1, columna + 2, Hoja1)
        Hoja1.Cells(contador + 17, columna).Value = "BANCO"
        Hoja1.Cells(contador + 17, columna + 1).Value = banco
     
        'Caja de Ahorro
        Call mergearCeldas(contador + 18, columna + 1, columna + 2, Hoja1)
        Hoja1.Cells(contador + 18, columna).VerticalAlignment = xlCenter
        Hoja1.Cells(contador + 18, columna).Value = "Caja de Ahorro N°2"
        Hoja1.Cells(contador + 18, columna + 1).Value = cajaDeAhorro
        
    Else
        If Hoja4.Cells(fila, 3).Value <> 0 And Hoja4.Cells(fila, 4).Value = 0 Then
            'Banco
            Call mergearCeldas(contador + 17, columna + 1, columna + 2, Hoja1)
            Hoja1.Cells(contador + 17, columna).Value = "BANCO"
            Hoja1.Cells(contador + 17, columna + 1).Value = banco
     
            'Efectivo
            Call mergearCeldas(contador + 18, columna + 1, columna + 2, Hoja1)
            Hoja1.Cells(contador + 18, columna).Value = "EFECTIVO"
            Hoja1.Cells(contador + 18, columna + 1).Value = cajaDeAhorro
        Else
            'Efectivo
            Call mergearCeldas(contador + 18, columna + 1, columna + 2, Hoja1)
            Hoja1.Cells(contador + 18, columna).Value = "EFECTIVO"
            Hoja1.Cells(contador + 18, columna + 1).Value = banco + cajaDeAhorro
        End If
    End If

    'Pegar en recuento total
    Hoja10.Cells(fila - 7, 1 + 3).Interior.color = color
    Hoja10.Cells(fila - 7, 5 + 3).Value = totalQuincena
    Hoja10.Cells(fila - 7, 1 + 3).Value = nombre
    Hoja10.Cells(fila - 7, 2).Value = Hoja4.Cells(fila, 3).Value
    Hoja10.Cells(fila - 7, 6 + 5).Value = Hoja4.Cells(fila, 27).Value
    
    If (Hoja4.Cells(fila, 3).Interior.color = RGB(255, 0, 0)) Then
        Hoja10.Cells(fila - 7, 2).Interior.color = Hoja4.Cells(fila, 3).Interior.color
    End If
    Hoja10.Cells(fila - 7, 3).Value = Hoja4.Cells(fila, 4).Value
    Hoja10.Cells(fila - 7, 1).Value = Hoja9.Cells(fila, 2).Value
    Hoja10.Cells(fila - 7, 6 + 3).Value = Hoja9.Cells(fila, 25).Value
    Hoja10.Cells(fila - 7, 6 + 4).Value = Hoja4.Cells(fila, 5).Value
    
    If IsEmpty(Hoja4.Cells(fila, 3)) And IsEmpty(Hoja4.Cells(fila, 4)) Then
        Hoja10.Cells(fila - 7, 4 + 3).Value = banco + cajaDeAhorro
    Else
        If IsEmpty(Hoja4.Cells(fila, 3)) And Not IsEmpty(Hoja4.Cells(fila, 4)) Then
            Hoja10.Cells(fila - 7, 4 + 3).Value = banco
        Else
            Hoja10.Cells(fila - 7, 2 + 3).Value = banco
        End If
        If IsEmpty(Hoja4.Cells(fila, 4)) And Not IsEmpty(Hoja4.Cells(fila, 3)) Then
            Hoja10.Cells(fila - 7, 4 + 3).Value = cajaDeAhorro
        Else
            Hoja10.Cells(fila - 7, 3 + 3).Value = cajaDeAhorro
        End If
    End If

End Sub


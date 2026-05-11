'Blanco - horas normales 8, 4 sabado si trabaja domingo no se considera, no se le cuentan las horas, si se cuenta cuando son feriado.
'            Vacaciones: 0 horas, pero no pierde presentismo.
'            Falto: 0 horas, pierde presentismo.
'            Enfermo y certificado: horas normales, pierde presentismo.
'            C/A o cortaron: 0 horas pero no pierde presentismo.
'            Lluvia: 2,5 horas
'Naranja - Quilmes
'Verde - Papelera
'Azul - Toyota
'Marron - Nasa



Sub copiarEnContador()
    Application.EnableAnimations = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws2 As Worksheet: Set ws2 = Hoja2 ' Usar alias más cortos
    Dim ws9 As Worksheet: Set ws9 = Hoja9
    Dim fila As Long, columna As Long
    Dim colorHoras As Long, colorEmpleado As Long
    Dim horas As Variant, Dia As String
    Dim categoria As String, apellido As String
    Dim precioCategoria As Double
    Dim columnaBase As Integer
    Dim startFila As Long, endFila As Long
    Dim startCol As Long, endCol As Long
    
    ' Definir rangos una sola vez
    startFila = 8 + ws2.Range("U3").Value
    endFila = ws2.Range("U4").Value + 8
    startCol = 2 + ws2.Range("U1").Value
    endCol = ws2.Range("U2").Value + 2
    
    ' Limpiar celdas de destino en bloques para mayor eficiencia
    For fila = startFila To endFila
        ws9.Range(ws9.Cells(fila, 6), ws9.Cells(fila, 19)).Value = 0
        ws9.Cells(fila, 27).Value = 0
        
        For columna = startCol To endCol
            colorHoras = ws2.Cells(fila, columna).Interior.color
            horas = ws2.Cells(fila, columna).Value
            Dia = LCase(ws2.Cells(8, columna).Value) ' Convertir a minúsculas para simplificar comparaciones
            categoria = ws2.Cells(fila, 2).Value
            colorEmpleado = ws2.Cells(fila, 1).Interior.color
            apellido = ws2.Cells(fila, 1).Value
            
            ' Determinar columna base basada en color
            Select Case colorHoras
                Case RGB(255, 255, 255): columnaBase = 6
                Case RGB(255, 192, 0): columnaBase = 7
                Case RGB(112, 173, 71), RGB(153, 102, 0): columnaBase = 8
            End Select
            
            ' Determinar precio de categoría con estructura más clara
            Select Case categoria
                Case "ESPECIALIZADO"
                    precioCategoria = IIf(colorEmpleado = RGB(153, 102, 0), ws2.Cells(1, 36).Value, ws2.Cells(1, 2).Value)
                Case "OFICIAL"
                    precioCategoria = IIf(colorEmpleado = RGB(153, 102, 0), ws2.Cells(2, 36).Value, ws2.Cells(2, 2).Value)
                Case "MEDIO OFICIAL"
                    precioCategoria = IIf(colorEmpleado = RGB(153, 102, 0), ws2.Cells(3, 36).Value, ws2.Cells(3, 2).Value)
                Case "AYUDANTE"
                    precioCategoria = IIf(colorEmpleado = RGB(153, 102, 0), ws2.Cells(4, 36).Value, ws2.Cells(4, 2).Value)
            End Select
            
            ws9.Cells(fila, 23).Value = categoria
            ws9.Cells(fila, 21).Value = precioCategoria
            
            ' Procesar días y horas
            If Not IsEmpty(ws2.Cells(7, columna)) Then
                ' Días laborables normales
                ws9.Cells(fila, 27).Value = ws9.Cells(fila, 27).Value + IIf(Dia = "viernes" Or Dia = "sábado" Or Dia = "domingo", 8, 9)
            Else
                Select Case LCase(horas) ' Convertir a minúsculas para simplificar comparaciones
                    Case "vacaciones", "cortaron", "c/aviso", "c/a"
                        ' No hacer nada para estos casos
                    Case "falto"
                        ws9.Cells(fila, 5).Value = "X"
                    Case "enfermo", "certif", "cert", "certificado"
                        ws9.Cells(fila, 5).Value = "X"
                        If Not (Dia = "sábado" Or Dia = "domingo") Then
                            ws9.Cells(fila, columnaBase + 10).Value = ws9.Cells(fila, columnaBase + 10).Value + IIf(Dia = "viernes", 8, 9)
                        End If
                    Case "art"
                        If Not (Dia = "sábado" Or Dia = "domingo") Then
                            ws9.Cells(fila, columnaBase + 7).Value = ws9.Cells(fila, columnaBase + 7).Value + IIf(Dia = "viernes", 8, 9)
                        End If
                    Case "lluvia"
                        ws9.Cells(fila, columnaBase).Value = ws9.Cells(fila, columnaBase).Value + 2.5
                    Case Else
                        If horas <> 0 And IsNumeric(horas) Then
                            generarHorasContador fila, columna, horas, colorHoras, Dia
                        End If
                End Select
            End If
        Next columna
        
        ' Procesar empleados especiales
        If colorEmpleado = RGB(153, 102, 0) Then
            With ws9
                .Cells(fila, columnaBase + 1).Value = ws2.Cells(fila, 21).Value
                .Cells(fila, columnaBase + 2).Value = ws2.Cells(fila, 22).Value
                .Cells(fila, columnaBase + 3).Value = ws2.Cells(fila, 35).Value
                .Cells(fila, columnaBase + 4).Value = ws2.Cells(fila, 36).Value
            End With
        End If
    Next fila
    
    Application.EnableAnimations = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
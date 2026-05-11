' ======================================================================================
' MACRO PRINCIPAL PARA COPIAR Y FORMATEAR DATOS BANCARIOS
' ======================================================================================
Sub CopiarFilasBancos()
    On Error GoTo ErrorHandler
    
    ' --- CONFIGURACIÓN INICIAL PARA MEJORAR EL RENDIMIENTO ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- DECLARACIÓN DE VARIABLES ---
    Dim ws As Worksheet
    Const HOJA_PROCESO As String = "RECUENTO TOTAL (2)"
    
    Const FORMATO_CONTADOR As String = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    Dim lastRow As Long, i As Long, destRow As Long
    Dim FechaActual As String
    Dim sumaSantander As Double, sumaMacro As Double, sumaProvincia As Double
    Dim sumaComafi As Double, sumaOtrosBancos As Double, sumaTotal As Double
    Dim sumaBanco As Double, sumaEfectivo As Double
    Dim tempVal As Variant, valorRedondeado As Double
    
    ' --- ASIGNAR HOJA DE TRABAJO (MÉTODO ROBUSTO) ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HOJA_PROCESO)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "La hoja '" & HOJA_PROCESO & "' no fue encontrada. Por favor, revise el nombre." & vbCrLf & "El proceso se cancelará.", vbCritical, "Error de Hoja"
        GoTo Fin
    End If

    ' --- OBTENER FECHA DE DEPÓSITO (FORMATO dd/mm/yyyy) ---
    Dim rawDateInput As String
    Dim dateParts() As String
    Dim tempDate As Date

    rawDateInput = InputBox("Por favor, ingrese la fecha de depósito (dd/mm/yyyy):", "Fecha de Depósito", Format(Date, "dd/mm/yyyy"))

    If rawDateInput = "" Then
        MsgBox "No se ingresó fecha. El proceso se cancelará.", vbInformation, "Proceso Cancelado"
        GoTo Fin
    End If
    
    dateParts = Split(rawDateInput, "/")

    If UBound(dateParts) = 2 And IsNumeric(dateParts(0)) And IsNumeric(dateParts(1)) And IsNumeric(dateParts(2)) Then
        On Error Resume Next
        tempDate = DateSerial(CInt(dateParts(2)), CInt(dateParts(1)), CInt(dateParts(0)))
        If Err.Number <> 0 Then
            MsgBox "La fecha '" & rawDateInput & "' no es válida. Se utilizará la fecha actual.", vbExclamation
            FechaActual = Format(Date, "dd/mm/yyyy")
            Err.Clear
        Else
            FechaActual = Format(tempDate, "dd/mm/yyyy")
        End If
        On Error GoTo ErrorHandler
    Else
        MsgBox "Formato de fecha no reconocido. Se utilizará la fecha actual.", vbExclamation
        FechaActual = Format(Date, "dd/mm/yyyy")
    End If
    
    ' --- DETERMINAR RANGOS DE TRABAJO ---
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    destRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 3
    
    ' --- APLICAR FORMATO INICIAL AL RANGO DE SALIDA ---
    With ws.Range("A" & destRow & ":P" & destRow + 500)
        .Font.Name = "Calibri"
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
    End With

    ' =================================================================================
    ' INICIO DEL PROCESAMIENTO POR BANCO
    ' =================================================================================

    ' --- PROCESAR SANTANDER ---
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = "SANTANDER" And ws.Cells(i, 5).Value > 0 Then
            tempVal = ws.Cells(i, 5).Value
            With ws
                .Range("H" & destRow).Value = .Range("B" & i).Value
                .Range("C" & destRow).Value = .Range("D" & i).Value
                .Range("G" & destRow).Value = tempVal
                .Range("G" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("N" & destRow).Value = tempVal
                .Range("N" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("E" & destRow).Value = .Range("I" & i).Value
                .Cells(destRow, 1).Value = "T"
                .Cells(destRow, 4).Value = "CUIL"
                ' <<< CORRECCIÓN APLICADA AQUÍ >>>
                With .Cells(destRow, 6)
                    .NumberFormat = "@" ' Formato Texto
                    .Value = FechaActual
                End With
                .Cells(destRow, 15).Value = "SANTANDER"
            End With
            sumaSantander = sumaSantander + tempVal
            destRow = destRow + 1
        End If
    Next i
    If sumaSantander > 0 Then
        With ws.Cells(destRow, 7)
            .Value = sumaSantander
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If
    sumaTotal = sumaTotal + sumaSantander
    destRow = destRow + 2

    ' --- PROCESAR MACRO ---
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = "MACRO" And ws.Cells(i, 5).Value > 0 Then
            tempVal = ws.Cells(i, 5).Value
            With ws
                .Range("A" & destRow).Value = .Range("A" & i).Value
                .Range("D" & destRow).Value = Val(.Range("B" & i).Value)
                .Range("C" & destRow).Value = .Range("D" & i).Value
                .Range("F" & destRow).Value = tempVal
                .Range("F" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("B" & destRow).Value = .Range("I" & i).Value
                .Range("E" & destRow).Value = .Range("K" & i).Value
                .Cells(destRow, 15).Value = "MACRO"
                With .Range("D" & destRow)
                    .NumberFormat = "0"
                    .HorizontalAlignment = xlCenter
                End With
            End With
            sumaMacro = sumaMacro + tempVal
            destRow = destRow + 1
        End If
    Next i
    If sumaMacro > 0 Then
        With ws.Cells(destRow, 6)
            .Value = sumaMacro
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If
    sumaTotal = sumaTotal + sumaMacro
    destRow = destRow + 2

    ' --- PROCESAR PROVINCIA ---
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = "PROVINCIA" And ws.Cells(i, 5).Value > 0 Then
            tempVal = ws.Cells(i, 5).Value
            With ws
                .Range("A" & destRow).Value = .Range("B" & i).Value
                .Range("C" & destRow).Value = .Range("D" & i).Value
                .Range("D" & destRow).Value = tempVal
                .Range("D" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("B" & destRow).Value = .Range("I" & i).Value
                .Cells(destRow, 15).Value = "PROVINCIA"
            End With
            sumaProvincia = sumaProvincia + tempVal
            destRow = destRow + 1
        End If
    Next i
    If sumaProvincia > 0 Then
        With ws.Cells(destRow, 4)
            .Value = sumaProvincia
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If
    sumaTotal = sumaTotal + sumaProvincia
    destRow = destRow + 2

    ' --- PROCESAR COMAFI ---
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = "COMAFI" And ws.Cells(i, 5).Value > 0 Then
            tempVal = ws.Cells(i, 5).Value
            With ws
                .Range("G" & destRow).Value = .Range("B" & i).Value
                .Range("C" & destRow).Value = .Range("D" & i).Value
                .Range("D" & destRow).Value = tempVal
                .Range("D" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("B" & destRow).Value = .Range("I" & i).Value
                ' <<< CORRECCIÓN APLICADA AQUÍ >>>
                With .Cells(destRow, 5)
                    .NumberFormat = "@" ' Formato Texto
                    .Value = FechaActual
                End With
                .Cells(destRow, 15).Value = "COMAFI"
                With .Range("G" & destRow)
                    .NumberFormat = "0"
                    .HorizontalAlignment = xlCenter
                End With
            End With
            sumaComafi = sumaComafi + tempVal
            destRow = destRow + 1
        End If
    Next i
    If sumaComafi > 0 Then
        With ws.Cells(destRow, 4)
            .Value = sumaComafi
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If
    sumaTotal = sumaTotal + sumaComafi
    destRow = destRow + 2

    ' --- PROCESAR OTROS BANCOS ---
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = "OTROS BANCOS" And ws.Cells(i, 5).Value > 0 Then
            tempVal = ws.Cells(i, 5).Value
            With ws
                .Range("H" & destRow).Value = .Range("B" & i).Value
                .Range("C" & destRow).Value = .Range("D" & i).Value
                .Range("G" & destRow).Value = tempVal
                .Range("G" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("N" & destRow).Value = tempVal
                .Range("N" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("E" & destRow).Value = .Range("I" & i).Value
                .Cells(destRow, 1).Value = "T"
                .Cells(destRow, 4).Value = "CUIL"
                ' <<< CORRECCIÓN APLICADA AQUÍ >>>
                With .Cells(destRow, 6)
                    .NumberFormat = "@" ' Formato Texto
                    .Value = FechaActual
                End With
                .Cells(destRow, 15).Value = "OTROS BANCOS"
            End With
            sumaOtrosBancos = sumaOtrosBancos + tempVal
            destRow = destRow + 1
        End If
    Next i
    If sumaOtrosBancos > 0 Then
        With ws.Cells(destRow, 7)
            .Value = sumaOtrosBancos
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If
    sumaTotal = sumaTotal + sumaOtrosBancos
    With ws.Cells(destRow + 1, 5)
        .Value = sumaTotal
        .NumberFormat = FORMATO_CONTADOR
    End With
    destRow = destRow + 4

    ' --- PROCESAR BANCO II ---
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value <> "" And ws.Cells(i, 6).Value <> 0 Then
            tempVal = ws.Cells(i, 6).Value
            With ws
                .Range("B" & destRow).Value = .Range("A" & i).Value
                .Range("C" & destRow).Value = .Range("C" & i).Value
                .Range("D" & destRow).Value = .Range("D" & i).Value
                .Range("E" & destRow).Value = tempVal
                .Range("E" & destRow).NumberFormat = FORMATO_CONTADOR
                .Range("F" & destRow).Value = .Range("I" & i).Value
                .Cells(destRow, 15).Value = "BANCO II"
            End With
            sumaBanco = sumaBanco + tempVal
            destRow = destRow + 1
        End If
    Next i
    If sumaBanco > 0 Then
        With ws.Cells(destRow, 5)
            .Value = sumaBanco
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If
    destRow = destRow + 4

    ' --- PROCESAR EFECTIVO CON CÁLCULO DE BILLETES ---
    ws.Cells(destRow - 1, "F").Value = "20000"
    ws.Cells(destRow - 1, "G").Value = "10000"
    ws.Cells(destRow - 1, "H").Value = "2000"
    ws.Cells(destRow - 1, "I").Value = "1000"
    
    For i = 2 To lastRow
        If ws.Cells(i, "G").Value > 0 Then
            With ws
                .Range("C" & destRow).Value = .Range("A" & i).Value
                .Range("D" & destRow).Value = .Range("D" & i).Value
                
                valorRedondeado = Application.WorksheetFunction.Ceiling_Math(.Cells(i, "G").Value, 100)
                .Cells(destRow, "E").Value = valorRedondeado
                .Cells(destRow, "E").NumberFormat = FORMATO_CONTADOR
                
                .Cells(destRow, "F").Value = Int(valorRedondeado / 20000)
                .Cells(destRow, "G").Value = Int(valorRedondeado / 10000)
                .Cells(destRow, "H").Value = Int(valorRedondeado / 2000)
                .Cells(destRow, "I").Value = Int(valorRedondeado / 1000)
                
                .Cells(destRow, 15).Value = "EFECTIVO"
            End With
            sumaEfectivo = sumaEfectivo + valorRedondeado
            destRow = destRow + 1
        End If
    Next i

    If sumaEfectivo > 0 Then
        With ws.Cells(destRow, 5)
            .Value = sumaEfectivo
            .NumberFormat = FORMATO_CONTADOR
        End With
    End If

    ' --- LIMPIEZA FINAL ---
    ws.Columns("A:O").AutoFit
    Application.CutCopyMode = False
    
Fin:
    ' --- RESTABLECER CONFIGURACIÓN DE EXCEL ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    ' --- MANEJADOR DE ERRORES ---
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error en Macro"
    Resume Fin
End Sub


' ======================================================================================
' MACRO PARA BORRAR CONTENIDO Y COLOREAR FILAS
' ======================================================================================
Sub BorrarContenidoYColorear()
    Dim ws As Worksheet
    ' <<< CAMBIO REALIZADO AQUÍ: Se actualizó el nombre de la hoja >>>
    Const HOJA_PROCESO As String = "RECUENTO TOTAL (2)"
    
    Dim lastRow As Long
    Dim i As Long
    
    ' --- ASIGNAR HOJA DE TRABAJO (MÉTODO ROBUSTO) ---
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(HOJA_PROCESO)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La hoja '" & HOJA_PROCESO & "' no fue encontrada. Por favor, revise el nombre." & vbCrLf & "El proceso se cancelará.", vbCritical, "Error de Hoja"
        Exit Sub
    End If
    
    ' Encontrar la última fila con datos en la columna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Buscar la primera fila en blanco a partir de la fila 2
    For i = 2 To lastRow + 1
        If IsEmpty(ws.Range("A" & i).Value) Then
            With ws.Range("A" & i & ":P2000")
                .ClearContents
                .Interior.color = RGB(211, 235, 247)
            End With
            Exit For
        End If
    Next i
End Sub

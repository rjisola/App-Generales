Sub BuscarDia()
    Dim i As Integer
    Dim Dia As String
    Dim Fecha As Date
    
    Fecha = WorksheetFunction.EoMonth(Date, -1) + 1
    Dia = Format(Fecha, "dddd")
    
    If UCase(Range("B5").Value) = "X" Then
        For i = 3 To 17
            Range(Cells(8, i), Cells(8, i)).Value = Dia
            Dia = Format(DateAdd("d", 1, Fecha), "dddd")
            Fecha = DateAdd("d", 1, Fecha)
        Next i
    End If
Range("C1:R8").Columns.AutoFit
BuscarDia2
End Sub
Sub BuscarDia2()
    Dim i As Integer
    Dim Dia As String
    Dim Fecha As Date
    
    Fecha = DateSerial(Year(Date), Month(Date), 16)
    Dia = Format(Fecha, "dddd")
    
    If UCase(Range("B6").Value) = "X" Then
        For i = 3 To 18
            If Month(Fecha) = Month(Date) Then
                Range(Cells(8, i), Cells(8, i)).Value = Dia
                Dia = Format(DateAdd("d", 1, Fecha), "dddd")
                Fecha = DateAdd("d", 1, Fecha)
            Else
                Exit For
            End If
        Next i
    End If
 Range("C1:R8").Columns.AutoFit
End Sub
Sub PintarRango()
    Dim celda As Range
    Dim UltimaFila As Long
    UltimaFila = Range("A" & Rows.Count).End(xlUp).Row 'Obtener la última fila con datos en la columna A
    
    Range("C9:R" & UltimaFila).Interior.color = RGB(255, 255, 255)
    For Each celda In Range("C8:R8") 'Recorrer las celdas del rango C8:R8
        If celda.Value = "sábado" Or celda.Value = "domingo" Then 'Verificar si contienen "sábado" o "domingo"
            Range(celda.Offset(1, 0), Cells(UltimaFila, celda.Column)).Interior.color = RGB(211, 211, 211) 'Pintar de gris claro la columna desde C9 hasta la última fila
        End If
    Next celda
    For Each celda In Range("C7:R7") 'Recorrer las celdas del rango C7:R7
        If UCase(celda.Value) = "X" Then 'Verificar si contienen "X" o "x"
            Range(celda.Offset(2, 0), Cells(UltimaFila, celda.Column)).Interior.color = RGB(255, 255, 0) 'Pintar de amarillo la columna desde C9 hasta la última fila
        End If
    Next celda
End Sub
Sub BuscarYBorrar()

Dim valor As Variant
Dim Hoja As Worksheet
Dim Matriz As Range
Dim encontrado As Range
Dim UltimaFila As Long
Dim Resultado As String 'Variable para guardar el resultado final

UltimaFila = Worksheets("ENVIO CONTADOR").Range("C" & Rows.Count).End(xlUp).Row

Resultado = "OK" 'Inicializar el resultado como OK

For i = 9 To UltimaFila
    
    valor = Worksheets("ENVIO CONTADOR").Range("C" & i).Value
    
    If Worksheets("ENVIO CONTADOR").Range("C" & i).Interior.color = RGB(255, 51, 0) Then
        
        For Each Hoja In ActiveWorkbook.Worksheets
            
            If Hoja.Name <> "Hoja99" Then
                
                If Hoja.Name = "CALCULAR HORAS" Then
                    
                    Set Matriz = Hoja.Range("A:A")
                    
                ElseIf Hoja.Name = "SUELDO_ALQ_GASTOS" Then
                    
                    Set Matriz = Hoja.Range("K:K")
                    
                ElseIf Hoja.Name = "ENVIO CONTADOR" Then
                    
                    Set Matriz = Hoja.Range("C:C")
                    
                End If
                
                Set encontrado = Matriz.Find(valor)
                
                If Not encontrado Is Nothing Then
                    
                    encontrado.EntireRow.Delete XlDeleteShiftDirection.xlShiftUp
                    
                End If
                
            End If
            
        Next Hoja
        
    Else 'Si la celda no es roja
        
        'Chequear si las columnas de las otras hojas son iguales al valor de la celda C en la fila i
        If Worksheets("CALCULAR HORAS").Range("A" & i).Value <> valor Or Worksheets("SUELDO_ALQ_GASTOS").Range("K" & i).Value <> valor Then
            
            'Cambiar el resultado a NO si hay algún valor diferente
            Resultado = "NO"
            
        End If
        
    End If
    
Next i

'Poner el resultado final en la celda U6 de la hoja ENVIO CONTADOR
Worksheets("CALCULAR HORAS").Range("U6").Value = Resultado
End Sub
Sub UltimoRegistro()
  Dim n As Long
  n = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row - 8
  ActiveSheet.Range("U4").Value = n
End Sub
Sub Limpieza()

    Call limpiarHoras
    Call limpiarImportes
    Call LimpiarValores
End Sub
Sub limpiarHoras()
    Application.Calculation = xlCalculationManual
    Dim fila As Long
    Dim columna As Long
    Dim maximoPersonas As Integer
    
    maximoPersonas = ActiveSheet.Range("u4").Value

    For fila = 9 To maximoPersonas + 8
        For columna = 19 To 25
            ActiveSheet.Cells(fila, columna).Value = Null
            ActiveSheet.Cells(fila, columna).Font.color = black
        Next columna
    Next fila
    
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub limpiarImportes()
    Application.Calculation = xlCalculationManual
    Dim fila As Long
    Dim columna As Long
    Dim maximoPersonas As Integer
    
    maximoPersonas = ActiveSheet.Range("u4").Value

    For fila = 9 To maximoPersonas + 8
        For columna = 25 To 32
        
            If columna = 31 Then
            Else
                ActiveSheet.Cells(fila, columna).Font.color = black
                ActiveSheet.Cells(fila, columna).Value = Null
            End If

        Next columna
    Next fila
    
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub Descuentos()
    Dim SUELDO_ALQ_GASTOS As Worksheet
    Dim Descuentos As Worksheet
    Dim fecha_actual As Date
    Dim ultima_fila As Long
    Dim i As Long
    Dim j As Long
    
    fecha_actual = Date
    fecha_nueva = DateSerial(Year(fecha_actual), Month(fecha_actual), 1)
    fecha_nueva1 = DateSerial(Year(fecha_actual), Month(fecha_actual), 16)
    
    
    Set SUELDO_ALQ_GASTOS = ThisWorkbook.Sheets("SUELDO_ALQ_GASTOS")
    Set Descuentos = ThisWorkbook.Sheets("Descuentos")
    
    fecha_actual = Date
    
    ultima_fila = SUELDO_ALQ_GASTOS.Cells(Rows.Count, "K").End(xlUp).Row
    
    SUELDO_ALQ_GASTOS.Range("P9:P" & ultima_fila).ClearContents ' Borrar la columna P desde la fila 9 hasta el último registro en K.
    
    For i = 9 To ultima_fila
    
        For j = 9 To 100
            If SUELDO_ALQ_GASTOS.Cells(i, "K").Value = Descuentos.Cells(j, "C").Value Then
                        
                If Descuentos.Cells(j, "E").Value = fecha_nueva And fecha_nueva <= Date Then
                
                    SUELDO_ALQ_GASTOS.Cells(i, "P").Value = Descuentos.Cells(j, "D").Value
                    Descuentos.Range("C" & j & ":E" & j).Interior.ColorIndex = 6 ' Amarillo
                ElseIf Descuentos.Cells(j, "E").Value = fecha_nueva1 And fecha_nueva1 <= Date Then
                    SUELDO_ALQ_GASTOS.Cells(i, "P").Value = Descuentos.Cells(j, "D").Value
                    Descuentos.Range("C" & j & ":E" & j).Interior.ColorIndex = 6 ' Amarillo
                End If
            End If
        Next j
        
        If i Mod 1000 = 0 Then ' Para evitar que la macro se bloquee si hay demasiados registros.
            DoEvents
        End If
        
        If i Mod 100 = 0 Then ' Para actualizar el estado de la macro cada 100 registros.
            Application.StatusBar = "Procesando fila " & i & " de " & ultima_fila & "..."
        End If
        
    Next i
    
    Application.StatusBar = False ' Restablecer la barra de estado al finalizar.
    Call CopiarDatos
End Sub
Sub CopiarDatos()
    Dim wb As Workbook
    Dim wsArreglos As Worksheet
    Dim wsSueldos As Worksheet
    Dim lastRowArreglos As Long
    Dim lastRowSueldos As Long
    Dim i As Long
    Dim j As Long
    
    ' Establecer referencias al libro y hoja de trabajo necesarios
    Set wb = ThisWorkbook
    Set wsArreglos = wb.Sheets("ARREGLOS_ALQUILERES")
    
    ' Verificar si la hoja "SUELDO_ALQ_GASTOS" existe en el libro
    On Error Resume Next
    Set wsSueldos = wb.Sheets("SUELDO_ALQ_GASTOS")
    On Error GoTo 0
    
    If wsSueldos Is Nothing Then
        MsgBox "No se encontró la hoja ""SUELDO_ALQ_GASTOS"" en el libro.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la última fila con datos en cada hoja
    lastRowArreglos = wsArreglos.Cells(wsArreglos.Rows.Count, "H").End(xlUp).Row
    lastRowSueldos = wsSueldos.Cells(wsSueldos.Rows.Count, "B").End(xlUp).Row
    
    ' Recorrer los registros en ARREGLOS_ALQUILERES y pegar los datos en SUELDOS_ALQ_GASTOS
    For i = 9 To lastRowArreglos
        ' Obtener el número de legajo de la columna H en ARREGLOS_ALQUILERES
        Dim legajoArreglos As Long
        legajoArreglos = wsArreglos.Cells(i, "H").Value
        
        ' Buscar el número de legajo en la columna B de SUELDOS_ALQ_GASTOS
        For j = 9 To lastRowSueldos
            Dim legajoSueldos As Long
            legajoSueldos = wsSueldos.Cells(j, "B").Value
            
            ' Si el número de legajo coincide, copiar el valor de la columna M de ARREGLOS_ALQUILERES a la columna L de SUELDOS_ALQ_GASTOS
            If legajoArreglos = legajoSueldos Then
                wsSueldos.Cells(j, "L").Value = wsArreglos.Cells(i, "M").Value
                Exit For ' Salir del bucle de búsqueda una vez que se haya encontrado una coincidencia
            End If
        Next j
    Next i
    
    ' Limpiar variables y liberar memoria
    Set wsSueldos = Nothing
    Set wsArreglos = Nothing
    Set wb = Nothing
End Sub

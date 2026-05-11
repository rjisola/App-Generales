Sub CopiarHorasDeposito()
    Dim hojaHoras As Worksheet
    Dim hojaCalculo As Worksheet
    Dim ultimaFilaHoras As Long
    Dim ultimaFilaCalculo As Long
    Dim i As Long, j As Long, k As Long
    Dim valor As Variant
    
    ' Asignar las hojas a variables
    Set hojaHoras = ThisWorkbook.Sheets("HorasDeposito")
    Set hojaCalculo = ThisWorkbook.Sheets("CALCULAR HORAS")
    
    ' Obtener la última fila con datos en cada hoja
    ultimaFilaHoras = hojaHoras.Cells(Rows.Count, 1).End(xlUp).Row
    ultimaFilaCalculo = hojaCalculo.Cells(Rows.Count, 38).End(xlUp).Row ' Columna AL = 38
    
    ' Comparar los valores de las columnas A y AL
    For i = 6 To ultimaFilaHoras
        For j = 9 To ultimaFilaCalculo
            If hojaHoras.Cells(i, 1).Value = hojaCalculo.Cells(j, 38).Value Then
                ' Verificar si el rango en la hoja "CALCULAR HORAS" está vacío o contiene ceros
                For k = 3 To 18 ' Columnas C a R
                    ' Si la celda ya contiene un valor diferente de vacío o cero, saltar a la siguiente celda
                    If hojaCalculo.Cells(j, k).Value = "" Or hojaCalculo.Cells(j, k).Value = 0 Then
                        ' Copiar el valor correspondiente desde la hoja Horas
                        valor = hojaHoras.Cells(i, k + 1).Value
                        
                        ' Reemplazar ciertos valores
                        Select Case UCase(valor)
                            Case "CERT", "ENFERMO"
                                valor = "CERTIF"
                            Case "PERMISO", "C/A"
                                valor = "C/AVISO"
                            Case "FALTO"
                                valor = "FALTO"
                            Case "VAC"
                                valor = "VACACIONES"
                        End Select
                        
                        ' Si el valor es texto, convertirlo a mayúsculas
                        If Not IsNumeric(valor) Then
                            valor = UCase(valor)
                        End If
                        
                        hojaCalculo.Cells(j, k).Value = valor
                    End If
                Next k
                Exit For ' Salir del bucle interior si se encuentra una coincidencia
            End If
        Next j
    Next i
End Sub

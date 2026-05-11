Sub CompararYActualizarHoja5()


Dim indice As Worksheet
Dim datos As Worksheet
Dim ultimaFilaIndice As Long
Dim ultimaFilaDatos As Long
Dim i As Long
Dim j As Long
Dim encontrado As Boolean

Set indice = ThisWorkbook.Sheets("ENVIO CONTADOR")
Set datos = ThisWorkbook.Sheets("ARREGLOS_ALQUILERES")

ultimaFilaIndice = indice.Cells(Rows.Count, "C").End(xlUp).Row
ultimaFilaDatos = datos.Cells(Rows.Count, "C").End(xlUp).Row

For i = ultimaFilaDatos To 9 Step -1
    encontrado = False
    For j = 9 To ultimaFilaIndice
        If datos.Cells(i, "C").Value = indice.Cells(j, "C").Value And datos.Cells(i, "F").Value = indice.Cells(j, "W").Value Then
            encontrado = True
            Exit For
        End If
    Next j
    'If Not encontrado Then
        'datos.Rows(i).Delete
        'ultimaFilaDatos = ultimaFilaDatos - 1
    'End If
Next i

For j = 9 To ultimaFilaIndice
    encontrado = False
    For i = 9 To ultimaFilaDatos
        If indice.Cells(j, "C").Value = datos.Cells(i, "C").Value Then
            encontrado = True
            Exit For
        End If
    Next i
    If Not encontrado Then
        datos.Rows(ultimaFilaDatos + 1).Insert
        datos.Cells(ultimaFilaDatos + 1, "C").Value = indice.Cells(j, "C").Value
        datos.Cells(ultimaFilaDatos + 1, "B").Value = indice.Cells(j, "B").Value
        ultimaFilaDatos = ultimaFilaDatos + 1
    End If
Next j

datos.Range("A9:E" & ultimaFilaDatos).Sort Key1:=datos.Range("C9:C" & ultimaFilaDatos), _
                                            Order1:=xlAscending, Header:=xlNo

Dim k As Long

For k = 9 To ultimaFilaDatos

    If Application.WorksheetFunction.CountIf(indice.Range("C:C"), datos.Cells(k, "C").Value) > 0 Then

        datos.Cells(k, "F").Value = "Ok"

    End If

Next k
'Selecciona la columna de fórmulas que desea actualizar en la hoja 2
    Sheets("Comprobar Lista").Select
    Range("C3:E400").Select
    
    'Actualiza las fórmulas en la hoja 2
    ActiveSheet.Calculate
Application.EnableAnimations = True

End Sub



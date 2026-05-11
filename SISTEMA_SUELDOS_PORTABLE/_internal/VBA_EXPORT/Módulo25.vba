Sub CompararYActualizarHoja2()
    Dim indice As Worksheet
    Dim datos As Worksheet
    Dim ultimaFilaIndice As Long
    Dim ultimaFilaDatos As Long
    Dim i As Long
    Dim j As Long
    Dim encontrado As Boolean

    Set indice = ThisWorkbook.Sheets("ENVIO CONTADOR")
    Set datos = ThisWorkbook.Sheets("CALCULAR HORAS")

    ultimaFilaIndice = indice.Cells(Rows.Count, "C").End(xlUp).Row
    ultimaFilaDatos = datos.Cells(Rows.Count, "A").End(xlUp).Row

    For i = ultimaFilaDatos To 9 Step -1
        encontrado = False
        For j = 9 To ultimaFilaIndice
            If datos.Cells(i, "A").Value = indice.Cells(j, "C").Value Then
                encontrado = True
                Exit For
            End If
        Next j
        If Not encontrado Then
            datos.Rows(i).Delete
            ultimaFilaDatos = ultimaFilaDatos - 1
        End If
    Next i

    For j = 9 To ultimaFilaIndice
        encontrado = False
        For i = 9 To ultimaFilaDatos
            If indice.Cells(j, "C").Value = datos.Cells(i, "A").Value Then
                encontrado = True
                Exit For
            End If
        Next i
        If Not encontrado Then
            datos.Rows(ultimaFilaDatos + 1).Insert Shift:=xlDown
            datos.Cells(ultimaFilaDatos + 1, "A").Value = indice.Cells(j, "C").Value
            datos.Cells(ultimaFilaDatos + 1, "B").Value = indice.Cells(j, "W").Value
            datos.Cells(ultimaFilaDatos + 1, "AL").Value = indice.Cells(j, "B").Value
            ultimaFilaDatos = ultimaFilaDatos + 1
        End If
    Next j

    datos.Range("A9:AM" & ultimaFilaDatos).Sort Key1:=datos.Range("A9:A" & ultimaFilaDatos), Order1:=xlAscending, Header:=xlNo

    Dim k As Long

    For k = 9 To ultimaFilaDatos
        If Application.WorksheetFunction.CountIf(indice.Range("C:C"), datos.Cells(k, "A").Value) > 0 Then
            datos.Cells(k, "AM").Value = "Ok"
        End If
    Next k

    Call CompararYActualizarHoja3
End Sub

Sub CompararYActualizarHoja3()
    Dim indice As Worksheet
    Dim datos As Worksheet
    Dim ultimaFilaIndice As Long
    Dim ultimaFilaDatos As Long
    Dim i As Long
    Dim j As Long
    Dim encontrado As Boolean

    Set indice = ThisWorkbook.Sheets("ENVIO CONTADOR")
    Set datos = ThisWorkbook.Sheets("SUELDO_ALQ_GASTOS")

    ultimaFilaIndice = indice.Cells(Rows.Count, "C").End(xlUp).Row
    ultimaFilaDatos = datos.Cells(Rows.Count, "K").End(xlUp).Row

    For i = ultimaFilaDatos To 9 Step -1
        encontrado = False
        For j = 9 To ultimaFilaIndice
            If datos.Cells(i, "K").Value = indice.Cells(j, "C").Value Then
                encontrado = True
                Exit For
            End If
        Next j
        If Not encontrado Then
            datos.Rows(i).Delete
            ultimaFilaDatos = ultimaFilaDatos - 1
        End If
    Next i

    For j = 9 To ultimaFilaIndice
        encontrado = False
        For i = 9 To ultimaFilaDatos
            If indice.Cells(j, "C").Value = datos.Cells(i, "K").Value Then
                encontrado = True
                Exit For
            End If
        Next i
        If Not encontrado Then
            datos.Rows(ultimaFilaDatos + 1).Insert Shift:=xlDown
            datos.Cells(ultimaFilaDatos + 1, "K").Value = indice.Cells(j, "C").Value
            datos.Cells(ultimaFilaDatos + 1, "Z").Value = indice.Cells(j, "W").Value
            datos.Cells(ultimaFilaDatos + 1, "B").Value = indice.Cells(j, "B").Value
            datos.Cells(ultimaFilaDatos + 1, "C").Value = indice.Cells(j, "A").Value
            ultimaFilaDatos = ultimaFilaDatos + 1
        End If
    Next j

    datos.Range("A9:AM" & ultimaFilaDatos).Sort Key1:=datos.Range("K9:K" & ultimaFilaDatos), Order1:=xlAscending, Header:=xlNo

    Dim k As Long
    For k = 9 To ultimaFilaDatos
        If Application.WorksheetFunction.CountIf(indice.Range("C:C"), datos.Cells(k, "K").Value) > 0 Then
            datos.Cells(k, "AM").Value = "Ok"
        End If
    Next k
End Sub

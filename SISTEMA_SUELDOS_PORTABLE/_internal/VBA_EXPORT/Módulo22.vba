Sub NOIMPRIMIRperdida()
Dim i As Long
Dim celda As Range
For i = 9 To Range("AL" & Rows.Count).End(xlUp).Row
    Set celda = Range("A" & i)
    If celda.Interior.color = RGB(255, 204, 102) And Range("AL" & i) = "NO IMPRIMIR" Then
        celda.Interior.color = RGB(255, 51, 0)
    End If
Next i
End Sub
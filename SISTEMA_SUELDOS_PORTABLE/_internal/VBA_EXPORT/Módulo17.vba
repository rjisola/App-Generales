Sub Gastos()
Dim i As Integer
Dim j As Integer
Dim lastRow As Long
Dim Hoja2 As Worksheet
Dim Hoja3 As Worksheet

Set Hoja2 = ThisWorkbook.Sheets("SUELDO_ALQ_GASTOS")
Set Hoja3 = ThisWorkbook.Sheets("ARREGLOS_ALQUILERES")

lastRow = Hoja2.Range("K" & Rows.Count).End(xlUp).Row

For i = 9 To lastRow
    For j = 9 To lastRow
    If Hoja2.Cells(i, "N").Interior.color = RGB(255, 255, 0) Then
        If Hoja2.Cells(i, "K").Value = Hoja3.Cells(j, "C").Value Then
            
                Hoja2.Cells(i, "N").Value = Hoja3.Cells(j, "D").Value
            End If
          
        End If
    Next j
Next i
Call Descuentos
End Sub

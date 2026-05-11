Sub mergearCeldas(contador, columna1, columna2, Hoja)
    If Not Hoja.Cells(contador, columna1).MergeCells Then
        Hoja.Range(Hoja.Cells(contador, columna1), Hoja.Cells(contador, columna2)).Merge
    End If
    Hoja.Cells(contador, columna1).NumberFormat = "$#,##0.00"
    Hoja.Cells(contador, columna1).Font.Bold = True
    Hoja.Cells(contador, columna1).Font.Size = 10.5
    Hoja.Cells(contador, columna1).HorizontalAlignment = xlLeft
    Hoja.Cells(contador, columna1).VerticalAlignment = xlCenter
End Sub

Sub unMergearCeldas(contador, columna1, columna2, Hoja)

    If Hoja.Cells(contador, columna1).MergeCells Then
        Hoja.Range(Hoja.Cells(contador, columna1), Hoja.Cells(contador, columna2)).UnMerge
    End If
    Hoja.Cells(contador, columna2).Borders.Weight = xlThin
    Hoja.Cells(contador, columna2).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Cells(contador, columna1).Borders.Weight = xlThin
    Hoja.Cells(contador, columna1).NumberFormat = "0.0"
    Hoja.Cells(contador, columna2).NumberFormat = "$#,##0.00"
    Hoja.Range(Hoja.Cells(contador, columna1), Hoja.Cells(contador, columna2)).Font.Bold = True
    Hoja.Range(Hoja.Cells(contador, columna1), Hoja.Cells(contador, columna2)).Font.Size = 10.5
    Hoja.Range(Hoja.Cells(contador, columna1), Hoja.Cells(contador, columna2)).HorizontalAlignment = xlLeft
    Hoja.Range(Hoja.Cells(contador, columna1), Hoja.Cells(contador, columna2)).VerticalAlignment = xlCenter

End Sub

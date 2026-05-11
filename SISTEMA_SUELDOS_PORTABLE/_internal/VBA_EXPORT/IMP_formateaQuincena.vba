Sub formatearQuincena(contador, columna, Hoja)
    
    Hoja.Range(Hoja.Cells(contador, columna), Hoja.Cells(contador + 2, columna)).Merge
    Hoja.Range(Hoja.Cells(contador, columna + 1), Hoja.Cells(contador + 2, columna + 2)).Merge
    
    Hoja.Range(Hoja.Cells(contador, columna), Hoja.Cells(contador + 2, columna)).Borders.Weight = xlThin

    Hoja.Cells(contador, columna + 1).NumberFormat = "$#,##0.00"
    Hoja.Cells(contador, columna + 1).Font.Bold = True
    Hoja.Cells(contador, columna + 1).Font.Size = 16
    Hoja.Cells(contador, columna + 1).HorizontalAlignment = xlCenter
    Hoja.Cells(contador, columna + 1).VerticalAlignment = xlCenter
    Hoja.Cells(contador, columna).HorizontalAlignment = xlCenter
    Hoja.Cells(contador, columna).VerticalAlignment = xlCenter
    Hoja.Cells(contador, columna).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Cells(contador + 1, columna).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Cells(contador + 2, columna).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Cells(contador, columna + 1).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Cells(contador + 1, columna + 1).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Cells(contador + 2, columna + 1).Borders(xlEdgeRight).Weight = xlMedium

End Sub

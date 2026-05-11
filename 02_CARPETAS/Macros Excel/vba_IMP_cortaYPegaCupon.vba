Sub cortarYPegarCupon(filaLlena, filaVacia, columna, desplazamiento, columnaVacia)
    
    If filaLlena <> 0 Then
        ActiveSheet.Range(ActiveSheet.Cells(filaLlena, columna), ActiveSheet.Cells(filaLlena + desplazamiento, columna + 2)).Copy ActiveSheet.Cells(filaVacia, columnaVacia)
        ActiveSheet.Range(ActiveSheet.Cells(filaLlena, columna), ActiveSheet.Cells(filaLlena + desplazamiento, columna + 2)).Copy
        ActiveSheet.Cells(filaVacia, columnaVacia).PasteSpecial xlFormats
        ActiveSheet.Range(ActiveSheet.Cells(filaLlena, columna), ActiveSheet.Cells(filaLlena + desplazamiento, columna + 2)).ClearContents
        ActiveSheet.Range(ActiveSheet.Cells(filaLlena, columna), ActiveSheet.Cells(filaLlena + desplazamiento, columna + 2)).Interior.color = RGB(255, 255, 255)
    End If

    
End Sub

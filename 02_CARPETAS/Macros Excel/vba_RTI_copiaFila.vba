Sub copiarFila(filaLlena, filaVacia)
    Dim corrimiento As Integer
    
    corrimiento = 5
    

    ActiveSheet.Range(ActiveSheet.Cells(filaLlena, 1), ActiveSheet.Cells(filaLlena, 5 + corrimiento)).Copy
    ActiveSheet.Cells(filaVacia, 1).PasteSpecial xlValues
    ActiveSheet.Cells(filaVacia, 1).PasteSpecial xlFormats
    ActiveSheet.Range(ActiveSheet.Cells(filaLlena, 1), ActiveSheet.Cells(filaLlena, 5 + corrimiento)).ClearContents
    ActiveSheet.Cells(filaLlena, 1).Interior.color = RGB(221, 235, 247)

End Sub

Sub vaciarRecuentoTotal()
    ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(200, 5 + 6)).ClearContents
    ActiveSheet.Range(ActiveSheet.Cells(2, 1), ActiveSheet.Cells(200, 5 + 6)).Interior.color = RGB(211, 235, 247)
End Sub

Sub BorrarEnvioContador()

    For fila = 9 To Hoja2.Cells(4, 21).Value + 9
        Hoja9.Range(Hoja9.Cells(fila, 4), Hoja9.Cells(fila, 19)).Value = ""
    Next fila
Range("AA9:AA95").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-90
    Range("AA9").Select
End Sub

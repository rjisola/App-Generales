Sub informarError()

    ActiveSheet.Range("S4").Interior.color = RGB(255, 0, 0)
    ActiveSheet.Range("S4").Font.color = RGB(255, 255, 255)
    ActiveSheet.Range("S4").Value = "HAY ERROR"

End Sub

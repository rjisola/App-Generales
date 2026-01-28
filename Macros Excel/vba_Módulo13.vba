Sub CopiarRECUENTO()
    Sheets("RECUENTO TOTAL").Range("K1:A" & Range("A" & Rows.Count).End(xlUp).Row).Copy
    Sheets("RECUENTO TOTAL (2)").Range("A1").PasteSpecial xlPasteFormats
    Sheets("RECUENTO TOTAL (2)").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    ' Activa la hoja "RECUENTO TOTAL (2)" para posicionar el cursor
    Sheets("RECUENTO TOTAL (2)").Activate
End Sub
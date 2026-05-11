Sub MacroVBA()
  Dim n As Long
  n = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row - 8
  ActiveSheet.Range("U4").Value = n
End Sub
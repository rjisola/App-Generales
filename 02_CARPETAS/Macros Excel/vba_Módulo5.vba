
  Option Explicit
                                                                                                                   
  Public Sub ExportarMedidas()
      Dim sh As Worksheet, out As Worksheet
      Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
      Dim r As Long, c As Long, i As Long
                                                                                                                   
      Set sh = ActiveSheet
                                                                                                                   
      On Error Resume Next
      Application.DisplayAlerts = False
      Worksheets("MEDIDAS").Delete
      Application.DisplayAlerts = True
      On Error GoTo 0
      Set out = Worksheets.Add
      out.Name = "MEDIDAS"
                                                                                                                   
      With sh.UsedRange
          r1 = .Row: r2 = .Rows(.Rows.Count).Row
          c1 = .Column: c2 = .Columns(.Columns.Count).Column
      End With
                                                                                                                   
      out.Range("A1").Value = "Fila"
      out.Range("B1").Value = "Altura (pt)"
      out.Range("D1").Value = "Columna"
      out.Range("E1").Value = "Ancho (unid. Excel)"
                                                                                                                   
      i = 2
      For r = r1 To r2
          out.Cells(i, 1).Value = r
          out.Cells(i, 2).Value = sh.Rows(r).RowHeight
          i = i + 1
      Next r
                                                                                                                   
      out.Cells(i, 1).Value = "TOTAL ALTO"
      out.Cells(i, 2).Value = Application.WorksheetFunction.Sum(out.Range("B2:B" & i - 1))
                                                                                                                   
      i = 2
      For c = c1 To c2
          out.Cells(i, 4).Value = c
          out.Cells(i, 5).Value = sh.Columns(c).ColumnWidth
          i = i + 1
      Next c
                                                                                                                   
      out.Columns("A:E").AutoFit
      out.Activate
      MsgBox "Medidas exportadas a la hoja 'MEDIDAS'.", vbInformation
  End Sub
                                                                                                                   
  Public Sub AplicarAlturasDesdeMedidas()
      Dim sh As Worksheet, src As Worksheet
      Dim lastRow As Long, r As Long, fila As Variant, altura As Variant
                                                                                                                   
      On Error Resume Next
      Set src = Worksheets("MEDIDAS")
      On Error GoTo 0
      If src Is Nothing Then
          MsgBox "No existe la hoja 'MEDIDAS'. Ejecute primero ExportarMedidas.", vbExclamation
          Exit Sub
      End If
                                                                                                                   
      Set sh = ActiveSheet
      lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).Row
                                                                                                                   
      For r = 2 To lastRow
          fila = src.Cells(r, "A").Value
          altura = src.Cells(r, "B").Value
          If IsNumeric(fila) And IsNumeric(altura) Then
              On Error Resume Next
              sh.Rows(CLng(fila)).RowHeight = CDbl(altura)
              On Error GoTo 0
          End If
      Next r
                                                                                                                   
      MsgBox "Alturas aplicadas desde 'MEDIDAS'.", vbInformation
  End Sub
                                                                                                                   
  Public Sub FijarAlturaUniforme()
      Dim h As Variant
      h = Application.InputBox("Altura a aplicar (en puntos):", "Fijar altura", 15, Type:=1)
      If h = False Then Exit Sub
      If h <= 0 Then
          MsgBox "Valor inválido.", vbExclamation
          Exit Sub
      End If
      ActiveSheet.Rows.RowHeight = h
      MsgBox "Altura uniforme aplicada: " & h & " pt.", vbInformation
  End Sub
                                                                                                                   
  Public Sub CompactarFilasVacias()
      Dim hVacia As Variant, hNoVacia As Variant
      Dim r As Long, r1 As Long, r2 As Long, sh As Worksheet
                                                                                                                   
      Set sh = ActiveSheet
      hVacia = Application.InputBox("Altura para filas VACÍAS (pt):", "Compactar vacías", 6, Type:=1)
      If hVacia = False Then Exit Sub
      hNoVacia = Application.InputBox("Altura para filas con contenido (pt):", "Compactar vacías", 18, Type:=1)
      If hNoVacia = False Then Exit Sub
                                                                                                                   
      With sh.UsedRange
          r1 = .Row: r2 = .Rows(.Rows.Count).Row
      End With
                                                                                                                   
      Application.ScreenUpdating = False
      For r = r1 To r2
          If Application.WorksheetFunction.CountA(sh.Rows(r)) = 0 Then
              sh.Rows(r).RowHeight = hVacia
          Else
              sh.Rows(r).RowHeight = hNoVacia
          End If
      Next r
      Application.ScreenUpdating = True
                                                                                                                   
      MsgBox "Filas vacías compactadas a " & hVacia & " pt; resto a " & hNoVacia & " pt.", vbInformation
  End Sub
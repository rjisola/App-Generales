Sub ImprimirSobres()
    Dim wsData As Worksheet
    Dim wsPrint As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim printRow As Long
    Dim legajo As String
    Dim nombre As String
    Dim texto As String

    ' ==========================================
    ' CONFIGURACIÓN
    ' ==========================================
    ' Nombre de la hoja donde están los datos
    Const DATA_SHEET_NAME As String = "Hoja1"
    ' Nombre de la hoja temporal para imprimir
    Const PRINT_SHEET_NAME As String = "ImpresionSobres"
    ' Margen superior en cm
    Const MARGIN_TOP_CM As Double = 3.0
    
    ' Verificar si existe la hoja de datos
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    On Error GoTo 0
    
    If wsData Is Nothing Then
        MsgBox "No se encontró la hoja '" & DATA_SHEET_NAME & "'.", vbCritical
        Exit Sub
    End If

    ' Encontrar la última fila con datos en la columna A
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No hay datos suficientes para procesar.", vbExclamation
        Exit Sub
    End If

    ' Crear o limpiar hoja de impresión
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Set wsPrint = ThisWorkbook.Sheets(PRINT_SHEET_NAME)
    On Error GoTo 0
    
    If wsPrint Is Nothing Then
        Set wsPrint = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsPrint.Name = PRINT_SHEET_NAME
    Else
        wsPrint.Cells.Clear
        wsPrint.ResetAllPageBreaks
        wsPrint.DrawingObjects.Delete
    End If

    ' ==========================================
    ' CONFIGURACIÓN DE PÁGINA (SOBRE C5)
    ' ==========================================
    With wsPrint.PageSetup
        ' xlPaperEnvelopeC5 = 28 (162mm x 229mm)
        On Error Resume Next
        .PaperSize = xlPaperEnvelopeC5 
        On Error GoTo 0
        
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(MARGIN_TOP_CM)
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(2)
        .BottomMargin = Application.CentimetersToPoints(2)
        .CenterHorizontally = True
        .CenterVertically = False
        
        ' Áreas de impresión limpias
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With

    ' ==========================================
    ' GENERACIÓN DE ETIQUETAS
    ' ==========================================
    printRow = 1
    
    For i = 2 To lastRow ' Asumiendo encabezados en fila 1
        legajo = Trim(wsData.Cells(i, 1).Value)
        nombre = Trim(wsData.Cells(i, 2).Value)
        
        If legajo <> "" Or nombre <> "" Then
            ' Formato: (leg) Nombre y apellido
            texto = "(" & legajo & ") " & nombre
            
            With wsPrint.Cells(printRow, 1)
                .Value = texto
                
                ' Formato de fuente
                .Font.Name = "Arial"
                .Font.Size = 14
                .Font.Bold = True
                
                ' Ajuste de celda para asegurar visibilidad
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
                .WrapText = True
                
                ' Altura de fila suficiente
                .RowHeight = 30 
            End With
            
            ' Insertar salto de página horizontal después de cada registro
            If i < lastRow Then
                wsPrint.HPageBreaks.Add Before:=wsPrint.Cells(printRow + 1, 1)
            End If
            
            printRow = printRow + 1
        End If
    Next i
    
    wsPrint.Columns("A:A").ColumnWidth = 50
    
    Application.ScreenUpdating = True
    
    ' Confirmación y Vista Previa
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Se ha generado la hoja '" & PRINT_SHEET_NAME & "' con " & (printRow - 1) & " sobres." & vbNewLine & vbNewLine & _
                       "¿Desea ver la vista preliminar de impresión ahora?", vbYesNo + vbQuestion, "Proceso Completado")
                       
    wsPrint.Activate
    If respuesta = vbYes Then
        wsPrint.PrintPreview
    End If

End Sub
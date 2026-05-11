
Sub CopiarRangoYBuscarColor()
    Dim LibroDestino As Workbook
    Dim HojaDestino As Worksheet
    Dim HojaOrigen As Worksheet
    Dim RangoOrigen As Range
    Dim FilePath As String
    Dim UltimaFila As Long
    Dim totalFilas As Long
    Dim totalColumnas As Long
    Dim fila As Long
    Dim col As Long
    Dim colIndex As Long
    Dim COLOR_ROJO As Long
    Dim COLOR_CREMA As Long
    Dim COLOR_DATOS_RESALTADOS As Long
    Dim columnasPermitidas As Variant
    Dim columnasSoloColor As Variant
    Dim columnasColorFijas As Variant
    Dim colorA As Long
    Dim colorB As Long
    Dim cabecera As String
    Dim colColor As Variant
    Dim colColorFija As Variant
    Dim helperColIndex As Long
    Dim firstCremaRow As Long
    Dim sortRange As Range
    Dim dataRange As Range
    Dim dataStartRow As Long
    Dim sourceRow As Long

    Const PRIMERA_FILA_ORIGEN As Long = 7
    Const HEADER_ROWS As Long = 2

On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set HojaOrigen = ActiveSheet

    COLOR_ROJO = RGB(255, 51, 0)
    COLOR_CREMA = RGB(251, 226, 213)
    COLOR_DATOS_RESALTADOS = RGB(251, 51, 0)
    columnasPermitidas = Array("A", "B", "T", "U", "V", "X")
    columnasSoloColor = Array("H", "I", "J", "K", "AA", "AB")
    columnasColorFijas = Array("F", "M", "P")

    ' Ultima fila con datos en columna C del origen
    UltimaFila = HojaOrigen.Cells(HojaOrigen.Rows.Count, "C").End(xlUp).Row
    If UltimaFila < PRIMERA_FILA_ORIGEN Then GoTo Salir

    ' Rango de origen: B7:AG hasta la ultima fila encontrada
    Set RangoOrigen = HojaOrigen.Range("B" & PRIMERA_FILA_ORIGEN & ":AG" & UltimaFila)
    totalFilas = RangoOrigen.Rows.Count
    totalColumnas = RangoOrigen.Columns.Count

    ' Archivo destino
    FilePath = "C:\Users\rjiso\OneDrive\Escritorio\HORAS CONTADOR.xlsx"

    ' Crear u abrir el archivo destino
    If Dir(FilePath) = "" Then
        Set LibroDestino = Workbooks.Add(xlWBATWorksheet)
        Set HojaDestino = LibroDestino.Sheets(1)
        HojaDestino.Name = "HORAS CONTADOR"
        LibroDestino.SaveAs Filename:=FilePath, FileFormat:=xlOpenXMLWorkbook
    Else
        Set LibroDestino = Workbooks.Open(FilePath)
        On Error Resume Next
        Set HojaDestino = LibroDestino.Sheets("HORAS CONTADOR")
        On Error GoTo ErrHandler
        If HojaDestino Is Nothing Then
            Set HojaDestino = LibroDestino.Worksheets.Add(After:=LibroDestino.Sheets(LibroDestino.Sheets.Count))
            HojaDestino.Name = "HORAS CONTADOR"
        End If
    End If

    ' Copiar desde A1 en el destino
    HojaDestino.Cells.Clear
    RangoOrigen.Copy
    With HojaDestino.Range("A1")
        .PasteSpecial Paste:=xlPasteAll
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With
    Application.CutCopyMode = False

    For fila = 1 To totalFilas
        sourceRow = PRIMERA_FILA_ORIGEN + fila - 1
        For Each colColorFija In columnasColorFijas
            HojaDestino.Range(colColorFija & fila).Interior.color = _
                HojaOrigen.Range(colColorFija & sourceRow).DisplayFormat.Interior.color
        Next colColorFija
    Next fila


    dataStartRow = HEADER_ROWS + 1

    If totalFilas > HEADER_ROWS Then
        ' Ajustar datos segun reglas de color
        For fila = dataStartRow To totalFilas
            colorA = HojaDestino.Cells(fila, "A").Interior.color
            colorB = HojaDestino.Cells(fila, "B").Interior.color

            If colorA <> COLOR_ROJO And colorA <> COLOR_CREMA Then
                HojaDestino.Cells(fila, "A").Interior.Pattern = xlNone
            End If

            If colorB <> COLOR_ROJO And colorB <> COLOR_CREMA Then
                HojaDestino.Cells(fila, "B").Interior.Pattern = xlNone
            End If

            If colorB = COLOR_CREMA Then
                For col = 1 To totalColumnas
                    cabecera = Split(HojaDestino.Cells(1, col).Address(False, False), "1")(0)
                    If IsError(Application.Match(cabecera, columnasPermitidas, 0)) Then
                        HojaDestino.Cells(fila, col).ClearContents
                    End If
                Next col
            End If

            For Each colColor In columnasSoloColor
                colIndex = HojaDestino.Columns(colColor).Column
                If HojaDestino.Cells(fila, colIndex).Interior.color <> COLOR_DATOS_RESALTADOS Then
                    HojaDestino.Cells(fila, colIndex).ClearContents
                End If
            Next colColor
        Next fila

        ' Ordenar por columna A (ascendente)
        Set dataRange = HojaDestino.Range(HojaDestino.Cells(dataStartRow, 1), HojaDestino.Cells(totalFilas, totalColumnas))
        dataRange.Sort Key1:=HojaDestino.Cells(dataStartRow, 1), Order1:=xlAscending, Header:=xlNo

        ' Marcar filas crema y reagruparlas al final
        helperColIndex = totalColumnas + 1
        For fila = 1 To totalFilas
            HojaDestino.Cells(fila, helperColIndex).Value = 0
        Next fila
        For fila = dataStartRow To totalFilas
            HojaDestino.Cells(fila, helperColIndex).Value = IIf(HojaDestino.Cells(fila, "B").Interior.color = COLOR_CREMA, 1, 0)
        Next fila

        Set sortRange = HojaDestino.Range(HojaDestino.Cells(dataStartRow, 1), HojaDestino.Cells(totalFilas, helperColIndex))
        sortRange.Sort Key1:=HojaDestino.Cells(dataStartRow, helperColIndex), Order1:=xlAscending, _
                       Key2:=HojaDestino.Cells(dataStartRow, 1), Order2:=xlAscending, Header:=xlNo

        firstCremaRow = 0
        For fila = dataStartRow To totalFilas
            If HojaDestino.Cells(fila, helperColIndex).Value = 1 Then
                firstCremaRow = fila
                Exit For
            End If
        Next fila

        If firstCremaRow > 0 Then
            HojaDestino.Rows(firstCremaRow).Insert Shift:=xlDown
            HojaDestino.Rows(firstCremaRow).ClearContents
            HojaDestino.Rows(firstCremaRow).Interior.Pattern = xlNone
            HojaDestino.Cells(firstCremaRow, helperColIndex).ClearContents
            totalFilas = totalFilas + 1
        End If

        HojaDestino.Columns(helperColIndex).Delete
        
        Call OrganizarHojaHorasContador(HojaDestino)
        
    End If
    
    LibroDestino.Save
    LibroDestino.Close SaveChanges:=False

Salir:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "CopiarRangoYBuscarColor"
   
End Sub

Sub OrganizarHojaHorasContador(ws As Worksheet)
'
' Macro mejorada para organizar una hoja de trabajo específica.
' Recibe la hoja como un parámetro (ws) para saber dónde actuar.
'
    Dim lastRow As Long
    
    ' Comprobación de seguridad: si por alguna razón la hoja no es válida, salimos.
    If ws Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    ' La actualización de pantalla ya está desactivada por la macro principal,
    ' pero no hace daño dejarlo por si se usa de forma independiente.
    Application.ScreenUpdating = False

    ' Usamos un bloque "With" para trabajar directamente sobre la hoja (ws)
    ' que nos pasaron como parámetro.
    With ws
        ' Con el objeto Sort de la hoja, realizamos la ordenación.
        With .Sort
            .SortFields.Clear
            
            ' Añadimos el criterio de ordenación.
            .SortFields.Add2 Key:=ws.Range("A1:AF1"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal

            ' Configuramos las propiedades de la ordenación.
            ' FIX: Usar lastRow dinámico en lugar de rango fijo A1:AF100
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            ' Aseguramos que haya datos para ordenar
            If lastRow < 2 Then lastRow = 2
            
            .SetRange ws.Range("A1:AF" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlLeftToRight
            
            ' Aplicamos la ordenación.
            .Apply
        End With
        
        ' Autoajustamos el ancho de las columnas A hasta AF.
        .Columns("A:AF").AutoFit
    End With
    
    ' Nos desplazamos a la celda A1 en un solo paso.
    Application.GoTo ws.Range("A1"), Scroll:=True

ExitSub:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error en OrganizarHojaHorasContador: " & Err.Description, vbCritical
    Resume ExitSub
    
End Sub


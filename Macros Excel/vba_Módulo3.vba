Sub OrdenarPorApellido()

' OrdenarPorApellido Macro
'

    Application.EnableAnimations = False
    Range("A9:AV100").Select
    ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort.SortFields.Add2 Key:=Range( _
        "A9:A100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort.SortFields.Add2 Key:=Range( _
        "AL9:AL100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort
        .SetRange Range("A9:AV100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("SUELDO_ALQ_GASTOS").Select
    Range("B9:AB100").Select
    ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort.SortFields.Add2 Key:= _
        Range("K9:K100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort.SortFields.Add2 Key:= _
        Range("B9:B100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort
        .SetRange Range("B9:AB100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ENVIO CONTADOR").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("A9:BL100").Select
    ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort.SortFields.Add2 Key:=Range( _
        "C9:C100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort.SortFields.Add2 Key:=Range( _
        "B9:B100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort
        .SetRange Range("A9:BL100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("CALCULAR HORAS").Select
    ActiveWindow.SmallScroll Down:=-3
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 1
    Range("A9").Select
    Application.EnableAnimations = True
End Sub
Sub OrdenatPorLegajo()
'
' OrdenatPorLegajo Macro
'

    Application.EnableAnimations = False
    Range("A9:AV100").Select
    ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort.SortFields.Add2 Key:=Range( _
        "AL9:AL100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort.SortFields.Add2 Key:=Range( _
        "A9:A100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("CALCULAR HORAS").Sort
        .SetRange Range("A9:AV100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("SUELDO_ALQ_GASTOS").Select
    Range("B9:AB100").Select
    ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort.SortFields.Add2 Key:= _
        Range("B9:B100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort.SortFields.Add2 Key:= _
        Range("K9:K100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SUELDO_ALQ_GASTOS").Sort
        .SetRange Range("B9:AB100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("ENVIO CONTADOR").Select
    Range("A9:BL100").Select
    ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort.SortFields.Add2 Key:=Range( _
        "B9:B100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort.SortFields.Add2 Key:=Range( _
        "C9:C100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ENVIO CONTADOR").Sort
        .SetRange Range("A9:BL100")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("CALCULAR HORAS").Select
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("A9").Select
    Application.EnableAnimations = True
End Sub
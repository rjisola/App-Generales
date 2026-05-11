Sub Macro1()
'
' Macro1 Macro
'

'
    Range("AC9:AC95").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-90
    Range("AC9").Select
End Sub
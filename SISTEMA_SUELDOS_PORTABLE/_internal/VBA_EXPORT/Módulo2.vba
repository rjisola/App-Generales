Sub DETALLEQUINC()
If Range("B5").Value = "X" Then
Range("T6").Value = "1ERA " & UCase(Format(Date, "MMMM")) & " " & Year(Date)
Else
Range("B6").Value = "X"
Range("T6").Value = "2DA " & UCase(Format(Date, "MMMM")) & " " & Year(Date)
End If
End Sub
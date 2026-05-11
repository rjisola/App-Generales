Function Redondear(numero)

    Dim diferencia
    
    diferencia = numero - CLng(numero)

    If diferencia <> 0 Then
        If diferencia < 0 Then
            Redondear = CLng(numero)
        Else
            Redondear = CLng(numero) + 1
        End If
    Else
        Redondear = numero
    End If
    
End Function

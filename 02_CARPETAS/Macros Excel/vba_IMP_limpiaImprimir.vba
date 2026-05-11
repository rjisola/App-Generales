Sub limpiarImprimir(contador, par)

    If par = 0 Then
        
        Hoja3.Cells(contador, 2).Value = 0
        Hoja3.Cells(contador, 3).Value = 0
    
        Hoja3.Cells(contador + 1, 2).Value = 0
        Hoja3.Cells(contador + 1, 3).Value = 0
    
        Hoja3.Cells(contador + 2, 2).Value = 0
        Hoja3.Cells(contador + 2, 3).Value = 0
        
        If Hoja3.Cells(contador + 3, 1).Value <> "ADICIONAL" Then
            Hoja3.Cells(contador + 3, 2).Value = 0
            Hoja3.Cells(contador + 3, 3).Value = 0
        End If
        
        Hoja3.Cells(contador + 4, 2).Value = 0
        Hoja3.Cells(contador + 4, 3).Value = 0
        
        Hoja3.Cells(contador + 5, 2).Value = 0
        Hoja3.Cells(contador + 5, 3).Value = 0
        
        If Hoja3.Cells(contador + 6, 1).Value <> "AJUSTE" Then
            Hoja3.Cells(contador + 6, 2).Value = 0
            Hoja3.Cells(contador + 6, 3).Value = 0
        End If
        
        Hoja3.Cells(contador + 7, 2).Value = 0
        Hoja3.Cells(contador + 7, 3).Value = 0
        
        Hoja3.Cells(contador + 8, 2).Value = 0
        Hoja3.Cells(contador + 8, 3).Value = 0
        
        Hoja3.Cells(contador + 10, 2).Value = 0
        Hoja3.Cells(contador + 10, 3).Value = 0
        
        Hoja3.Cells(contador + 11, 2).Value = 0
        Hoja3.Cells(contador + 11, 3).Value = 0
    
    Else
    
        Hoja3.Cells(contador, 5).Value = 0
        Hoja3.Cells(contador, 6).Value = 0
    
        Hoja3.Cells(contador + 1, 5).Value = 0
        Hoja3.Cells(contador + 1, 6).Value = 0
    
        Hoja3.Cells(contador + 2, 5).Value = 0
        Hoja3.Cells(contador + 2, 6).Value = 0
    
        If Hoja3.Cells(contador + 3, 1).Value <> "ADICIONAL" Then
            Hoja3.Cells(contador + 3, 5).Value = 0
            Hoja3.Cells(contador + 3, 6).Value = 0
        End If
        
        Hoja3.Cells(contador + 4, 5).Value = 0
        Hoja3.Cells(contador + 4, 6).Value = 0
    
        Hoja3.Cells(contador + 5, 5).Value = 0
        Hoja3.Cells(contador + 5, 6).Value = 0
    
        If Hoja3.Cells(contador + 6, 4).Value <> "AJUSTE" Then
            Hoja3.Cells(contador + 6, 5).Value = 0
            Hoja3.Cells(contador + 6, 6).Value = 0
        End If
    
        Hoja3.Cells(contador + 7, 5).Value = 0
        Hoja3.Cells(contador + 7, 6).Value = 0
    
        Hoja3.Cells(contador + 8, 5).Value = 0
        Hoja3.Cells(contador + 8, 6).Value = 0
    
        Hoja3.Cells(contador + 10, 5).Value = 0
        Hoja3.Cells(contador + 10, 6).Value = 0
    
        Hoja3.Cells(contador + 11, 5).Value = 0
        Hoja3.Cells(contador + 11, 6).Value = 0
    
    End If

    

End Sub

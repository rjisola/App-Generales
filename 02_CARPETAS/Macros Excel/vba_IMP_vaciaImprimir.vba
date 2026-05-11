Sub vaciarImprimir()

    'Hasta el maximo de gente multiplicado por el desplazamiento

    ActiveSheet.Columns("A:F").ClearContents
    ActiveSheet.Columns("A:F").Interior.color = RGB(255, 255, 255)

End Sub

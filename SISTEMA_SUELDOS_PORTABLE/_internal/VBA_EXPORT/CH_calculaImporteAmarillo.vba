' ============================================================================
' MÓDULO: CH_calculaImporteAmarillo
' DESCRIPCIÓN: Calcula el importe total para empleados de categoría AMARILLO
'              aplicando multiplicadores según el sub-proyecto (Quilmes/Papelera)
' ============================================================================

' --- CONSTANTES DE MULTIPLICADORES ---
' Estos valores se aplican a las horas extras según el sub-proyecto
Const MULTIPLICADOR_QUILMES As Double = 1.2        ' 20% adicional para Quilmes
Const MULTIPLICADOR_PAPELERA As Double = 1.344     ' 34.4% adicional para Papelera (1.2 * 1.12)

' --- CONSTANTES DE COLORES RGB ---
' Colores usados para identificar sub-proyectos en las celdas
Const COLOR_QUILMES_R As Integer = 255
Const COLOR_QUILMES_G As Integer = 192
Const COLOR_QUILMES_B As Integer = 0

Const COLOR_PAPELERA_R As Integer = 112
Const COLOR_PAPELERA_G As Integer = 173
Const COLOR_PAPELERA_B As Integer = 71

' --- CONSTANTES DE COLUMNAS ---
Const COL_HORAS_50 As Integer = 21
Const COL_HORAS_100 As Integer = 22
Const COL_HORAS_FERIADO As Integer = 23
Const COL_IMPORTE_FERIADO As Integer = 25
Const COL_IMPORTE_NORMAL As Integer = 26
Const COL_IMPORTE_50 As Integer = 27
Const COL_IMPORTE_100 As Integer = 28
Const COL_TOTAL As Integer = 29
Const COL_TOTAL_DUPLICADO As Integer = 30

Sub calcularImporteAmarillo(fila, ByRef presentismo, categoria, horasQuilmesCincuenta, horasPapeleraCincuenta, horasQuilmesCien, horasPapeleraCien)
    ' ========================================================================
    ' PROPÓSITO: Calcula importes para empleados AMARILLO con multiplicadores
    '            diferenciados por sub-proyecto (Quilmes/Papelera)
    '
    ' PARÁMETROS:
    '   fila                    - Número de fila en la hoja de cálculo
    '   presentismo             - Indica si el empleado tiene presentismo
    '   categoria               - Categoría del empleado (ESPECIALIZADO, OFICIAL, etc.)
    '   horasQuilmesCincuenta   - Horas al 50% en proyecto Quilmes
    '   horasPapeleraCincuenta  - Horas al 50% en proyecto Papelera
    '   horasQuilmesCien        - Horas al 100% en proyecto Quilmes
    '   horasPapeleraCien       - Horas al 100% en proyecto Papelera
    '
    ' LÓGICA DE CÁLCULO:
    '   - Horas blancas (sin sub-proyecto): tarifa normal
    '   - Horas Quilmes: tarifa * 1.2
    '   - Horas Papelera: tarifa * 1.344 (1.2 * 1.12)
    ' ========================================================================
    
    Dim valorHoraNormal As Double
    Dim valorHoraAlCincuenta As Double
    Dim valorHoraAlCien As Double
    Dim valorHoraFeriado As Double
    Dim importeHoraNormal As Double
    Dim importeHoraAlCincuenta As Double
    Dim importeHoraAlCien As Double
    Dim importeHorasQuilmesCien As Double
    Dim importeHorasPapeleraCien As Double
    Dim importeHorasQuilmesCincuenta As Double
    Dim importeHorasPapeleraCincuenta As Double
    Dim importeHorasAlCincuentaBlancas As Double
    Dim importeHorasAlCienBlancas As Double
    Dim total As Double
    
    ' Inicializar valores
    valorHoraNormal = 0
    valorHoraAlCincuenta = 0
    valorHoraAlCien = 0
    valorHoraFeriado = 0

    ' --- DETERMINAR VALOR HORA NORMAL SEGÚN CATEGORÍA ---
    If categoria <> vbNullString Then
        ' Marcar celda como válida (color celeste)
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(189, 215, 238)
    
        ' Obtener tarifa base según categoría del empleado
        If categoria = "ESPECIALIZADO" Or categoria = "MAQUINISTA" Then
            valorHoraNormal = ActiveSheet.Range("B1").Value
        ElseIf categoria = "OFICIAL" Then
            valorHoraNormal = ActiveSheet.Range("B2").Value
        ElseIf categoria = "MEDIO OFICIAL" Then
            valorHoraNormal = ActiveSheet.Range("B3").Value
        ElseIf categoria = "AYUDANTE" Then
            valorHoraNormal = ActiveSheet.Range("B4").Value
        End If
    Else
        ' Marcar celda como error (color rojo) si no hay categoría
        ActiveSheet.Cells(fila, 2).Interior.color = RGB(255, 0, 0)
    End If

    ' --- CALCULAR TARIFAS PARA HORAS EXTRAS ---
    valorHoraAlCincuenta = valorHoraNormal * 1.5  ' 50% adicional
    valorHoraAlCien = valorHoraNormal * 2         ' 100% adicional
    valorHoraFeriado = valorHoraAlCien            ' Feriados = 100%
    
    ' --- CALCULAR IMPORTES (PRESENTISMO NO AFECTA EL CÁLCULO) ---
    
    ' Calcular horas blancas (sin sub-proyecto) al 50%
    horasBlancasCincuenta = ActiveSheet.Cells(fila, COL_HORAS_50).Value - horasPapeleraCincuenta - horasQuilmesCincuenta
    
    ' Importes al 50%
    importeHorasAlCincuentaBlancas = horasBlancasCincuenta * valorHoraAlCincuenta
    importeHorasQuilmesCincuenta = horasQuilmesCincuenta * valorHoraAlCincuenta * MULTIPLICADOR_QUILMES
    importeHorasPapeleraCincuenta = horasPapeleraCincuenta * valorHoraAlCincuenta * MULTIPLICADOR_PAPELERA
    importeHoraAlCincuenta = importeHorasQuilmesCincuenta + importeHorasPapeleraCincuenta + importeHorasAlCincuentaBlancas
    
    ' Calcular horas blancas (sin sub-proyecto) al 100%
    horasBlancasCien = ActiveSheet.Cells(fila, COL_HORAS_100).Value - horasPapeleraCien - horasQuilmesCien
    
    ' Importes al 100%
    importeHorasAlCienBlancas = horasBlancasCien * valorHoraAlCien
    importeHorasQuilmesCien = horasQuilmesCien * valorHoraAlCien * MULTIPLICADOR_QUILMES
    importeHorasPapeleraCien = horasPapeleraCien * valorHoraAlCien * MULTIPLICADOR_PAPELERA
    importeHoraAlCien = importeHorasQuilmesCien + importeHorasPapeleraCien + importeHorasAlCienBlancas
   
    ' Importes de feriados
    importeHoraFeriado = ActiveSheet.Cells(fila, COL_HORAS_FERIADO).Value * valorHoraFeriado
    
    ' --- ESCRIBIR RESULTADOS EN HOJA ---
    ActiveSheet.Cells(fila, COL_IMPORTE_FERIADO).Value = importeHoraFeriado
    ActiveSheet.Cells(fila, COL_IMPORTE_NORMAL).Value = importeHoraNormal
    ActiveSheet.Cells(fila, COL_IMPORTE_50).Value = importeHoraAlCincuenta
    ActiveSheet.Cells(fila, COL_IMPORTE_100).Value = importeHoraAlCien

    ' Calcular y escribir total
    total = importeHoraAlCincuenta + importeHoraAlCien + importeHoraFeriado
    ActiveSheet.Cells(fila, COL_TOTAL).Value = total
    ActiveSheet.Cells(fila, COL_TOTAL_DUPLICADO).Value = total

End Sub

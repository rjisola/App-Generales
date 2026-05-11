' ============================================================================
' MÓDULO: CH_generaHorasGris
' DESCRIPCIÓN: Genera y distribuye horas para empleados de categoría GRIS
'              aplicando límites según el día de la semana
' ============================================================================

' --- CONSTANTES DE LÍMITES DE HORAS ---
Const LIMITE_HORAS_NORMALES_LUN_JUE As Integer = 9   ' Lunes a Jueves
Const LIMITE_HORAS_NORMALES_VIERNES As Integer = 8   ' Viernes
Const LIMITE_HORAS_50_SABADO As Integer = 4          ' Sábado

' --- CONSTANTES DE COLUMNAS ---
Const COL_HORAS_NORMALES As Integer = 20
Const COL_HORAS_50 As Integer = 21
Const COL_HORAS_100 As Integer = 22
Const COL_HORAS_FERIADO As Integer = 23
Const COL_PRESENTISMO As Integer = 24

' --- CONSTANTES DE CÓDIGOS ESPECIALES ---
Const CODIGO_AUSENCIA As Integer = -1        ' Ausencia sin justificación
Const CODIGO_AUSENCIA_CERT As Integer = -8   ' Ausencia con certificado

Sub generarHorasGris(fila, columna, Dia, ByRef presentismo, feriado, ByRef horas)
    ' ========================================================================
    ' PROPÓSITO: Distribuye las horas trabajadas en categorías según el día
    '            y aplica límites para horas normales, 50% y 100%
    '
    ' PARÁMETROS:
    '   fila        - Número de fila en la hoja de cálculo
    '   columna     - Número de columna del día
    '   Dia         - Nombre del día (lunes, martes, etc.)
    '   presentismo - Indica si el empleado mantiene presentismo
    '   feriado     - Indica si el día es feriado
    '   horas       - Cantidad de horas trabajadas
    '
    ' REGLAS GRIS:
    '   - Lunes a Jueves: 9 hs normales, 9<12 al 50%, >12 al 100%
    '   - Viernes: 8 hs normales, 8<12 al 50%, >12 al 100%
    '   - Sábado: 4 hs al 50%, >4 al 100%
    '   - Domingo: No trabaja (si trabaja, todo al 100%)
    '   - Feriado: Todo al 100%
    '   - PRESENTISMO: Si falta lo pierde, con certificado cobra y pierde pres.
    ' ========================================================================

    Dim horasAlCien As Single
    Dim horasAlCincuenta As Single
    Dim horasNormales As Single
    Dim horasFeriado As Single
    Dim vac As Single

    ' Inicializar valores
    horasAlCien = 0
    horasAlCincuenta = 0
    horasNormales = 0
    horasFeriado = 0
    
    ' --- PROCESAMIENTO DE FERIADOS ---
    If feriado Then
        If horas <= CODIGO_AUSENCIA Or horas > 24 Then
            If horas = CODIGO_AUSENCIA Then
                ' Ausencia en feriado: asignar horas normales según el día
                If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Then
                    horasNormales = LIMITE_HORAS_NORMALES_LUN_JUE
                ElseIf Dia = "viernes" Then
                    horasNormales = LIMITE_HORAS_NORMALES_VIERNES
                ElseIf Dia = "sábado" Then
                    horasAlCincuenta = LIMITE_HORAS_50_SABADO
                End If
            Else
                ' Código de horas inválido
                Call informarError
            End If
        Else
            ' Horas trabajadas en feriado: todas al 100%
            horasFeriado = horas
        End If
        
    ' --- PROCESAMIENTO DE DÍAS NORMALES ---
    Else
        ' --- LUNES A JUEVES ---
        If Dia = "lunes" Or Dia = "martes" Or Dia = "miércoles" Or Dia = "jueves" Then
            ' Procesar códigos especiales
            If horas < 0 Or horas > 24 Then
                If horas = CODIGO_AUSENCIA Then
                    presentismo = False  ' Pierde presentismo
                ElseIf horas = CODIGO_AUSENCIA_CERT Then
                    horasNormales = LIMITE_HORAS_NORMALES_LUN_JUE
                    presentismo = False  ' Cobra pero pierde presentismo
                Else
                    Call informarError
                End If
            End If
        
            ' Distribuir horas normales (0-9 horas)
            If horas <= LIMITE_HORAS_NORMALES_LUN_JUE And horas > 0 Then
                horasNormales = horas
            End If
        
            ' Distribuir horas extras (>9 horas)
            If horas > LIMITE_HORAS_NORMALES_LUN_JUE Then
                horasNormales = LIMITE_HORAS_NORMALES_LUN_JUE
                horasAlCincuenta = horas - horasNormales  ' Excedente al 50%
            End If
        End If
    
        ' --- VIERNES ---
        If Dia = "viernes" Then
            ' Procesar códigos especiales
            If horas < 0 Or horas > 24 Then
                If horas = CODIGO_AUSENCIA Then
                    presentismo = False
                ElseIf horas = CODIGO_AUSENCIA_CERT Then
                    horasNormales = LIMITE_HORAS_NORMALES_VIERNES
                    presentismo = False
                Else
                    Call informarError
                End If
            End If
        
            ' Distribuir horas normales (0-8 horas)
            If horas <= LIMITE_HORAS_NORMALES_VIERNES And horas > 0 Then
                horasNormales = horas
            End If
        
            ' Distribuir horas extras (>8 horas)
            If horas > LIMITE_HORAS_NORMALES_VIERNES Then
                horasNormales = LIMITE_HORAS_NORMALES_VIERNES
                horasAlCincuenta = horas - horasNormales  ' Excedente al 50%
            End If
        End If
    
        ' --- SÁBADO ---
        If Dia = "sábado" Then
            ' Procesar códigos especiales
            If horas < 0 Or horas > 24 Then
                If horas = CODIGO_AUSENCIA Then
                    ' No hace nada, ya está en 0
                ElseIf horas = CODIGO_AUSENCIA_CERT Then
                    horasNormales = LIMITE_HORAS_NORMALES_VIERNES
                    presentismo = False
                Else
                    Call informarError
                End If
            End If
        
            ' Distribuir horas al 50% (0-4 horas)
            If horas <= LIMITE_HORAS_50_SABADO And horas > 0 Then
                horasAlCincuenta = horas
            End If
        
            ' Distribuir horas mixtas (>4 horas)
            If horas > LIMITE_HORAS_50_SABADO Then
                horasAlCincuenta = LIMITE_HORAS_50_SABADO  ' Primeras 4 al 50%
                horasAlCien = horas - horasAlCincuenta     ' Excedente al 100%
            End If
        End If
    End If
    
    ' --- DOMINGO Y FERIADO ---
    If Dia = "domingo" Or Dia = "feriado" Then
        horasAlCien = horas  ' Todo al 100%
    End If
    
    If Dia = "feriado" Then
        horasFeriado = horas
    End If
    
    ' --- ACUMULAR HORAS EN LA HOJA ---
    ActiveSheet.Cells(fila, COL_HORAS_NORMALES).Value = ActiveSheet.Cells(fila, COL_HORAS_NORMALES).Value + horasNormales
    ActiveSheet.Cells(fila, COL_HORAS_50).Value = ActiveSheet.Cells(fila, COL_HORAS_50).Value + horasAlCincuenta
    ActiveSheet.Cells(fila, COL_HORAS_100).Value = ActiveSheet.Cells(fila, COL_HORAS_100).Value + horasAlCien
    ActiveSheet.Cells(fila, COL_HORAS_FERIADO).Value = ActiveSheet.Cells(fila, COL_HORAS_FERIADO).Value + horasFeriado
        
    ' --- MARCAR ESTADO DE PRESENTISMO ---
    ' Nota: GRIS no muestra estado de presentismo (siempre muestra "-")
    If presentismo Then
        ActiveSheet.Cells(fila, COL_PRESENTISMO).Value = "-"
    Else
        ActiveSheet.Cells(fila, COL_PRESENTISMO).Value = "-"
    End If
End Sub

import pandas as pd
import math
import os
import sys

# Añadir el directorio actual al path para importar los submódulos
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, sys.path[0]) # Mantener prioridad

from payroll_azul    import calculate_azul
from payroll_celeste import calculate_celeste
from payroll_gris    import calculate_gris
from payroll_amarillo import calculate_amarillo
from payroll_blanco  import calculate_blanco

def _unify_day_input(day_text, config):
    if pd.isna(day_text) or str(day_text).strip() == '': return {"hours": 0.0, "status": "OK"}
    text = str(day_text).lower().strip()
    for mapping in config['text_input_mappings']['mappings']:
        if text in mapping['inputs']: return {"hours": float(mapping['output']['hours']), "status": mapping['output']['status']}
    try:
        num = float(text)
        return {"hours": num if 0 <= num <= 24 else 0.0, "status": "OK"}
    except (ValueError, TypeError): return {"hours": 0.0, "status": "UNKNOWN_TEXT"}

def unify_name(n):
    if n is None or pd.isna(n): return ""
    import unicodedata
    s = str(n).strip().upper()
    s = "".join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return " ".join(s.split())

def process_payroll_for_employee(employee_data, config, day_definitions, rate_config,
                                  wb_styles=None, row_idx=None):
    """
    Orquestador principal de nómina.
    Regla de Oro: Se prioriza el ACUERDO (Col L) para todos los colores.

    Parámetros adicionales:
        wb_styles  – workbook openpyxl cargado con data_only=False (para leer colores en AMARILLO).
        row_idx    – número de fila Excel del empleado actual (1-based).
    """
    color_cat      = str(employee_data.get('Cat_Color_Name', 'CELESTE')).upper().strip()
    sueldo_acuerdo = float(employee_data.get('Sueldo_Acordado', 0.0))
    sueldo_contador = float(employee_data.get('Sueldo_Sobre', 0.0))
    
    # Pre-cargar valores UOCRA del archivo de configuración (INICIO/Tarifas)
    control_rates = rate_config.get("control_rates", {})
    uocra_50     = float(control_rates.get("C1", 0.0))
    uocra_100    = float(control_rates.get("D1", 0.0))
    price_normal = float(control_rates.get("B1", 0.0))
    
    # Determinar Base de Cálculo
    base_de_calculo = sueldo_acuerdo
    if base_de_calculo <= 0:
        base_de_calculo = sueldo_contador
        
    # Determinar tarifas UOCRA basadas en el PUESTO detectado
    puesto    = str(employee_data.get('PUESTO_UOCRA', 'AYUDANTE')).upper().strip()
    job_rates = rate_config.get('job_title_rates', {}).get(puesto, {})
    
    # Priorizar tarifas del puesto. Si no existen, usar el control_rates general.
    uocra_50_final  = job_rates.get('rate_50_value',  uocra_50)
    uocra_100_final = job_rates.get('rate_100_value', uocra_100)
    
    # Si sigue siendo 0, es un empleado UOCRA puro sin acuerdos especiales
    if base_de_calculo <= 0:
        base_de_calculo = price_normal * 88  # 88 horas quincenales base
        res = calculate_gris(employee_data, base_de_calculo, day_definitions, config,
                             _unify_day_input, uocra_50_final, uocra_100_final)
    else:
        # ── DELEGACIÓN POR COLOR ──────────────────────────────────────────────
        if color_cat == 'AZUL':
            res = calculate_azul(employee_data, base_de_calculo, day_definitions, config,
                                 _unify_day_input)

        elif color_cat == 'AMARILLO':
            # AMARILLO recibe wb_styles + row_idx para detectar colores por celda
            # y aplicar multiplicadores Quilmes (×1.2) y Papelera (×1.344)
            res = calculate_amarillo(employee_data, base_de_calculo, day_definitions, config,
                                     _unify_day_input, uocra_50_final, uocra_100_final,
                                     wb_styles=wb_styles, row_idx=row_idx)

        elif color_cat == 'GRIS':
            res = calculate_gris(employee_data, base_de_calculo, day_definitions, config,
                                 _unify_day_input, uocra_50_final, uocra_100_final)

        elif color_cat == 'BLANCO':
            # BLANCO: misma lógica que GRIS pero con tasa UOCRA × 1.2
            res = calculate_blanco(employee_data, base_de_calculo, day_definitions, config,
                                   _unify_day_input, uocra_50_final, uocra_100_final)

        else:
            # CELESTE o fallback
            res = calculate_celeste(employee_data, base_de_calculo, day_definitions, config,
                                    _unify_day_input)
            
    sueldo_base, extras_without_feriado, v_50, v_100, h_50, h_100 = res
    
    # Consolidar Ajustes según Spreadsheet
    premio         = float(employee_data.get('Premio',          0.0))  # Col W en SUELDO_ALQ_GASTOS
    reintegro      = float(employee_data.get('Reintegro',       0.0))  # Col N
    ajuste_alq     = float(employee_data.get('Ajuste_Alquiler', 0.0))  # Col O
    adelanto       = float(employee_data.get('Adelanto',        0.0))  # Col M
    obra_social    = float(employee_data.get('Obra_Social',     0.0))  # Col Q
    gasto_personal = float(employee_data.get('Gasto_Personal', 0.0))  # Col P
    patente        = float(employee_data.get('Patente',        0.0))  # Col 18 de Hoja4 (VBA detectado)
    sueldo_sobre   = float(employee_data.get('Sueldo_Sobre',    0.0))  # Col J (índice 9) en SUELDO_ALQ_GASTOS
    metodo_pago    = str(employee_data.get('Metodo_Pago', '')).strip().upper()

    # Calcular Horas Normales y Feriados
    total_hours_worked = 0.0
    feriado_hours = 0.0
    presentismo = "SI"

    for day in day_definitions:
        col_key = day['col_key_in_df']
        val = employee_data.get(col_key, "")
        
        # Chequeo de presentismo (igual que en logic_accountant)
        val_str = str(val).lower().strip()
        if val_str in ['falto', 'enfermo', 'certif', 'cert', 'certificado']:
            presentismo = "NO"
            
        parsed = _unify_day_input(val, config)
        h = parsed['hours']
        total_hours_worked += h
        if day['is_holiday']:
            feriado_hours += h
            
    # El importe de feriado se paga al 100% de la tarifa horaria
    importe_feriado = round(feriado_hours * v_100)
    
    # El TOTAL de extras (para el recibo) incluye 50%, 100% Y Feriados
    extras_totales = extras_without_feriado + importe_feriado

    # Extraer Importe Altura si existe (Cols 31/32)
    importe_altura = 0.0
    for key, val in employee_data.items():
        if 'ALTURA' in str(key).upper():
            try:
                # Usar pd.isna para evitar que NaN se convierta en float y rompa el round()
                if pd.notna(val):
                    importe_altura = float(val)
                    if math.isnan(importe_altura):
                        importe_altura = 0.0
                break
            except: pass

    horas_normales = total_hours_worked - h_50 - h_100 - feriado_hours
    if horas_normales < 0: horas_normales = 0.0

    # Fórmula Final de Paridad (Neto)
    # 1. Total Conceptos Positivos
    total_bruto = sueldo_base + extras_totales + premio + reintegro + ajuste_alq + importe_altura
    
    # 2. Total Descuentos
    total_descuentos = adelanto + obra_social + gasto_personal + patente
    
    # 3. Neto Final
    neto_final = round(total_bruto - total_descuentos)

    # LÓGICA DE DISTRIBUCIÓN DE FONDOS (VBA: Regla de Banco vs Caja de Ahorro)
    # El monto de 'Sueldo Sobre' es lo que se pretende pagar por Banco.
    # El resto del neto va a Caja de Ahorro 2 o Efectivo.
    monto_banco = sueldo_sobre
    monto_ca2   = neto_final - monto_banco
    
    # REGLA DE ABSORCIÓN: Si CA2 es negativo (descuentos superan el excedente), 
    # se descuenta del Banco y CA2 queda en 0.
    if monto_ca2 < 0:
        monto_banco += monto_ca2 # Sumar un negativo es restar
        monto_ca2 = 0.0
        
    # PARIDAD MACRO: 
    # 1. En CALCULAR HORAS el total solo incluye Sueldo + Extras (sin reintegros ni ajustes)
    total_calculo_horas = round(sueldo_base + extras_totales)
    
    # 2. En RECUENTO TOTAL el ajuste es (Neto Final - Sueldo Sobre)
    # Nota: Este ajuste es informativo para la planilla de control
    ajuste_recuento = neto_final - sueldo_sobre

    # Legajo: primer columna de la fila (numérica), guardada en Excel_Row_Index
    # El legajo real lo inyecta el caller (B-PROCESARSUELDOS) si lo tiene
    legajo_real = employee_data.get('Legajo_Real', employee_data.get('Excel_Row_Index'))

    # Fondo desempleo para recibo
    fondo_desempleo = 0.0
    if color_cat == 'BLANCO':
        fondo_desempleo = (h_50 * v_50 + h_100 * v_100) * 0.12

    return {
        'Legajo':           employee_data.get('Legajo'),
        'Cuenta1':          employee_data.get('Cuenta1'),
        'Cuenta2':          employee_data.get('Cuenta2'),
        'CUIL':             employee_data.get('CUIL'),
        'Banco':            employee_data.get('Banco'),
        'Nombre':           employee_data.get('NOMBRE Y APELLIDO', ''),
        'Categoría':        color_cat,
        'Puesto':           puesto,
        'Total Quincena':   total_bruto, # Volver al bruto para paridad con Hoja 10
        'Total_CH':         total_calculo_horas,
        'Sueldo Sobre':     sueldo_sobre,
        'Sueldo_Acordado':  sueldo_acuerdo,
        'Presentismo':      presentismo,
        'Ajuste_Recuento':  ajuste_recuento,
        'Total Extras':     extras_totales,
        'Sueldo Base':      sueldo_base,
        'Monto_Banco':      monto_banco,
        'Neto_Real':        neto_final, # Mantener el neto real para los recibos
        'Adelanto':         adelanto,
        'Reintegro':        reintegro,
        'Ajuste Alquiler':  ajuste_alq,
        'Gasto Personal':   gasto_personal,
        'Obra Social':      obra_social,
        'Patente':          patente,
        'Premio':           premio,
        'V.HORA 50%':       v_50,
        'V.HORA 100%':      v_100,
        'Horas Normales':   horas_normales,
        'Horas Feriados':   feriado_hours,
        'Horas al 50%':     h_50,
        'Horas al 100%':    h_100,
        'Fondo_Desempleo':  fondo_desempleo,
        'Importe_Altura':   importe_altura,
        'Metodo_Pago':      metodo_pago,
        'Caja_Ahorro_2':    monto_ca2,
        'Legajo_Real':      legajo_real,
        'Excel_Row_Index':  employee_data.get('Excel_Row_Index')
    }

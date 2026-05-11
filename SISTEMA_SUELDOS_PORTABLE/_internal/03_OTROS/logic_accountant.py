import pandas as pd

def process_accountant_summary_for_employee(employee_data, config, day_definitions, rate_config, wb_styles=None, row_idx=None, sheet_name='CALCULAR HORAS'):
    """
    Procesa la lógica del resumen para el contador para un solo empleado.
    Replica la funcionalidad de EC_generaHorasContador y EC_copiaEnContador de VBA.
    
    CRITICO: Utiliza wb_styles (cargado con data_only=False) para leer el color de CADA celda de día.
    
    Categorías por color (VBA):
    - BLANCO (RGB 255,255,255): Horas normales (Columna 6 en Contador)
    - NARANJA (RGB 255,192,0): Quilmes (Columna 7 en Contador)
    - VERDE (RGB 112,173,71) o MARRON (RGB 153,102,0): Papelera (Columna 8 en Contador)
    """
    from data_loader import get_rgb_from_openpyxl_color
    
    # Inicializar contadores por categoría de proyecto
    # En la hoja ENVIO CONTADOR:
    # Col 6 (F): Normal (Blanco)
    # Col 7 (G): Quilmes (Naranja)
    # Col 8 (H): Papelera (Verde/Marrón)
    horas_blanco = 0.0
    horas_quilmes = 0.0
    horas_papelera = 0.0
    
    # Contadores separados para ART (columnas M, N, O)
    horas_blanco_art = 0.0
    horas_quilmes_art = 0.0
    horas_papelera_art = 0.0
    
    # Contadores separados para Enfermo (columnas P, Q, R)
    horas_blanco_enfermo = 0.0
    horas_quilmes_enfermo = 0.0
    horas_papelera_enfermo = 0.0
    
    perdio_presentismo = False
    
    # Obtener hoja para leer colores si está disponible
    ws_styles = None
    if wb_styles and sheet_name in wb_styles.sheetnames:
        ws_styles = wb_styles[sheet_name]
    
    # Determinar precio de categoría
    categoria_empleado = str(employee_data.get('PUESTO_UOCRA', '')).strip()
    
    # IMPORTANTE: Empleados sin categoría o categorías administrativas/capataces 
    # tienen una lógica especial en el contador.
    is_admin_or_capataz = any(x in categoria_empleado.upper() for x in ['ADMINISTRACION', 'CAPATAZ', 'SOCIOS', 'ANALISTA'])
    
    if not categoria_empleado or categoria_empleado == '':
        categoria_empleado = 'ESPECIALIZADO'
    
    job_title_rates = rate_config.get("job_title_rates", {}) 
    
    # Buscar la info de la categoría
    cat_info = job_title_rates.get(categoria_empleado.upper(), {})
    
    # Lógica de Paridad con VBA (Efecto memoria/arrastre):
    if not cat_info:
        # Casos que en el original terminan con precio de ESPECIALIZADO (6011)
        if any(x in categoria_empleado.upper() for x in ['ADMINISTRACION', 'ANALISTA', 'SOCIO']):
            # Excepción: ADMINISTRACION2 está al final cerca de AYUDANTES
            if 'ADMINISTRACION2' in categoria_empleado.upper():
                cat_info = job_title_rates.get('AYUDANTE', {})
            else:
                cat_info = job_title_rates.get('ESPECIALIZADO', {})
        # Casos que en el original terminan con precio de AYUDANTE (4374)
        elif 'CAPATAZ' in categoria_empleado.upper():
            cat_info = job_title_rates.get('AYUDANTE', {})
        else:
            cat_info = job_title_rates.get('ESPECIALIZADO', {})
    
    precio_categoria = cat_info.get('base_rate_cell_value', 0.0)





    # Definir colores RGB de referencia (con tolerancia)
    RGB_BLANCO = (255, 255, 255)
    RGB_NARANJA = (255, 192, 0)
    RGB_NARANJA_CLARO = (255, 204, 102) # Nuevo color detectado en Quilmes (FFCC66)
    RGB_VERDE = (112, 173, 71)
    RGB_MARRON = (153, 102, 0)
    
    def matches_color(rgb, target_rgb, tolerance=5):
        if not rgb: return False
        return (abs(rgb[0] - target_rgb[0]) <= tolerance and
                abs(rgb[1] - target_rgb[1]) <= tolerance and
                abs(rgb[2] - target_rgb[2]) <= tolerance)

    # DEBUG: Mansilla
    is_debug_target = "MANSILLA" in str(employee_data.get('NOMBRE Y APELLIDO', '')).upper()
    if is_debug_target:
        print(f"DEBUG MANSILLA: Fila Excel {row_idx}")

    # Procesar cada día
    for day_info in day_definitions:
        day_input_raw = employee_data.get(day_info['col_key_in_df'])
        day_name = day_info['day_name']
        is_holiday = day_info['is_holiday']
        
        # Saltar domingos y sábados para el contador (no se cuentan)
        if day_name in ['domingo', 'sábado']:
            continue
        
        # Obtener columna en Excel para leer el color (1-based index)
        col_idx = day_info.get('col_idx')
        
        day_color_type = 'BLANCO' # Default es BLANCO
        
        if ws_styles and row_idx and col_idx:
            cell = ws_styles.cell(row=row_idx, column=col_idx)
            if cell.fill and cell.fill.fgColor:
                color = cell.fill.fgColor
                
                # Caso 1: Colores de Tema
                if color.type == 'theme':
                    if color.theme == 7:
                        day_color_type = 'NARANJA'
                    elif color.theme == 8:
                        day_color_type = 'PAPELERA'
                
                # Caso 2: Colores RGB
                else:
                    rgb = get_rgb_from_openpyxl_color(color)
                    if rgb:
                        if matches_color(rgb, RGB_NARANJA) or matches_color(rgb, RGB_NARANJA_CLARO):
                            day_color_type = 'NARANJA'
                        elif matches_color(rgb, RGB_VERDE) or matches_color(rgb, RGB_MARRON):
                            day_color_type = 'PAPELERA'
        
        # --- Lógica de acumulación (VBA generarHorasContador) ---
        
        # CRÍTICO: VBA solo procesa 'art', 'enfermo', 'lluvia', etc. cuando fila 7 ESTÁ VACÍA
        # If Not IsEmpty(ws2.Cells(7, columna)) Then ... Else [procesa casos especiales]
        # Por lo tanto, si is_holiday=True (fila 7 NO vacía), NO procesamos estos casos
        if is_holiday:
            # Día marcado como feriado (fila 7 no vacía) - VBA no procesa casos especiales
            continue
        
        day_input_str = str(day_input_raw).lower().strip() if not pd.isna(day_input_raw) else ''
        
        # Determinar límite de horas para este día
        limit_horas = 8 if day_name == 'viernes' else 9
        
        # Procesar textos
        if day_input_str in ['vacaciones', 'cortaron', 'c/aviso', 'c/a']:
            continue
        elif day_input_str == 'falto':
            perdio_presentismo = True
            continue
        elif day_input_str in ['enfermo', 'certif', 'cert', 'certificado']:
            perdio_presentismo = True
            # En VBA: enfermo suma a columnaBase + 10 (columnas P, Q, R)
            if day_name not in ['sábado', 'domingo']:
                # Sumar a contadores de ENFERMO (no normales)
                if day_color_type == 'BLANCO': horas_blanco_enfermo += limit_horas
                elif day_color_type == 'NARANJA': horas_quilmes_enfermo += limit_horas
                elif day_color_type == 'PAPELERA': horas_papelera_enfermo += limit_horas
            continue
        elif day_input_str == 'art':
            # En VBA: ART suma a columnaBase + 7 (columnas M, N, O)
            if day_name not in ['sábado', 'domingo']:
                # Sumar a contadores de ART (no normales)
                if day_color_type == 'BLANCO': horas_blanco_art += limit_horas
                elif day_color_type == 'NARANJA': horas_quilmes_art += limit_horas
                elif day_color_type == 'PAPELERA': horas_papelera_art += limit_horas
            continue
        elif day_input_str == 'lluvia':
            # VBA: Lluvia suma 2.5
             # CORREGIDO: Usar day_color_type
             if day_color_type == 'BLANCO': horas_blanco += 2.5
             elif day_color_type == 'NARANJA': horas_quilmes += 2.5
             elif day_color_type == 'PAPELERA': horas_papelera += 2.5
             continue
             
        # Procesar numéricos
        try:
            val = float(day_input_raw)
            if val > 0 and day_name not in ['sábado', 'domingo']:
                horas_a_sumar = min(val, limit_horas)
                
                # CORREGIDO: Usar day_color_type
                if not is_admin_or_capataz:
                    if day_color_type == 'BLANCO': horas_blanco += horas_a_sumar
                    elif day_color_type == 'NARANJA': 
                        horas_quilmes += horas_a_sumar
                    elif day_color_type == 'PAPELERA': horas_papelera += horas_a_sumar
                
        except (ValueError, TypeError):
            pass

    # Calcular Total Horas Contador (basado en VBA línea 72-74)
    # VBA: If Not IsEmpty(ws2.Cells(7, columna)) Then
    #      ws9.Cells(fila, 27).Value = ws9.Cells(fila, 27).Value + IIf(Dia = "viernes"..., 8, 9)
    # 
    # IMPORTANTE: El VBA cuenta días donde la fila 7 NO está vacía
    # En data_loader.py, is_holiday=True significa que la fila 7 tiene contenido
    # Por lo tanto, debemos contar días donde is_holiday=True (no False como antes)
    total_horas_contador = 0
    for day_def in day_definitions:
        if day_def['is_holiday']:  # Contar días donde fila 7 NO está vacía
            day_name = day_def['day_name'].lower()
            # Viernes/Sábado/Domingo: 8 horas
            # Lunes/Martes/Miércoles/Jueves: 9 horas
            if day_name in ['viernes', 'sábado', 'domingo']:
                total_horas_contador += 8
            else:
                total_horas_contador += 9
    
    return {
        'Nombre': employee_data.get('NOMBRE Y APELLIDO', ''),
        'Categoría': categoria_empleado,
        'Precio Categoría': precio_categoria,
        'Perdió Presentismo': perdio_presentismo,
        'Horas Blanco': horas_blanco,
        'Horas Quilmes': horas_quilmes,
        'Horas Papelera': horas_papelera,
        # Horas ART (columnas M, N, O)
        'Horas Blanco ART': horas_blanco_art,
        'Horas Quilmes ART': horas_quilmes_art,
        'Horas Papelera ART': horas_papelera_art,
        # Horas Enfermo (columnas P, Q, R)
        'Horas Blanco Enfermo': horas_blanco_enfermo,
        'Horas Quilmes Enfermo': horas_quilmes_enfermo,
        'Horas Papelera Enfermo': horas_papelera_enfermo,
        'Total Horas': horas_blanco + horas_quilmes + horas_papelera,
        'Total Horas Contador': total_horas_contador  # Nuevo campo para ENVIO CONTADOR
    }
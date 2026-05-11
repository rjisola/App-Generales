import os
import sys

# Importar helper de colores (mismo que logic_accountant.py)
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Constantes de multiplicadores (VBA: CH_calculaImporteAmarillo)
MULT_QUILMES  = 1.2     # +20% adicional para horas en proyecto Quilmes (naranja)
MULT_PAPELERA = 1.344   # +34.4% adicional para horas en proyecto Papelera (= 1.2 × 1.12)

# Colores RGB de referencia (VBA: COLOR_QUILMES_*, COLOR_PAPELERA_*)
RGB_NARANJA       = (255, 192, 0)
RGB_NARANJA_CLARO = (255, 204, 102)   # FFCC66 – variante detectada en algunos archivos
RGB_VERDE         = (112, 173, 71)
RGB_MARRON        = (153, 102, 0)     # Alternativo para Papelera (VBA usa ambos)


def _matches_color(rgb, target, tolerance=5):
    """Compara dos tripletas RGB con tolerancia (igual que logic_accountant)."""
    if not rgb:
        return False
    return all(abs(rgb[i] - target[i]) <= tolerance for i in range(3))


def _get_day_color_type(wb_styles, row_idx, col_idx, sheet_name='CALCULAR HORAS'):
    """
    Lee el color de fondo de la celda del día en la hoja de estilos
    y devuelve 'QUILMES', 'PAPELERA' o 'BLANCO' (default).
    """
    if not (wb_styles and row_idx and col_idx):
        return 'BLANCO'

    try:
        if sheet_name not in wb_styles.sheetnames:
            return 'BLANCO'
        ws = wb_styles[sheet_name]
        cell = ws.cell(row=row_idx, column=col_idx)
        if not (cell.fill and cell.fill.fgColor):
            return 'BLANCO'

        sc = cell.fill.start_color
        if not sc:
            return 'BLANCO'

        # Caso 1: Colores de Tema (común en archivos nuevos)
        if sc.type == 'theme':
            if sc.theme == 7: return 'QUILMES'   # Naranja de Tema
            if sc.theme == 8: return 'PAPELERA'  # Verde de Tema
            return 'BLANCO'

        # Caso 2: Colores RGB fijos
        from data_loader import get_rgb_from_openpyxl_color
        rgb = get_rgb_from_openpyxl_color(cell.fill.fgColor)
        if not rgb:
            return 'BLANCO'

        if _matches_color(rgb, RGB_NARANJA) or _matches_color(rgb, RGB_NARANJA_CLARO):
            return 'QUILMES'
        if _matches_color(rgb, RGB_VERDE) or _matches_color(rgb, RGB_MARRON):
            return 'PAPELERA'
    except Exception:
        pass

    return 'BLANCO'


def calculate_amarillo(employee_data, sueldo_contador, day_definitions, config, unify_day_fn,
                        uocra_50, uocra_100,
                        wb_styles=None, row_idx=None, sheet_name='CALCULAR HORAS'):
    """
    Lógica para categoría AMARILLO con multiplicadores por sub-proyecto.

    Equivalente VBA: CH_generaHorasAmarillo + CH_calculaImporteAmarillo

    Umbrales de horas (igual que GRIS):
    - Lunes-Jueves:  > 9hs → extras al 50%
    - Viernes:       > 8hs → extras al 50%
    - Sábados:       <=4hs → al 50%; >4hs → 4 al 50% + resto al 100%
    - Domingos/Feriados: todo al 100%

    Multiplicadores por proyecto (VBA: MULTIPLICADOR_QUILMES / MULTIPLICADOR_PAPELERA):
    - Días con celda NARANJA  (255,192,0)  → tasa × 1.200  (Quilmes)
    - Días con celda VERDE    (112,173,71) → tasa × 1.344  (Papelera)
    - Días con celda BLANCA   (default)   → tasa × 1.000  (sin adicional)

    Si wb_styles=None (modo sin acceso a estilos), aplica tasa uniforme sin multiplicadores
    como fallback seguro.
    """
    # --- Acumuladores separados por sub-proyecto ---
    horas_50_blancas  = 0.0
    horas_50_quilmes  = 0.0
    horas_50_papelera = 0.0

    horas_100_blancas  = 0.0
    horas_100_quilmes  = 0.0
    horas_100_papelera = 0.0

    for day_info in day_definitions:
        day_input = unify_day_fn(employee_data.get(day_info['col_key_in_df']), config)
        horas = day_input['hours']
        if horas <= 0:
            continue

        day_name   = day_info['day_name'].lower()
        is_holiday = day_info.get('is_holiday', False)
        col_idx    = day_info.get('col_idx')

        extras_50  = 0.0
        extras_100 = 0.0

        if is_holiday:
            # Feriados se calculan aparte en logic_payroll para paridad
            continue

        if day_name == 'domingo':
            # Domingos: todo al 100%
            extras_100 = horas

        elif day_name == 'sábado':
            # Sábado: <=4hs al 50%; >4hs → 4 al 50% + resto al 100%
            if horas > 4:
                extras_50  = 4
                extras_100 = horas - 4
            else:
                extras_50 = horas

        else:
            # Lunes a Viernes: límite 9hs (L-J) u 8hs (Vie), extras al 50%
            limite = 9 if day_name != 'viernes' else 8
            if horas > limite:
                extras_50 = horas - limite

        # Si hay horas extras en este día, detectar proyecto por color de celda
        if extras_50 > 0 or extras_100 > 0:
            color_tipo = _get_day_color_type(wb_styles, row_idx, col_idx, sheet_name)

            if color_tipo == 'QUILMES':
                horas_50_quilmes  += extras_50
                horas_100_quilmes += extras_100
            elif color_tipo == 'PAPELERA':
                horas_50_papelera  += extras_50
                horas_100_papelera += extras_100
            else:
                horas_50_blancas  += extras_50
                horas_100_blancas += extras_100

    # --- Cálculo de importes con multiplicadores (VBA: calcularImporteAmarillo) ---
    importe_50 = (
        horas_50_blancas  * uocra_50 +
        horas_50_quilmes  * uocra_50 * MULT_QUILMES +
        horas_50_papelera * uocra_50 * MULT_PAPELERA
    )
    importe_100 = (
        horas_100_blancas  * uocra_100 +
        horas_100_quilmes  * uocra_100 * MULT_QUILMES +
        horas_100_papelera * uocra_100 * MULT_PAPELERA
    )

    # Redondear importes intermedios si es necesario para paridad (VBA suele redondear a 2 decimales o entero)
    total_extras = round(importe_50 + importe_100, 2)

    # Totales de horas (para mostrar en recibo)
    total_horas_50  = horas_50_blancas  + horas_50_quilmes  + horas_50_papelera
    total_horas_100 = horas_100_blancas + horas_100_quilmes + horas_100_papelera

    return sueldo_contador, total_extras, uocra_50, uocra_100, total_horas_50, total_horas_100

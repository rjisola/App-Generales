import os
from decimal import Decimal, ROUND_HALF_UP, ROUND_CEILING
from num2words import num2words
import openpyxl

# --- Configuración Global ---
# Fuente de tarifas: se busca primero el archivo compartido de la app del puerto 20000
# (asistente-sueldos-js/app_20000/) para garantizar que ambas apps usen siempre
# los mismos valores. Si no se encuentra, se usa la copia local como respaldo.
_project_root = os.path.dirname(os.path.dirname(__file__))

_EXCEL_PRIMARIO  = os.path.abspath(os.path.join(
    _project_root, '..', 'asistente-sueldos-js', 'app_20000', '2-VALOR_HORAS_SUELDOS.xlsx'
))
_EXCEL_RESPALDO  = os.path.join(_project_root, '02_CARPETAS', 'Datos', '2-VALOR_HORAS_SUELDOS.xlsx')

# Seleccionar automáticamente la fuente disponible - PRIORIDAD LOCAL
if os.path.exists(_EXCEL_RESPALDO):
    EXCEL_FILE_PATH = _EXCEL_RESPALDO
    print(f"[calculators] Usando Excel local (principal): {_EXCEL_RESPALDO}")
elif os.path.exists(_EXCEL_PRIMARIO):
    EXCEL_FILE_PATH = _EXCEL_PRIMARIO
    print(f"[calculators] Usando Excel compartido (respaldo): {_EXCEL_PRIMARIO}")
else:
    EXCEL_FILE_PATH = _EXCEL_RESPALDO # Por defecto para el error
    print(f"[calculators] ERROR: No se encontró el Excel en ninguna ubicación.")

EXCEL_SHEET_NAME = 'Hoja1'

# Cache para los datos del Excel
_excel_data_cache = {}

def _load_excel_data():
    global _excel_data_cache
    if _excel_data_cache:
        return _excel_data_cache

    try:
        workbook = openpyxl.load_workbook(EXCEL_FILE_PATH, data_only=True)
        sheet = workbook[EXCEL_SHEET_NAME]

        # Cargar tasas horarias para UOCRA/QUILMES/NASA
        hourly_rates = {
            "UOCRA": {},
            "QUILMES": {},
            "NASA": {}
        }
        # UOCRA / QUILMES (Columna D)
        uocra_quilmes_categories = [str(sheet[f'C{r}'].value).strip() for r in range(3, 7) if sheet[f'C{r}'].value]
        for i, cat in enumerate(uocra_quilmes_categories):
            hourly_rates["UOCRA"][cat] = _to_float(sheet[f'D{i+3}'].value)
            hourly_rates["QUILMES"][cat] = _to_float(sheet[f'D{i+3}'].value)

        # NASA (Columna F)
        nasa_categories = [str(sheet[f'E{r}'].value).strip() for r in range(3, 9) if sheet[f'E{r}'].value]
        for i, cat in enumerate(nasa_categories):
            hourly_rates["NASA"][cat] = _to_float(sheet[f'F{i+3}'].value)

        # Cargar datos UECARA
        uecara_data = {"categorias": [], "valores": {}, "adicionales": {}}
        for row in range(3, 11):
            categoria = sheet[f'G{row}'].value
            valor = sheet[f'H{row}'].value
            if categoria and str(categoria).strip():
                cat_clean = str(categoria).strip()
                uecara_data["categorias"].append(cat_clean)
                uecara_data["valores"][cat_clean] = _to_decimal(valor)

        uecara_data["adicionales"] = {
            "Antigüedad": _to_decimal(sheet['H11'].value),
            "Título Universitario": _to_decimal(sheet['H12'].value),
            "Título Técnico": _to_decimal(sheet['H13'].value),
            "Título Secundario": _to_decimal(sheet['H14'].value),
            "Sin Título": Decimal('0')
        }

        # Otros valores
        other_values = {
            "seguro_vida_con": _to_float(sheet['B3'].value),
            "seguro_vida_sin": _to_float(sheet['B4'].value),
        }

        _excel_data_cache = {
            "hourly_rates": hourly_rates,
            "uecara_data": uecara_data,
            "other_values": other_values
        }
        return _excel_data_cache

    except FileNotFoundError:
        print(f"Error: Archivo Excel no encontrado en {EXCEL_FILE_PATH}")
        return None
    except KeyError as e:
        print(f"Error: Hoja o celda no encontrada en Excel: {e}")
        return None
    except Exception as e:
        print(f"Error al cargar datos de Excel: {e}")
        return None

def _to_float(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, Decimal):
        return float(value)
    try:
        s = str(value).strip()
        if not s: return 0.0
        
        if ',' in s:
            s = s.replace('.', '').replace(',', '.')
        elif '.' in s:
            if s.count('.') > 1:
                s = s.replace('.', '')
            else:
                parts = s.split('.')
                if len(parts) > 1 and len(parts[1]) == 3:
                    s = s.replace('.', '')
        
        return float(s)
    except (ValueError, TypeError):
        return 0.0

def _to_decimal(value):
    if value is None:
        return Decimal('0')
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    try:
        s = str(value).strip()
        if not s: return Decimal('0')
        
        if ',' in s:
            s = s.replace('.', '').replace(',', '.')
        elif '.' in s:
            if s.count('.') > 1:
                s = s.replace('.', '')
            else:
                parts = s.split('.')
                if len(parts) > 1 and len(parts[1]) == 3:
                    s = s.replace('.', '')
        
        return Decimal(s)
    except Exception:
        return Decimal('0')

def _format_currency(value):
    try:
        v = _to_float(value)
        s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return s
    except Exception:
        return "0,00"

def _format_currency_decimal(dec_value: Decimal):
    try:
        q = dec_value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        s = f"{float(q):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return s
    except Exception:
        return "0,00"

def _convert_neto_a_letras(neto_valor):
    try:
        v = float(neto_valor)
    except Exception:
        v = 0.0
    if not (0 <= v < 1_000_000_000):
        return "VALOR FUERA DE RANGO"
    parte_entera = int(v)
    dec = Decimal(str(v)) - Decimal(parte_entera)
    parte_decimal = int(dec.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) * 100)
    try:
        texto_entero = num2words(parte_entera, lang='es').upper()
        return f"{texto_entero} PESOS CON {parte_decimal:02d}/100"
    except Exception:
        return "Error en conversión a letras"

def _get_hourly_rate(convenio, categoria, excel_data):
    if not convenio or not categoria:
        return 0.0
    return excel_data["hourly_rates"].get(convenio, {}).get(categoria, 0.0)

def _sum_values(*values):
    total = 0.0
    for val in values:
        total += _to_float(val)
    return total

def calculate_payroll_uocra_quilmes_nasa(data):

    excel_data = _load_excel_data()
    if not excel_data:
        return {"error": "No se pudieron cargar los datos de Excel."}

    convenio = data.get('convenio', '')
    categoria = data.get('categoria', '')
    presentismo_activo = data.get('presentismo', 'Presentismo') == "Presentismo"
    tipo_bono = data.get('bono_tipo', 'No Remunerativo (aporta OS)')

    tasa_horaria = _get_hourly_rate(convenio, categoria, excel_data)

    if tasa_horaria < 0.01:
        return {"error": "Categoría o convenio no encontrado o tasa horaria inválida.", "tasa_horaria": _format_currency(tasa_horaria)}

    # --- Inputs --- (usar _to_float para todos los inputs numéricos)
    horas_50 = _to_float(data.get('horas_50', 0))
    horas_100 = _to_float(data.get('horas_100', 0))
    horas_hormigon = _to_float(data.get('horas_hormigon', 0))
    horas_altura = _to_float(data.get('horas_altura', 0))
    horas_normales = _to_float(data.get('horas_normales', 0))
    horas_feriado = _to_float(data.get('horas_feriado', 0))
    horas_art = _to_float(data.get('horas_art', 0))
    horas_enfermedad = _to_float(data.get('horas_enfermedad', 0))
    dias_vacaciones = _to_float(data.get('dias_vacaciones', 0))
    horas_quincena_anterior = _to_float(data.get('horas_quincena_anterior', 0))
    horas_conv_especial = _to_float(data.get('horas_conv_especial', 0))
    porc_conv_especial = _to_float(data.get('porc_conv_especial', 0))
    porc_presentismo = _to_float(data.get('porc_presentismo', 0))
    porc_adicional = _to_float(data.get('porc_adicional', 0))
    monto_bono = _to_float(data.get('monto_bono', 0))
    porc_ret_ganancias = _to_float(data.get('porc_ret_ganancias', 0))
    porc_ret_judicial = _to_float(data.get('porc_ret_judicial', 0))
    monto_retroactivo = _to_float(data.get('monto_retroactivo', 0))
    seguro_vida_opcion = data.get('seguro_vida_opcion', 'Seg Vida')

    # --- Haberes ---
    importe_50 = horas_50 * 1.5 * tasa_horaria
    importe_100 = horas_100 * 2.0 * tasa_horaria

    importe_hormigon = horas_hormigon * 0.15 * tasa_horaria if convenio == "NASA" else 0.0
    importe_altura = horas_altura * 0.15 * tasa_horaria if convenio in ("QUILMES", "NASA") else 0.0

    importe_normal = horas_normales * tasa_horaria
    importe_feriado = horas_feriado * tasa_horaria
    importe_enfermedad = horas_enfermedad * tasa_horaria
    importe_art = horas_art * tasa_horaria

    # Vacaciones
    importe_vacaciones = 0.0
    if dias_vacaciones > 0 and tasa_horaria > 0:
        if convenio == "NASA":
            horas_por_dia = 8
            importe_vacaciones = dias_vacaciones * horas_por_dia * tasa_horaria * 1.2 * 1.03
        else:
            horas_por_dia = 9
            importe_vacaciones = dias_vacaciones * horas_por_dia * tasa_horaria

    # Quincena anterior
    importe_quincena_anterior = 0.0
    pres_quincena_anterior = 0.0
    if presentismo_activo and horas_quincena_anterior > 0:
        importe_quincena_anterior = horas_quincena_anterior * tasa_horaria
        pres_quincena_anterior = importe_quincena_anterior * 0.2

    # Convenio Especial
    importe_horas_conv_especial = 0.0
    importe_adic_conv_especial = 0.0
    if convenio in ("UOCRA", "NASA") and horas_conv_especial > 0:
        tasa_esp = _get_hourly_rate("UOCRA", categoria, excel_data) if convenio == "NASA" else tasa_horaria
        if tasa_esp > 0:
            importe_horas_conv_especial = horas_conv_especial * tasa_esp
            importe_adic_conv_especial = importe_horas_conv_especial * (porc_conv_especial / 100)

    # Presentismo
    importe_presentismo = 0.0
    base_calculo_presentismo = 0.0
    
    if presentismo_activo: # Calcular base aunque porc sea 0, por si se usa para Adicional
        if convenio in ("UOCRA", "QUILMES"):
            # Alinear con Nueva carpeta\10-CALCULAR_HORAS.pyw: Solo 50, 100, Normal
            base_calculo_presentismo = _sum_values(importe_50, importe_100, importe_normal)
        elif convenio == "NASA":
            # Alinear con Nueva carpeta\10-CALCULAR_HORAS.pyw: 50, 100, Normal, Hormigon, Altura, Feriado
            base_calculo_presentismo = _sum_values(importe_50, importe_100, importe_hormigon, importe_altura, importe_normal, importe_feriado)
            
        if porc_presentismo > 0:
            importe_presentismo = base_calculo_presentismo * (porc_presentismo / 100)

    # Adicional
    importe_adicional = 0.0
    if porc_adicional > 0 and presentismo_activo:
        if convenio == "QUILMES":
            importe_adicional = base_calculo_presentismo * (porc_adicional / 100)
        elif convenio == "NASA":
            # Alinear con Nueva carpeta: Base Adic = Base Pres + Importe Pres
            base_calculo_adicional = base_calculo_presentismo + importe_presentismo
            importe_adicional = base_calculo_adicional * (porc_adicional / 100)

    # Haberes con Descuento (remunerativos = "Total Bruto")
    total_bruto = _sum_values(
        importe_50, importe_100, importe_hormigon, importe_altura, importe_normal, importe_presentismo,
        importe_adicional, importe_feriado, importe_enfermedad, importe_art, importe_vacaciones,
        importe_quincena_anterior, pres_quincena_anterior, importe_adic_conv_especial, importe_horas_conv_especial,
        monto_retroactivo
    )
    if tipo_bono == "Remunerativo":
        total_bruto += monto_bono

    # Deducciones s/bruto
    jubilacion = total_bruto * 0.11
    ley19032 = total_bruto * 0.03
    obra_social_bruto = total_bruto * 0.03
    sindicato_bruto = total_bruto * 0.025

    obra_social_bono = 0.0
    if tipo_bono == "No Remunerativo (aporta OS)" and monto_bono > 0:
        obra_social_bono = monto_bono * 0.03

    seguro_vida = excel_data["other_values"]["seguro_vida_con"] if seguro_vida_opcion == "Seg Vida" else excel_data["other_values"]["seguro_vida_sin"]

    total_deducciones = (
        jubilacion + ley19032 + obra_social_bruto +
        sindicato_bruto + obra_social_bono +
        seguro_vida
    )

    # Neto antes de ganancias (incluye s/desc si corresponde)
    neto_antes_ganancias = total_bruto - total_deducciones
    if tipo_bono in ["No Remunerativo", "No Remunerativo (aporta OS)"]:
        neto_antes_ganancias += monto_bono

    # Retención Ganancias
    ret_ganancias = 0.0
    if porc_ret_ganancias > 0 and neto_antes_ganancias > 0:
        ret_ganancias = neto_antes_ganancias * (porc_ret_ganancias / 100)

    # Neto base y Retención Judicial
    neto_a_cobrar_base = neto_antes_ganancias - ret_ganancias

    importe_ret_judicial = neto_a_cobrar_base * (porc_ret_judicial / 100) if porc_ret_judicial > 0 else 0.0

    # Redondeo: SIEMPRE hacia arriba y se imputa a Haberes s/Desc.
    neto_final_antes_redondeo = neto_a_cobrar_base - importe_ret_judicial
    net_dec = Decimal(str(neto_final_antes_redondeo)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    neto_redondeado_dec = net_dec.quantize(Decimal('1'), rounding=ROUND_CEILING)
    ajuste_redondeo = float(neto_redondeado_dec - net_dec)
    neto_final_redondeado = float(neto_redondeado_dec)

    # Haberes sin Descuento = Bono NR + Redondeo
    haberes_sin_descuento = 0.0
    if tipo_bono in ("No Remunerativo", "No Remunerativo (aporta OS)"):
        haberes_sin_descuento += monto_bono
    haberes_sin_descuento += ajuste_redondeo

    results = {
        "importe_50": _format_currency(importe_50),
        "importe_100": _format_currency(importe_100),
        "importe_hormigon": _format_currency(importe_hormigon),
        "importe_altura": _format_currency(importe_altura),
        "importe_normal": _format_currency(importe_normal),
        "importe_feriado": _format_currency(importe_feriado),
        "importe_enfermedad": _format_currency(importe_enfermedad),
        "importe_art": _format_currency(importe_art),
        "importe_vacaciones": _format_currency(importe_vacaciones),
        "importe_quincena_anterior": _format_currency(importe_quincena_anterior),
        "pres_quincena_anterior": _format_currency(pres_quincena_anterior),
        "importe_horas_conv_especial": _format_currency(importe_horas_conv_especial),
        "importe_adic_conv_especial": _format_currency(importe_adic_conv_especial),
        "importe_presentismo": _format_currency(importe_presentismo),
        "importe_adicional": _format_currency(importe_adicional),
        "total_bruto": _format_currency(total_bruto),
        "haberes_con_descuento": _format_currency(total_bruto),
        "jubilacion": _format_currency(jubilacion),
        "ley19032": _format_currency(ley19032),
        "obra_social_bruto": _format_currency(obra_social_bruto),
        "sindicato_bruto": _format_currency(sindicato_bruto),
        "obra_social_bono": _format_currency(obra_social_bono),
        "seguro_vida": _format_currency(seguro_vida),
        "total_deducciones": _format_currency(total_deducciones),
        "ret_ganancias": _format_currency(ret_ganancias),
        "neto_antes_ret_judicial": _format_currency(neto_a_cobrar_base),
        "importe_ret_judicial": _format_currency(importe_ret_judicial),
        "ajuste_redondeo": _format_currency(ajuste_redondeo),
        "haberes_sin_descuento": _format_currency(haberes_sin_descuento),
        "total_final": _format_currency(neto_final_redondeado),
        "neto_en_letras": _convert_neto_a_letras(neto_final_redondeado),
        "tasa_horaria": _format_currency(tasa_horaria),
        "importe_bono": _format_currency(monto_bono)
    }
    return results

def calculate_uecara(data):
    excel_data = _load_excel_data()
    if not excel_data:
        return {"error": "No se pudieron cargar los datos de Excel."}

    uecara_config = excel_data["uecara_data"]

    categoria = data.get('categoria', '')
    if not categoria or categoria not in uecara_config["categorias"]:
        return {"error": "Categoría UECARA no válida."}

    D = Decimal
    valor_base_mensual = uecara_config["valores"].get(categoria, D('0'))
    if valor_base_mensual <= 0:
        return {"error": "Valor base mensual de UECARA inválido para la categoría seleccionada."}

    # Inputs
    dnt = D(_to_float(data.get('dnt', 0)))
    feriados = D(_to_float(data.get('feriados', 0)))
    anios_antiguedad = D(_to_float(data.get('anios_antiguedad', 0)))
    ajuste_sueldo = _to_decimal(data.get('ajuste_sueldo', 0))
    bono = _to_decimal(data.get('bono', 0))
    retroactivo = _to_decimal(data.get('retroactivo', 0))
    presentismo_activo = data.get('presentismo_opcion', 'Con Presentismo') == "Con Presentismo"
    bono_tipo = data.get('bono_tipo', 'Aporta Obra Social')
    titulo_seleccionado = data.get('titulo', 'Sin Título')
    ganancias_pct = _to_decimal(data.get('ganancias_pct', 0))
    porc_ret_judicial = _to_float(data.get('porc_ret_judicial', 0))

    sueldo_quincenal = valor_base_mensual / D('2')
    descuento_dnt = (valor_base_mensual / D('30')) * dnt
    importe_feriados = (valor_base_mensual / D('25')) * feriados

    valor_mensual_antiguedad = uecara_config["adicionales"].get("Antigüedad", D('0'))
    importe_antiguedad = (anios_antiguedad * valor_mensual_antiguedad) / D('2')

    importe_titulo = uecara_config["adicionales"].get(titulo_seleccionado, D('0')) / D('2')

    importe_presentismo = D('0')
    if presentismo_activo:
        importe_presentismo = (sueldo_quincenal * D('0.10')).quantize(D('1.'), rounding=ROUND_HALF_UP)

    total_bruto = (sueldo_quincenal - descuento_dnt + importe_feriados +
                   importe_presentismo + importe_antiguedad + importe_titulo +
                   ajuste_sueldo + retroactivo)

    if bono_tipo == "Con Retenciones":
        total_bruto += bono

    # Deducciones
    jubilacion = (total_bruto * D('0.11'))
    ley19032 = (total_bruto * D('0.03'))
    obra_social = (total_bruto * D('0.03'))
    sindicato = (total_bruto * D('0.025'))
    ret_ganancias = total_bruto * (ganancias_pct / D('100'))

    obra_social_bono = D('0')
    if bono_tipo == "Aporta Obra Social":
        obra_social_bono = bono * D('0.03')

    total_deducciones = jubilacion + ley19032 + obra_social + sindicato + ret_ganancias + obra_social_bono

    # Neto
    neto_exacto = total_bruto - total_deducciones
    if bono_tipo in ("Sin Retenciones", "Aporta Obra Social"):
        neto_exacto += bono

    # Retención Judicial: se aplica sobre el neto completo (igual que UOCRA/NASA)
    importe_ret_judicial = D('0')
    if porc_ret_judicial > 0:
        importe_ret_judicial = (neto_exacto * D(str(porc_ret_judicial)) / D('100')).quantize(D('0.01'), rounding=ROUND_HALF_UP)
        neto_exacto -= importe_ret_judicial

    neto_redondeado = neto_exacto.quantize(D('1.'), rounding=ROUND_HALF_UP)
    diferencia_redondeo = neto_redondeado - neto_exacto
    total_bruto += diferencia_redondeo # Ajustar bruto por redondeo

    results = {
        "sueldo_quincenal": _format_currency_decimal(sueldo_quincenal),
        "descuento_dnt": _format_currency_decimal(descuento_dnt),
        "importe_feriados": _format_currency_decimal(importe_feriados),
        "importe_antiguedad": _format_currency_decimal(importe_antiguedad),
        "importe_titulo": _format_currency_decimal(importe_titulo),
        "importe_presentismo": _format_currency_decimal(importe_presentismo),
        "bruto": _format_currency_decimal(total_bruto),
        "jubilacion": _format_currency_decimal(jubilacion),
        "ley19032": _format_currency_decimal(ley19032),
        "obra_social": _format_currency_decimal(obra_social),
        "sindicato": _format_currency_decimal(sindicato),
        "ganancias_importe": _format_currency_decimal(ret_ganancias),
        "os_bono": _format_currency_decimal(obra_social_bono),
        "importe_ret_judicial": _format_currency_decimal(importe_ret_judicial),
        "total_deducciones": _format_currency_decimal(total_deducciones + importe_ret_judicial),
        "neto_redondeado": _format_currency_decimal(neto_redondeado),
        "ajuste_redondeo": _format_currency_decimal(diferencia_redondeo),
        "neto_en_letras": _convert_neto_a_letras(float(neto_redondeado)),
        "importe_bono": _format_currency_decimal(bono)
    }
    return results

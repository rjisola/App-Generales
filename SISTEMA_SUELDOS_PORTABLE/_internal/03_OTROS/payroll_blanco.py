def calculate_blanco(employee_data, sueldo_acuerdo, day_definitions, config, unify_day_fn, uocra_50, uocra_100):
    """
    Lógica para categoría BLANCO (UOCRA puro con factor ×1.2).

    Equivalente VBA: CH_generaHorasBlanco + CH_calculaImporteBlanco

    Reglas de horas (igual que GRIS):
    - Lunes-Jueves:  9hs incluidas  → extras al 50%
    - Viernes:       8hs incluidas  → extras al 50%
    - Sábados:       <=4hs → al 50%; >4hs → 4 al 50% + resto al 100%
    - Domingos/Feriados: todo al 100%

    Diferencia vs GRIS:
    - Tarifa BLANCO = tasa_UOCRA × 1.2  (VBA: valorHoraNormal = Range("B1") * 1.2)
    - Las horas normales SÍ se acumulan (pero se incluyen en sueldo_acuerdo)
    """
    # Factor diferenciador BLANCO (20% adicional sobre la tasa UOCRA base)
    FACTOR_BLANCO = 1.2

    v_50 = uocra_50 * FACTOR_BLANCO
    v_100 = uocra_100 * FACTOR_BLANCO

    horas_50 = 0.0
    horas_100 = 0.0

    for day_info in day_definitions:
        day_input = unify_day_fn(employee_data.get(day_info['col_key_in_df']), config)
        horas = day_input['hours']
        if horas <= 0:
            continue

        day_name = day_info['day_name'].lower()
        is_holiday = day_info.get('is_holiday', False)

        if is_holiday:
            # Feriados se calculan aparte en logic_payroll para paridad
            continue

        if day_name == 'domingo':
            # Domingos: todo al 100%
            horas_100 += horas

        elif day_name == 'sábado':
            # Sábado: <=4hs al 50%, >4hs → 4 al 50% y el resto al 100%
            if horas > 4:
                horas_50 += 4
                horas_100 += (horas - 4)
            else:
                horas_50 += horas

        else:
            # Lunes a Viernes: límite 9hs (L-J) u 8hs (Vie), extras al 50%
            limite = 9
            if day_name == 'viernes':
                limite = 8
            if horas > limite:
                horas_50 += (horas - limite)

    # Cálculo de importes base de extras
    importe_50 = horas_50 * v_50
    importe_100 = horas_100 * v_100
    
    # Fondo de Desempleo (VBA: 12% sobre el total de extras)
    fondo_desempleo = (importe_50 + importe_100) * 0.12
    
    total_extras = importe_50 + importe_100 + fondo_desempleo

    return sueldo_acuerdo, total_extras, v_50, v_100, horas_50, horas_100

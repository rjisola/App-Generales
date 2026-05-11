def calculate_azul(employee_data, sueldo_acuerdo, day_definitions, config, unify_day_fn):
    """
    Lógica para categoría AZUL (Capataces).

    Equivalente VBA: CH_generaHorasAzul + CH_calculaImporteAzul

    Reglas generales (capataz estándar):
    - Lunes-Viernes: 10hs incluidas → extras al 50%
    - Sábados:       <=5hs → al 50%; >5hs → 5 al 50% + resto al 100%
    - Tarifas:       v_50 = Acuerdo / 100
                     v_100 = (Acuerdo / 110) × 2

    Excepción ALBORNOZ (Albornoz Claudio Gera):
    - Lunes-Viernes: 12hs incluidas → extras al 50%
    - v_50 = (Acuerdo / 120) × 1.5
    - v_100 = (Acuerdo / 120) × 2.0

    Excepción HOLGADO / SOUZA (Holgado Pedro Atilio, Souza Edgardo Andres):
    - Lunes-Viernes: 10hs incluidas → extras al 50%  (horas igual que estándar)
    - v_50 = (Acuerdo / 120) × 1.5
    - v_100 = (Acuerdo / 120) × 1.5   ← MISMA tasa para 50% y 100% (VBA confirmado)
    """
    empleado_nombre = str(employee_data.get('NOMBRE Y APELLIDO', '')).upper()

    is_albornoz      = "ALBORNOZ" in empleado_nombre
    is_tarifa_esp    = (not is_albornoz) and (
                           "HOLGADO" in empleado_nombre or "SOUZA" in empleado_nombre
                       )

    # --- Tarifas ---
    if is_albornoz:
        v_50  = (sueldo_acuerdo / 120) * 1.5
        v_100 = (sueldo_acuerdo / 120) * 2.0
    elif is_tarifa_esp:
        # Holgado y Souza: misma tasa para horas al 50% y al 100% (VBA Select Case)
        v_50  = (sueldo_acuerdo / 120) * 1.5
        v_100 = (sueldo_acuerdo / 120) * 1.5
    else:
        v_50  = sueldo_acuerdo / 100
        v_100 = (sueldo_acuerdo / 110) * 2

    # Límite de horas normales incluidas L-V
    limite_lv = 12 if is_albornoz else 10

    horas_50  = 0.0
    horas_100 = 0.0

    for day_info in day_definitions:
        day_input = unify_day_fn(employee_data.get(day_info['col_key_in_df']), config)
        horas = day_input['hours']
        if horas <= 0:
            continue

        day_name   = day_info['day_name'].lower()
        is_holiday = day_info.get('is_holiday', False)

        if is_holiday:
            # Feriados se calculan aparte en logic_payroll para paridad
            continue

        if day_name == 'domingo':
            # Domingos: todo al 100%
            horas_100 += horas

        elif day_name == 'sábado':
            # Sábado: <=5hs al 50%; >5hs → 5 al 50% + resto al 100%
            if horas > 5:
                horas_50  += 5
                horas_100 += (horas - 5)
            else:
                horas_50 += horas

        else:
            # Lunes a Viernes: extras más allá del límite → al 50%
            if horas > limite_lv:
                horas_50 += (horas - limite_lv)

    total_extras = (horas_50 * v_50) + (horas_100 * v_100)

    # El sueldo base es el acuerdo íntegro (VBA: sueldo base = Acuerdo)
    sueldo_base = sueldo_acuerdo

    return sueldo_base, total_extras, v_50, v_100, horas_50, horas_100

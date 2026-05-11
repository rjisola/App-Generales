def calculate_celeste(employee_data, sueldo_acuerdo, day_definitions, config, unify_day_fn):
    """
    Lógica para categoría CELESTE (Acuerdos).

    Equivalente VBA: CH_generaHorasCeleste + CH_calculaImporteCeleste

    Reglas:
    - Lunes-Viernes: 10hs incluidas → extras al 50%
    - Sábados:       <=5hs → al 50%; >5hs → 5 al 50% + resto al 100%
    - Domingos/Feriados: todo al 100%
    - Tarifas: v_50 = (Acuerdo / 120) × 1.5
               v_100 = (Acuerdo / 120) × 2.0

    Excepción FERREYRA (Ferreyra David Ismael):
    - Cap fijo de 10hs normales, SIN horas extras (VBA confirmado).
    """
    v_50  = (sueldo_acuerdo / 120) * 1.5
    v_100 = (sueldo_acuerdo / 120) * 2.0

    empleado_nombre = str(employee_data.get('NOMBRE Y APELLIDO', '')).upper()
    is_ferreyra     = "FERREYRA DAVID ISMAEL" in empleado_nombre or \
                      ("FERREYRA" in empleado_nombre and "DAVID" in empleado_nombre)

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
            # Lunes a Viernes: 10hs incluidas
            limite = 10
            if horas > limite:
                if is_ferreyra:
                    # Excepción Ferreyra: cap en 10hs, sin extras (VBA: horasAlCincuenta = 0)
                    pass
                else:
                    horas_50 += (horas - limite)

    total_extras = (horas_50 * v_50) + (horas_100 * v_100)

    # En esta categoría NO hay premios ni pluses adicionales sobre el acuerdo.
    return sueldo_acuerdo, total_extras, v_50, v_100, horas_50, horas_100

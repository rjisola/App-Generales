def calculate_gris(employee_data, sueldo_acuerdo, day_definitions, config, unify_day_fn, uocra_50, uocra_100):
    """
    Lógica para categoría GRIS (Acuerdo + UOCRA).
    - 9hs normales L-J, 8hs normales Viernes.
    - Sábados: 4hs al 50%, resto al 100%.
    - Extras: Se usan valores UOCRA (uocra_50, uocra_100).
    """
    v_50 = uocra_50
    v_100 = uocra_100
    
    horas_50 = 0.0
    horas_100 = 0.0
    
    for day_info in day_definitions:
        day_input = unify_day_fn(employee_data.get(day_info['col_key_in_df']), config)
        horas = day_input['hours']
        if horas <= 0: continue
        
        day_name = day_info['day_name'].lower()
        is_holiday = day_info.get('is_holiday', False)
        
        if is_holiday:
            # Feriados se calculan aparte en logic_payroll para paridad
            continue
            
        if day_name == 'domingo':
            horas_100 += horas
        elif day_name == 'sábado':
            if horas > 4:
                horas_50 += 4
                horas_100 += (horas - 4)
            else:
                horas_50 += horas
        else:
            # Lunes a Viernes
            limite = 9
            if day_name == 'viernes':
                limite = 8
            if horas > limite:
                horas_50 += (horas - limite)
                
    total_extras = (horas_50 * v_50) + (horas_100 * v_100)
    
    total = sueldo_acuerdo + total_extras
    return sueldo_acuerdo, total_extras, v_50, v_100, horas_50, horas_100

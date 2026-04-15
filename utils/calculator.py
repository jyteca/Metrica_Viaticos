"""
Motor de cálculos para la app de Viáticos.
Replica la lógica del Excel original con fórmulas de referencia cruzada.
Actualizado: considera número de técnicos para hospedaje y viáticos.
"""


def calcular_peaje(ref_data, ciudad_destino, tipo_vehiculo="Auto/Camioneta", ida_vuelta=True):
    """
    Calcula el peaje para un destino.
    """
    campo = "auto_ida" if tipo_vehiculo == "Auto/Camioneta" else "camion_ida"
    
    for seccion in ["norte", "sur"]:
        for entry in ref_data.get("peajes_por_destino", {}).get(seccion, []):
            if entry["ciudad"].lower() == ciudad_destino.lower():
                ida = entry[campo]
                factor = 2 if ida_vuelta else 1
                return {
                    "peaje_ida": ida,
                    "peaje_ida_vuelta": ida * factor
                }
    
    return {"peaje_ida": 0, "peaje_ida_vuelta": 0}


def calcular_combustible(ref_data, km_ida, km_vuelta, tipo_vehiculo="Auto/Camioneta"):
    """
    Calcula el costo de combustible.
    """
    comb = ref_data.get("combustible", {})
    
    if tipo_vehiculo == "Auto/Camioneta":
        precio_lt = comb.get("precio_litro_auto", 1450)
        consumo = comb.get("consumo_km_lt_auto", 10)
    else:
        precio_lt = comb.get("precio_litro_camion", 1100)
        consumo = comb.get("consumo_km_lt_camion", 6)
    
    km_totales = km_ida + km_vuelta
    litros_totales = km_totales / consumo if consumo > 0 else 0
    costo_total = litros_totales * precio_lt
    
    return {
        "km_ida": km_ida,
        "km_vuelta": km_vuelta,
        "km_totales": km_totales,
        "precio_litro": precio_lt,
        "consumo_km_lt": consumo,
        "litros_totales": round(litros_totales, 1),
        "costo_total": round(costo_total)
    }


def calcular_alojamiento(ref_data, noches, num_tecnicos, tipo_habitacion="habitacion_doble", rango="promedio"):
    """
    Calcula el costo de alojamiento considerando número de técnicos.
    Si habitación doble, se necesitan ceil(num_tecnicos/2) habitaciones.
    Si single, se necesita 1 por técnico.
    """
    aloj = ref_data.get("alojamiento", {})
    precios = aloj.get(tipo_habitacion, {})
    precio_noche = precios.get(rango, 0)
    
    if tipo_habitacion == "habitacion_doble":
        # Redondeo hacia arriba para parejas
        num_habitaciones = (num_tecnicos + 1) // 2 if num_tecnicos > 0 else 0
    else:
        num_habitaciones = num_tecnicos
    
    total = noches * precio_noche * num_habitaciones
    
    return {
        "noches": noches,
        "num_tecnicos": num_tecnicos,
        "num_habitaciones": num_habitaciones,
        "precio_noche": precio_noche,
        "total": round(total)
    }


def calcular_imprevistos(ref_data, subtotal_gastos):
    """
    Calcula el porcentaje de imprevistos.
    """
    pct = ref_data.get("imprevistos_porcentaje", 0.20)
    total = round(subtotal_gastos * pct)
    
    return {
        "porcentaje": pct,
        "subtotal_gastos": subtotal_gastos,
        "total_imprevistos": total
    }


def calcular_viaticos_tecnicos(tecnicos):
    """
    Suma los viáticos no rendibles de todos los técnicos.
    """
    total = sum(t.get("monto", 0) for t in tecnicos)
    return {
        "tecnicos": tecnicos,
        "total_viaticos_no_rendibles": total,
        "num_tecnicos": len(tecnicos)
    }


def calcular_todo(ref_data, solicitud):
    """
    Calcula todos los costos de una solicitud completa.
    """
    tecnicos = solicitud.get("tecnicos", [])
    num_tecnicos = len(tecnicos)
    
    # Determinar tipo de vehículo para cálculos
    tipo_auto_camion = solicitud.get("tipo_auto_camion", "Auto/Camioneta")
    
    # 1. Peajes
    destino = solicitud.get("destinos", [""])[0] or ""
    ida_vuelta = solicitud.get("ida_vuelta", True)
    peaje = calcular_peaje(ref_data, destino, tipo_auto_camion, ida_vuelta)
    
    # 2. Combustible
    km = get_km_for_destino(ref_data, destino)
    km_vuelta = km if ida_vuelta else 0
    combustible = calcular_combustible(ref_data, km, km_vuelta, tipo_auto_camion)
    
    # 3. Alojamiento
    noches = solicitud.get("noches", 0)
    tipo_hab = solicitud.get("tipo_habitacion", "habitacion_doble")
    rango = solicitud.get("rango_precio_alojamiento", "promedio")
    alojamiento = calcular_alojamiento(ref_data, noches, num_tecnicos, tipo_hab, rango)
    
    # 4. Imprevistos (% sobre peaje + combustible + alojamiento)
    subtotal = peaje["peaje_ida_vuelta"] + combustible["costo_total"] + alojamiento["total"]
    imprevistos = calcular_imprevistos(ref_data, subtotal)
    
    # 5. Viáticos de técnicos
    viaticos = calcular_viaticos_tecnicos(tecnicos)
    
    # 6. Totales por columna
    total_fondo_rendir = (
        peaje["peaje_ida_vuelta"] +
        alojamiento["total"] +
        combustible["costo_total"] +
        imprevistos["total_imprevistos"]
    )
    
    total_viatico_no_rendible = viaticos["total_viaticos_no_rendibles"]
    
    total_general = total_fondo_rendir + total_viatico_no_rendible
    
    return {
        "peaje": peaje,
        "combustible": combustible,
        "alojamiento": alojamiento,
        "imprevistos": imprevistos,
        "viaticos": viaticos,
        "total_fondo_rendir": total_fondo_rendir,
        "total_viatico_no_rendible": total_viatico_no_rendible,
        "total_general": total_general,
        "desglose": {
            "Peajes": peaje["peaje_ida_vuelta"],
            "Hotel": alojamiento["total"],
            "Combustible": combustible["costo_total"],
            "Imprevistos": imprevistos["total_imprevistos"],
            "Viaticos Tecnicos": total_viatico_no_rendible
        }
    }


def get_km_for_destino(ref_data, ciudad):
    """Obtiene km desde Santiago para la ciudad."""
    for item in ref_data.get("kilometraje", []):
        if item["ciudad"].lower() == ciudad.lower():
            return item["km"]
    return 0

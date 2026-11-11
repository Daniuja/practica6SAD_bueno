import argparse
import html
import json
import math
import sys
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
import shapefile
import xlrd

CARPETA_PRINCIPAL = Path(__file__).resolve().parent
CARPETA_DATOS = CARPETA_PRINCIPAL / "datos"
CARPETA_MAPAS = CARPETA_PRINCIPAL / "Mapas"
CRITERIOS = ("crimenes", "robos", "vehiculos", "orden_publico", "drogas")
NOMBRES_CRITERIOS = {
    "crimenes": "Crímenes",
    "robos": "Robos",
    "vehiculos": "Vehículos",
    "orden_publico": "Orden público",
    "drogas": "Drogas",
}
DATOS_ANDALUCIA = [
    {"provincia": "Almería", "crimenes": 23, "robos": 17284, "vehiculos": 2101, "orden_publico": 387, "drogas": 267},
    {"provincia": "Cádiz", "crimenes": 61, "robos": 26571, "vehiculos": 3945, "orden_publico": 722, "drogas": 818},
    {"provincia": "Córdoba", "crimenes": 20, "robos": 14508, "vehiculos": 1691, "orden_publico": 219, "drogas": 148},
    {"provincia": "Granada", "crimenes": 32, "robos": 19139, "vehiculos": 2299, "orden_publico": 368, "drogas": 622},
    {"provincia": "Huelva", "crimenes": 8, "robos": 11222, "vehiculos": 1409, "orden_publico": 156, "drogas": 139},
    {"provincia": "Jaén", "crimenes": 17, "robos": 7548, "vehiculos": 718, "orden_publico": 175, "drogas": 129},
    {"provincia": "Málaga", "crimenes": 58, "robos": 46137, "vehiculos": 6377, "orden_publico": 830, "drogas": 934},
    {"provincia": "Sevilla", "crimenes": 86, "robos": 52849, "vehiculos": 9662, "orden_publico": 716, "drogas": 415},
]
PROVINCIAS_ANDALUCIA = ("Almería", "Cádiz", "Córdoba", "Granada", "Huelva", "Jaén", "Málaga", "Sevilla")
EXCEL_ANDALUCIA = CARPETA_DATOS / "Andalucía.xls"
EXCEL_NACIONAL_2017 = CARPETA_DATOS / "Datos_nacionales_2017_por_10000_ES.xlsx"
EXCEL_NACIONAL_2023 = CARPETA_DATOS / "Datos_nacionales_2023_por_10000_ES.xlsx"
EXCEL_HABITANTES = CARPETA_DATOS / "habitantes.xlsx"
MAPA_ANDALUCIA = CARPETA_MAPAS / "Andalucía" / "da03_provincia.shp"
MAPA_ESPANA = None
for ruta in CARPETA_MAPAS.rglob("*.shp"):
    if ruta.name.startswith("recintos_"):
        MAPA_ESPANA = ruta
        break

# Matriz de comparación por pares de la tabla 4.
MATRIZ_COMPARACION = [
    [1, 5, 3, 7, 3],
    [1 / 5, 1, 3, 1 / 2, 1 / 4],
    [1 / 3, 1 / 3, 1, 1 / 5, 1 / 5],
    [1 / 7, 2, 5, 1, 1 / 5],
    [1 / 3, 4, 5, 5, 1],
]

# Puntos de las funciones de preferencia de la tabla 6, ejercicio 1.
PUNTOS_EJERCICIO_1 = {
    "crimenes": [0, 20, 41.2, 82.4, 100, 123.6, 160, 164.8, 206],
    "robos": [0, 5000, 50914.4, 95000, 101828.8, 152743.2, 180000, 203657.6, 254572],
    "vehiculos": [0, 5000, 7025.4, 14050.8, 19000, 21076.2, 27000, 28101.6, 35127],
    "orden_publico": [0, 300, 764.8, 1529.6, 1800, 2294.4, 2500, 3059.2, 3824],
    "drogas": [0, 300, 321.6, 600, 643.2, 900, 964.8, 1286.4, 1608],
}
PRIORIDADES_EJERCICIO_1 = {
    "crimenes": [0.336, 0.277, 0.132, 0.086, 0.058, 0.032, 0.029, 0.026, 0.023],
    "robos": [0.320, 0.261, 0.139, 0.075, 0.074, 0.047, 0.032, 0.031, 0.022],
    "vehiculos": [0.459, 0.138, 0.124, 0.069, 0.063, 0.048, 0.040, 0.034, 0.026],
    "orden_publico": [0.308, 0.259, 0.149, 0.071, 0.066, 0.060, 0.034, 0.032, 0.023],
    "drogas": [0.385, 0.262, 0.112, 0.070, 0.062, 0.035, 0.033, 0.023, 0.018],
}

# Prioridades locales en los puntos RP1–RP6.
PRIORIDADES_RP = {
    "crimenes": [0.336, 0.132, 0.086, 0.032, 0.026, 0.023],
    "robos": [0.320, 0.139, 0.074, 0.047, 0.031, 0.022],
    "vehiculos": [0.459, 0.124, 0.069, 0.048, 0.034, 0.026],
    "orden_publico": [0.308, 0.149, 0.071, 0.060, 0.032, 0.023],
    "drogas": [0.385, 0.112, 0.062, 0.033, 0.023, 0.018],
}

PERFILES_ANDALUCIA = {
    "crimenes": [20, 100, 160],
    "robos": [5000, 95000, 180000],
    "vehiculos": [5000, 19000, 27000],
    "orden_publico": [300, 1800, 2500],
    "drogas": [300, 600, 900],
}
PERFILES_ANDALUCIA_NORMALIZADOS = {
    "crimenes": [0.26, 0.34, 0.41],
    "robos": [194.29, 215.32, 261.32],
    "vehiculos": [22.95, 28.44, 36.20],
    "orden_publico": [2.87, 3.86, 5.32],
    "drogas": [2.06, 3.23, 6.25],
}
RPS_ANDALUCIA = {
    "crimenes": [0, 0.10, 0.20, 0.30, 0.39, 0.49],
    "robos": [0, 56.59, 113.18, 169.77, 226.35, 282.94],
    "vehiculos": [0, 9.96, 19.93, 29.89, 39.85, 49.82],
    "orden_publico": [0, 1.17, 2.33, 3.50, 4.66, 5.83],
    "drogas": [0, 1.36, 2.73, 4.09, 5.45, 6.81],
}
PUNTOS_ANDALUCIA = {
    "crimenes": [0, 0.10, 0.20, 0.26, 0.30, 0.34, 0.39, 0.41, 0.49],
    "robos": [0, 56.59, 113.18, 169.77, 194.29, 215.32, 226.35, 261.32, 282.94],
    "vehiculos": [0, 9.96, 19.93, 22.95, 28.44, 29.89, 36.20, 39.85, 49.82],
    "orden_publico": [0, 1.17, 2.33, 2.87, 3.50, 3.86, 4.66, 5.32, 5.83],
    "drogas": [0, 1.36, 2.06, 2.73, 3.23, 4.09, 5.45, 6.25, 6.81],
}
PRIORIDADES_ANDALUCIA = {
    "crimenes": [0.34, 0.13, 0.09, 0.05, 0.03, 0.03, 0.03, 0.03, 0.02],
    "robos": [0.32, 0.14, 0.07, 0.05, 0.04, 0.03, 0.03, 0.03, 0.02],
    "vehiculos": [0.46, 0.12, 0.07, 0.06, 0.05, 0.05, 0.04, 0.03, 0.03],
    "orden_publico": [0.31, 0.15, 0.07, 0.07, 0.06, 0.05, 0.03, 0.03, 0.02],
    "drogas": [0.39, 0.11, 0.09, 0.06, 0.05, 0.03, 0.02, 0.02, 0.02],
}
PERFILES_NACIONALES = {
    "crimenes": [0.17, 0.25, 0.36],
    "robos": [129.18, 179.38, 268.99],
    "vehiculos": [7.42, 13.69, 30.94],
    "orden_publico": [2.73, 3.38, 4.52],
    "drogas": [1.46, 2.00, 2.77],
}
def calcular_pesos_ahp(matriz):
    """Calcula los pesos AHP y la razón de consistencia."""
    tamano = len(matriz)
    vector = [1 / tamano] * tamano
    for _ in range(10000):
        producto = []
        for fila in matriz:
            valor = 0
            for j in range(tamano):
                valor += fila[j] * vector[j]
            producto.append(valor)
        suma = sum(producto)
        vector_nuevo = []
        for valor in producto:
            vector_nuevo.append(valor / suma)
        diferencia = 0
        for indice in range(tamano):
            cambio = abs(vector_nuevo[indice] - vector[indice])
            if cambio > diferencia:
                diferencia = cambio
        if diferencia < 1e-13:
            vector = vector_nuevo
            break
        vector = vector_nuevo
    producto = []
    for fila in matriz:
        valor = 0
        for j in range(tamano):
            valor += fila[j] * vector[j]
        producto.append(valor)
    valor_propio_maximo = 0
    for indice in range(tamano):
        valor_propio_maximo += producto[indice] / vector[indice]
    valor_propio_maximo = valor_propio_maximo / tamano
    indice_consistencia = (valor_propio_maximo - tamano) / (tamano - 1)
    indice_aleatorio = 1.12  # Índice aleatorio de Saaty para una matriz 5x5.
    return vector, indice_consistencia / indice_aleatorio


def interpolar(valor, puntos_x, puntos_y):
    if len(puntos_x) != len(puntos_y) or not puntos_x:
        raise ValueError("Cada función debe tener igual número de valores RP y prioridades.")
    for indice in range(len(puntos_x) - 1):
        if puntos_x[indice] >= puntos_x[indice + 1]:
            raise ValueError("Los puntos RP deben estar estrictamente ordenados.")
    valor = min(max(valor, puntos_x[0]), puntos_x[-1])
    for indice in range(len(puntos_x) - 1):
        if valor <= puntos_x[indice + 1]:
            proporcion = (valor - puntos_x[indice]) / (puntos_x[indice + 1] - puntos_x[indice])
            return puntos_y[indice] + proporcion * (puntos_y[indice + 1] - puntos_y[indice])
    return puntos_y[-1]


def calcular_rps_nacionales(filas):
    if not filas:
        raise ValueError("No hay datos nacionales para calcular los puntos de referencia.")
    puntos = {}
    for criterio in CRITERIOS:
        valor_maximo = max(fila[criterio] for fila in filas)
        puntos[criterio] = []
        for indice in range(6):
            puntos[criterio].append(round(valor_maximo * indice / 5, 2))
    return puntos


def calcular_prioridades_perfiles(criterio, perfiles, puntos_rp, prioridades_rp):
    prioridades = []
    for valor in perfiles[criterio]:
        prioridades.append(interpolar(valor, puntos_rp[criterio], prioridades_rp[criterio]))
    return prioridades


def clasificar_provincias(filas, pesos, perfiles, puntos_rp, prioridades_rp):
    prioridades_perfiles = {}
    for criterio in CRITERIOS:
        prioridades_perfiles[criterio] = calcular_prioridades_perfiles(
            criterio, perfiles, puntos_rp, prioridades_rp
        )

    perfiles_globales = []
    for indice_perfil in range(3):
        prioridad = 0
        for indice, criterio in enumerate(CRITERIOS):
            prioridad += pesos[indice] * prioridades_perfiles[criterio][indice_perfil]
        perfiles_globales.append(prioridad)
    if not (perfiles_globales[0] > perfiles_globales[1] > perfiles_globales[2]):
        raise ValueError(
            "Los perfiles calculados no están ordenados Alto > Medio > Bajo "
            f"({perfiles_globales}). Revise los datos de calibración."
        )

    resultados = []
    for fila in filas:
        prioridades_locales = {}
        for criterio in CRITERIOS:
            prioridades_locales[criterio] = interpolar(
                float(fila[criterio]), puntos_rp[criterio], prioridades_rp[criterio]
            )
        prioridad_global = 0
        for indice, criterio in enumerate(CRITERIOS):
            prioridad_global += pesos[indice] * prioridades_locales[criterio]
        indice_clase = 0
        distancia_menor = abs(prioridad_global - perfiles_globales[0])
        for indice in range(1, len(perfiles_globales)):
            distancia = abs(prioridad_global - perfiles_globales[indice])
            if distancia < distancia_menor:
                indice_clase = indice
                distancia_menor = distancia
        clases = ("Alto", "Medio", "Bajo")
        resultado = {"provincia": fila["provincia"]}
        for criterio in CRITERIOS:
            resultado[f"tasa_{criterio}"] = fila[criterio]
            resultado[f"prioridad_{criterio}"] = prioridades_locales[criterio]
        resultado["prioridad_global"] = prioridad_global
        resultado["nivel_seguridad"] = clases[indice_clase]
        resultados.append(resultado)
    return ordenar_por_seguridad(resultados)


def ordenar_por_seguridad(filas):
    """Ordena las provincias de mayor a menor prioridad global."""
    ordenadas = []
    for fila in filas:
        posicion = 0
        while posicion < len(ordenadas):
            otra = ordenadas[posicion]
            tiene_mas_prioridad = fila["prioridad_global"] > otra["prioridad_global"]
            mismo_nombre_antes = (
                fila["prioridad_global"] == otra["prioridad_global"]
                and fila["provincia"].casefold() < otra["provincia"].casefold()
            )
            if tiene_mas_prioridad or mismo_nombre_antes:
                break
            posicion += 1
        ordenadas.insert(posicion, fila)
    return ordenadas


def normalizar_tasas(filas, poblaciones):
    poblaciones_normalizadas = {}
    for nombre, habitantes in poblaciones.items():
        poblaciones_normalizadas[normalizar_nombre(nombre)] = habitantes
    normalizados = []
    faltantes = []
    for fila in filas:
        poblacion = poblaciones_normalizadas.get(normalizar_nombre(fila["provincia"]))
        if poblacion is None:
            faltantes.append(fila["provincia"])
            continue
        if poblacion <= 0:
            raise ValueError(f"La población de {fila['provincia']} debe ser mayor que cero.")
        tasas_provincia = {"provincia": fila["provincia"]}
        for criterio in CRITERIOS:
            tasas_provincia[criterio] = fila[criterio] * 10000 / poblacion
        normalizados.append(tasas_provincia)
    if faltantes:
        raise ValueError("Falta la población de: " + ", ".join(faltantes))
    return normalizados


def corregir_nombre_provincia(valor):
    clave = normalizar_nombre(valor)
    aliases = {
        "araba/alava": "Álava",
        "alacant/alicante": "Alicante",
        "alicante/alacant": "Alicante",
        "balears (illes)": "Illes Balears",
        "balears, illes": "Illes Balears",
        "illes balears": "Illes Balears",
        "coruna (a)": "A Coruña",
        "coruna, a": "A Coruña",
        "a coruna": "A Coruña",
        "castello/castellon": "Castellón",
        "castellon/castello": "Castellón",
        "rioja (la)": "La Rioja",
        "rioja, la": "La Rioja",
        "la rioja": "La Rioja",
        "palmas (las)": "Las Palmas",
        "palmas, las": "Las Palmas",
        "valencia/valencia": "Valencia",
    }
    return aliases.get(clave, valor.strip())


def leer_tasas_excel(ruta):
    libro = load_workbook(ruta, data_only=True, read_only=True)
    hoja = libro.worksheets[0]
    iterador = hoja.iter_rows(values_only=True)
    encabezados = next(iterador, None)
    if not encabezados:
        raise ValueError(f"{ruta}: el libro no contiene una cabecera.")
    columnas = {}
    for indice, valor in enumerate(encabezados):
        nombre_columna = normalizar_nombre(str(valor or "")).replace(" ", "")
        if nombre_columna in ("provincia", "crimenes", "robos", "vehiculos", "drogas"):
            columnas[nombre_columna] = indice
        elif nombre_columna == "ordenpublico":
            columnas["orden_publico"] = indice
    obligatorios = {"provincia"}
    for criterio in CRITERIOS:
        obligatorios.add(criterio)
    if not obligatorios.issubset(columnas):
        raise ValueError(f"{ruta}: no se encuentran las cinco columnas de tasas y la provincia.")
    filas = []
    vistos = set()
    for linea, valores in enumerate(iterador, 2):
        provincia = valores[columnas["provincia"]] if columnas["provincia"] < len(valores) else None
        if not isinstance(provincia, str) or not provincia.strip():
            continue
        fila = {"provincia": corregir_nombre_provincia(provincia)}
        clave = normalizar_nombre(fila["provincia"])
        if clave in vistos:
            raise ValueError(f"{ruta}, fila {linea}: provincia duplicada: {provincia}.")
        vistos.add(clave)
        for criterio in CRITERIOS:
            indice = columnas[criterio]
            valor = valores[indice] if indice < len(valores) else None
            if not isinstance(valor, (int, float)) or not math.isfinite(valor) or valor < 0:
                raise ValueError(f"{ruta}, fila {linea}: tasa no válida para {provincia}, {criterio}.")
            fila[criterio] = float(valor)
        filas.append(fila)
    if not filas:
        raise ValueError(f"{ruta}: el libro no contiene datos provinciales.")
    return filas


def leer_datos_andalucia(ruta=EXCEL_ANDALUCIA):
    libro = xlrd.open_workbook(ruta)
    hoja = libro.sheet_by_index(0)
    provincias_conocidas = {}
    for nombre in PROVINCIAS_ANDALUCIA:
        provincias_conocidas[normalizar_nombre(nombre)] = nombre
    provincia_actual = None
    datos = {}
    for indice_fila in range(hoja.nrows):
        valores = hoja.row_values(indice_fila)
        primero = str(valores[0]).strip() if valores else ""
        if normalizar_nombre(primero) in provincias_conocidas:
            provincia_actual = provincias_conocidas[normalizar_nombre(primero)]
            datos[provincia_actual] = {}
        if not provincia_actual or len(valores) < 4:
            continue
        etiqueta = normalizar_nombre(str(valores[2])).replace(" ", "")
        criterio = None
        if "crimen" in etiqueta:
            criterio = "crimenes"
        elif "robo" in etiqueta:
            criterio = "robos"
        elif "vehiculo" in etiqueta:
            criterio = "vehiculos"
        elif "orden" in etiqueta:
            criterio = "orden_publico"
        elif "droga" in etiqueta:
            criterio = "drogas"
        if criterio:
            datos[provincia_actual][criterio] = valores[3]
    faltan_provincias = set(datos) != set(PROVINCIAS_ANDALUCIA)
    faltan_criterios = False
    for fila in datos.values():
        if set(fila) != set(CRITERIOS):
            faltan_criterios = True
    if faltan_provincias or faltan_criterios:
        raise ValueError(f"{ruta}: no se pudieron leer los cinco totales de las ocho provincias.")
    filas = []
    for provincia in PROVINCIAS_ANDALUCIA:
        fila = {"provincia": provincia}
        for criterio in CRITERIOS:
            fila[criterio] = datos[provincia][criterio]
        filas.append(fila)
    return filas


def leer_poblaciones_excel(ruta=EXCEL_HABITANTES):
    libro = load_workbook(ruta, data_only=True, read_only=True)
    hoja = libro.worksheets[0]
    poblaciones = {}
    for indice_fila, fila in enumerate(hoja.iter_rows(values_only=True), 1):
        if len(fila) < 3 or not isinstance(fila[1], str):
            continue
        nombre = fila[1].strip()
        valor = fila[2]
        if normalizar_nombre(nombre) in {"ciudad", "provincia", "nombre"}:
            continue
        if not nombre or not isinstance(valor, (int, float)):
            continue
        poblacion = float(valor)
        if not math.isfinite(poblacion) or poblacion <= 0:
            raise ValueError(f"{ruta}, fila {indice_fila}: población no válida para {nombre}.")
        clave = normalizar_nombre(corregir_nombre_provincia(nombre))
        if clave in poblaciones:
            raise ValueError(f"{ruta}, fila {indice_fila}: población duplicada para {nombre}.")
        poblaciones[clave] = poblacion
    if not poblaciones:
        raise ValueError(f"{ruta}: no se encontraron datos de población.")
    return poblaciones


def normalizar_nombre(valor):
    texto = unicodedata.normalize("NFKD", valor.casefold())
    nombre_sin_tildes = ""
    for caracter in texto:
        if not unicodedata.combining(caracter):
            nombre_sin_tildes += caracter
    return nombre_sin_tildes


def leer_geometrias_mapa(ruta):
    if ruta.suffix.lower() != ".shp":
        with ruta.open("r", encoding="utf-8") as archivo:
            return json.load(archivo)
    lector = shapefile.Reader(str(ruta), encoding="utf-8")
    nombres_campos = []
    for campo in lector.fields[1:]:
        nombres_campos.append(campo[0])

    campo_nombre = None
    nombres_posibles = ("PROVINCIA", "NAMEUNIT", "name", "NAME_2")
    for posible in nombres_posibles:
        if posible in nombres_campos:
            campo_nombre = posible
            break
    if not campo_nombre:
        raise ValueError(f"{ruta}: no se encuentra el atributo con nombre provincial.")
    indice_nombre = nombres_campos.index(campo_nombre)
    elementos = []
    for registro in lector.iterShapeRecords():
        forma = registro.shape
        inicios = list(forma.parts)
        inicios.append(len(forma.points))
        anillos = []
        for indice in range(len(inicios) - 1):
            puntos_anillo = []
            for punto in forma.points[inicios[indice]:inicios[indice + 1]]:
                puntos_anillo.append([float(punto[0]), float(punto[1])])
            anillos.append(puntos_anillo)
        elementos.append({
            "type": "Feature",
            "properties": {"provincia": str(registro.record[indice_nombre])},
            "geometry": {"type": "Polygon", "coordinates": anillos},
        })
    return {"type": "FeatureCollection", "features": elementos}


def guardar_mapa(ruta, resultados, ruta_geojson, permitir_faltantes=False):
    geojson = leer_geometrias_mapa(ruta_geojson)
    if geojson.get("type") != "FeatureCollection":
        raise ValueError(f"{ruta_geojson}: el GeoJSON no contiene una lista de provincias.")
    resultado_por_nombre = {}
    for fila in resultados:
        clave = normalizar_nombre(corregir_nombre_provincia(fila["provincia"]))
        resultado_por_nombre[clave] = fila
    clases = {"Alto": "#2e8b57", "Medio": "#f0c84b", "Bajo": "#c94c4c"}
    entradas = []
    faltantes = set(resultado_por_nombre)
    coordenadas = []
    for elemento in geojson.get("features", []):
        propiedades = elemento.get("properties") or {}
        nombre = None
        nombres_posibles = ("provincia", "name", "NAME_2", "NAME_1", "NAME", "nom")
        for posible in nombres_posibles:
            if propiedades.get(posible):
                nombre = str(propiedades[posible])
                break
        if not nombre:
            continue
        clave = normalizar_nombre(corregir_nombre_provincia(nombre))
        resultado = resultado_por_nombre.get(clave)
        if resultado is None:
            continue
        faltantes.discard(clave)
        geometria = elemento.get("geometry") or {}
        if geometria.get("type") == "Polygon":
            poligonos = geometria.get("coordinates", [])
        elif geometria.get("type") == "MultiPolygon":
            poligonos = []
            for poligono in geometria.get("coordinates", []):
                for anillo in poligono:
                    poligonos.append(anillo)
        else:
            poligonos = []
        if not poligonos:
            raise ValueError(f"{ruta_geojson}: geometría no válida para {nombre}.")
        entradas.append((nombre, resultado["nivel_seguridad"], poligonos))
        for poligono in poligonos:
            for punto in poligono:
                coordenadas.append(punto)
    nombres_faltantes = []
    for clave in sorted(faltantes):
        nombres_faltantes.append(resultado_por_nombre[clave]["provincia"])
    if faltantes and not permitir_faltantes:
        raise ValueError(
            "El mapa no contiene las provincias de: " + ", ".join(nombres_faltantes)
            + ". Comprueba la propiedad provincial (provincia/name/NAME_2)."
        )
    puntos_peninsula = []
    puntos_canarias = []
    for punto in coordenadas:
        if punto[0] > -10:
            puntos_peninsula.append(punto)
        else:
            puntos_canarias.append(punto)
    if not puntos_peninsula:
        raise ValueError(f"{ruta_geojson}: no se encontraron geometrías peninsulares.")
    min_x = min(punto[0] for punto in puntos_peninsula)
    max_x = max(punto[0] for punto in puntos_peninsula)
    min_y = min(punto[1] for punto in puntos_peninsula)
    max_y = max(punto[1] for punto in puntos_peninsula)
    plot_x, plot_y, plot_width, plot_height = 55, 35, 840, 590
    scale = min(plot_width / (max_x - min_x), plot_height / (max_y - min_y))
    drawn_width, drawn_height = (max_x - min_x) * scale, (max_y - min_y) * scale
    offset_x = plot_x + (plot_width - drawn_width) / 2
    offset_y = plot_y + (plot_height - drawn_height) / 2

    def crear_trazado(anillo, limites, caja):
        bx0, bx1, by0, by1 = limites
        ox, oy, width, height = caja
        proporcion = min(width / (bx1 - bx0), height / (by1 - by0))
        shift_x = ox + (width - (bx1 - bx0) * proporcion) / 2
        shift_y = oy + (height - (by1 - by0) * proporcion) / 2
        comandos = []
        for indice, punto in enumerate(anillo):
            longitud = punto[0]
            latitud = punto[1]
            valor_x = shift_x + (longitud - bx0) * proporcion
            valor_y = shift_y + height - (
                (latitud - by0) * proporcion + (height - (by1 - by0) * proporcion) / 2
            )
            letra = "M" if indice == 0 else "L"
            comandos.append(letra + f"{valor_x:.2f},{valor_y:.2f}")
        return " ".join(comandos) + " Z"

    partes_svg = []
    for nombre, nivel, poligonos in entradas:
        for anillo in poligonos:
            if not anillo:
                continue
            longitud_media = 0
            for punto in anillo:
                longitud_media += punto[0]
            longitud_media = longitud_media / len(anillo)
            es_canarias = longitud_media <= -10
            if es_canarias:
                if not puntos_canarias:
                    continue
                limites = (
                    min(punto[0] for punto in puntos_canarias),
                    max(punto[0] for punto in puntos_canarias),
                    min(punto[1] for punto in puntos_canarias),
                    max(punto[1] for punto in puntos_canarias),
                )
                caja = (45, 490, 170, 105)
            else:
                caja = (offset_x, offset_y, drawn_width, drawn_height)
                limites = (min_x, max_x, min_y, max_y)
            trazado = crear_trazado(anillo, limites, caja)
            titulo = html.escape(f"{nombre}: {nivel}")
            partes_svg.append(
                f'<path d="{trazado}" fill="{clases[nivel]}" stroke="#fff" stroke-width="1.1">'
                f"<title>{titulo}</title></path>"
            )
    leyenda = "".join(
        f'<span><i style="background:{color}"></i>{etiqueta}</span>'
        for etiqueta, color in clases.items()
    )
    contenido = f"""<!doctype html>
<html lang="es"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Clasificación provincial de seguridad</title>
<style>
body{{font:16px system-ui,sans-serif;margin:24px;color:#222}}h1{{font-size:1.35rem}}
svg{{width:min(100%,900px);height:auto;border:1px solid #ddd;background:#f8fafc}}
path:hover{{stroke:#111;stroke-width:2.5}}.legend{{display:flex;gap:22px;margin:12px 0}}
i{{display:inline-block;width:14px;height:14px;border:1px solid #666;margin-right:6px;vertical-align:-2px}}
</style>
<h1>Clasificación provincial del nivel de seguridad</h1>
<div class="legend">{leyenda}</div>
<svg viewBox="0 0 950 680" role="img" aria-label="Mapa provincial coloreado por nivel de seguridad">
{"".join(partes_svg)}
</svg>
<p>Pasa el cursor por una provincia para consultar su clase. Verde: Alto; amarillo: Medio; rojo: Bajo.</p>
{"<p>Sin geometría en el shapefile: " + html.escape(", ".join(nombres_faltantes)) + ".</p>" if nombres_faltantes else ""}
</html>"""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(contenido, encoding="utf-8")
    return nombres_faltantes


def mostrar_transiciones(resultados_2017, resultados_2023):
    provincias_2017 = {}
    provincias_2023 = {}
    for fila in resultados_2017:
        provincias_2017[fila["provincia"].casefold()] = fila
    for fila in resultados_2023:
        provincias_2023[fila["provincia"].casefold()] = fila
    if set(provincias_2017) != set(provincias_2023):
        solo_2017 = sorted(set(provincias_2017) - set(provincias_2023))
        solo_2023 = sorted(set(provincias_2023) - set(provincias_2017))
        raise ValueError(f"Las provincias deben coincidir. Solo 2017: {solo_2017}; solo 2023: {solo_2023}")
    clases = ("Alto", "Medio", "Bajo")
    matriz = {}
    for clase_anterior in clases:
        matriz[clase_anterior] = {}
        for clase_nueva in clases:
            matriz[clase_anterior][clase_nueva] = 0
    for anterior in ordenar_por_seguridad(resultados_2017):
        clave = anterior["provincia"].casefold()
        nuevo = provincias_2023[clave]
        matriz[anterior["nivel_seguridad"]][nuevo["nivel_seguridad"]] += 1
    print("\nTabla de transición por provincia (ordenada por seguridad en 2017):")
    print(f"{'Provincia':18} {'2017':10} {'2023':10}")
    for anterior in ordenar_por_seguridad(resultados_2017):
        nuevo = provincias_2023[anterior["provincia"].casefold()]
        print(
            f"{anterior['provincia']:18} "
            f"{anterior['nivel_seguridad']:10} {nuevo['nivel_seguridad']:10}"
        )
    print("\nMatriz de cambios (filas: 2017; columnas: 2023):")
    print("               Alto  Medio  Bajo")
    for clase in clases:
        print(
            f"{clase:10} {matriz[clase]['Alto']:5} "
            f"{matriz[clase]['Medio']:6} {matriz[clase]['Bajo']:5}"
        )


def ejecutar_ejercicio(opciones):
    pesos, razon_consistencia = calcular_pesos_ahp(MATRIZ_COMPARACION)
    print("\nPesos AHP:")
    for criterio, peso in zip(CRITERIOS, pesos):
        print(f"  {NOMBRES_CRITERIOS[criterio]:16} {peso:.4f}")
    if razon_consistencia <= 0.10:
        print(f"Razón de consistencia: {razon_consistencia:.4f} (aceptable)")
    else:
        print(f"Razón de consistencia: {razon_consistencia:.4f} (revisar matriz)")
    carpeta_salida = Path(opciones.salida)
    if opciones.ejercicio == "1":
        resultados = clasificar_provincias(
            leer_datos_andalucia(), pesos, PERFILES_ANDALUCIA,
            PUNTOS_EJERCICIO_1, PRIORIDADES_EJERCICIO_1,
        )
    elif opciones.ejercicio == "2":
        datos_originales = leer_datos_andalucia()
        poblaciones = leer_poblaciones_excel()
        tasas = normalizar_tasas(datos_originales, poblaciones)
        resultados = clasificar_provincias(
            tasas, pesos, PERFILES_ANDALUCIA_NORMALIZADOS,
            PUNTOS_ANDALUCIA, PRIORIDADES_ANDALUCIA,
        )
    elif opciones.ejercicio == "3":
        tasas = leer_tasas_excel(Path(opciones.datos) if opciones.datos else EXCEL_NACIONAL_2017)
        if len(tasas) != 50:
            raise ValueError(f"El ejercicio 3 requiere 50 provincias; se han leído {len(tasas)}.")
        rps_nacionales = calcular_rps_nacionales(tasas)
        print(
            "\nEl libro nacional de 2017 ya contiene tasas por 10.000 habitantes. "
            "Sin los recuentos y la población originales no se puede verificar "
            "la normalización desde su origen."
        )
        resultados = clasificar_provincias(
            tasas, pesos, PERFILES_NACIONALES, rps_nacionales,
            PRIORIDADES_RP,
        )
    else:
        tasas_2017 = leer_tasas_excel(
            Path(opciones.datos_2017) if opciones.datos_2017 else EXCEL_NACIONAL_2017
        )
        tasas_2023 = leer_tasas_excel(
            Path(opciones.datos_2023) if opciones.datos_2023 else EXCEL_NACIONAL_2023
        )
        if len(tasas_2017) != 50 or len(tasas_2023) != 50:
            raise ValueError("El ejercicio 4 requiere datos de 50 provincias para ambos años.")
        rps_nacionales = calcular_rps_nacionales(tasas_2017)
        print(
            "\nLos libros de 2017 y 2023 ya contienen tasas por 10.000 habitantes; "
            "sin los recuentos y la población originales no se puede verificar "
            "la normalización de 2017 desde su origen."
        )
        resultados = clasificar_provincias(
            tasas_2017, pesos, PERFILES_NACIONALES,
            rps_nacionales, PRIORIDADES_RP,
        )
        resultados_2023 = clasificar_provincias(
            tasas_2023, pesos, PERFILES_NACIONALES,
            rps_nacionales, PRIORIDADES_RP,
        )
        for anio, clasificacion in (("2017", resultados), ("2023", resultados_2023)):
            print(f"\nClasificación de {anio}, ordenada de mayor a menor seguridad:")
            print(f"{'Puesto':>6} {'Provincia':16} {'Prioridad':>12} {'Seguridad':>12}")
            for puesto, fila in enumerate(clasificacion, start=1):
                print(
                    f"{puesto:6} {fila['provincia']:16} "
                    f"{fila['prioridad_global']:12.4f} {fila['nivel_seguridad']:>12}"
                )
        mostrar_transiciones(resultados, resultados_2023)
        ruta_mapa = Path(opciones.geojson) if opciones.geojson else MAPA_ESPANA
        if ruta_mapa:
            faltan_2017 = guardar_mapa(
                carpeta_salida / "mapa_espana_2017.html", resultados, ruta_mapa,
                permitir_faltantes=True,
            )
            faltan_2023 = guardar_mapa(
                carpeta_salida / "mapa_espana_2023.html", resultados_2023, ruta_mapa,
                permitir_faltantes=True,
            )
            if faltan_2017 or faltan_2023:
                print("Aviso de mapa incompleto: faltan geometrías para "
                      + ", ".join(faltan_2017 or faltan_2023) + ".")
        print("\nCalibración nacional de 2017 conservada para ambos años; tasas truncadas en RP1–RP6.")
        return
    print("\nClasificación, ordenada de mayor a menor seguridad:")
    print(f"{'Puesto':>6} {'Provincia':16} {'Prioridad':>12} {'Seguridad':>12}")
    for puesto, fila in enumerate(resultados, start=1):
        print(
            f"{puesto:6} {fila['provincia']:16} "
            f"{fila['prioridad_global']:12.4f} {fila['nivel_seguridad']:>12}"
        )
    if opciones.ejercicio == "2":
        print("\nTasas de delitos por cada 10.000 habitantes:")
        print(f"{'Provincia':16}", end="")
        for criterio in CRITERIOS:
            print(f"{NOMBRES_CRITERIOS[criterio]:>17}", end="")
        print()
        for fila in resultados:
            print(f"{fila['provincia']:16}", end="")
            for criterio in CRITERIOS:
                print(f"{fila[f'tasa_{criterio}']:17.3f}", end="")
            print()
        print("\nPrioridades locales por provincia:")
        print(f"{'Provincia':16}", end="")
        for criterio in CRITERIOS:
            print(f"{NOMBRES_CRITERIOS[criterio]:>17}", end="")
        print()
        for fila in resultados:
            print(f"{fila['provincia']:16}", end="")
            for criterio in CRITERIOS:
                print(f"{fila[f'prioridad_{criterio}']:17.3f}", end="")
            print()
    ruta_mapa = Path(opciones.geojson) if opciones.geojson else (MAPA_ANDALUCIA if opciones.ejercicio in ("1", "2") else MAPA_ESPANA)
    if ruta_mapa:
        nombre_mapa = "mapa_espana_2017.html" if opciones.ejercicio == "3" else f"mapa_andalucia_2017_ejercicio_{opciones.ejercicio}.html"
        geometrias_faltantes = guardar_mapa(
            carpeta_salida / nombre_mapa, resultados, ruta_mapa,
            permitir_faltantes=opciones.ejercicio == "3",
        )
        print(f"Mapa guardado en {carpeta_salida / nombre_mapa}")
        if geometrias_faltantes:
            print("Aviso: el shapefile no tiene geometría para " + ", ".join(geometrias_faltantes) + ".")
    elif opciones.ejercicio == "3":
        print("Para generar el mapa, añade --geojson con las geometrías provinciales de España.")
    if opciones.ejercicio == "1":
        perfiles = PERFILES_ANDALUCIA
        puntos_perfil = PUNTOS_EJERCICIO_1
        prioridades_perfil = PRIORIDADES_EJERCICIO_1
    elif opciones.ejercicio == "2":
        perfiles = PERFILES_ANDALUCIA_NORMALIZADOS
        puntos_perfil = PUNTOS_ANDALUCIA
        prioridades_perfil = PRIORIDADES_ANDALUCIA
    else:
        perfiles = PERFILES_NACIONALES
        puntos_perfil = rps_nacionales
        prioridades_perfil = PRIORIDADES_RP
    prioridades_perfiles = {}
    for criterio in CRITERIOS:
        prioridades_perfiles[criterio] = calcular_prioridades_perfiles(
            criterio, perfiles, puntos_perfil, prioridades_perfil,
        )
    perfiles_globales = []
    for indice_perfil in range(3):
        prioridad = 0
        for indice, criterio in enumerate(CRITERIOS):
            prioridad += pesos[indice] * prioridades_perfiles[criterio][indice_perfil]
        perfiles_globales.append(prioridad)
    print("\nPrioridades locales y globales de los perfiles centrales:")
    print(f"{'Criterio':16} {'Alto':>9} {'Medio':>9} {'Bajo':>9} {'Peso AHP':>10}")
    for indice, criterio in enumerate(CRITERIOS):
        valores = prioridades_perfiles[criterio]
        print(
            f"{NOMBRES_CRITERIOS[criterio]:16} "
            f"{valores[0]:9.4f} {valores[1]:9.4f} {valores[2]:9.4f} {pesos[indice]:10.4f}"
        )
    print(
        f"{'Global':16} {perfiles_globales[0]:9.4f} "
        f"{perfiles_globales[1]:9.4f} {perfiles_globales[2]:9.4f}"
    )
    if not ruta_mapa and opciones.ejercicio in ("1", "2", "3"):
        print(
            "Mapa no generado: indica --geojson con geometrías provinciales "
            "para guardar el mapa HTML."
        )


def leer_opciones():
    analizador = argparse.ArgumentParser(description="Clasificación provincial de criminalidad con AHP-Sort.")
    analizador.add_argument("--ejercicio", choices=("1", "2", "3", "4"), help="Ejercicio que se va a ejecutar.")
    analizador.add_argument("--datos", help="Libro Excel de tasas normalizadas para el ejercicio 3.")
    analizador.add_argument("--datos-2017", help="Libro Excel de tasas normalizadas de 2017.")
    analizador.add_argument("--datos-2023", help="Libro Excel de tasas normalizadas de 2023.")
    analizador.add_argument("--geojson", help="GeoJSON o shapefile provincial para mapas.")
    analizador.add_argument("--salida", default="resultados", help="Carpeta de resultados (por defecto: resultados).")
    opciones = analizador.parse_args()
    opciones.menu_interactivo = not bool(opciones.ejercicio)
    return opciones


def mostrar_menu(opciones):
    ejercicios = (
        ("1", "Andalucía 2017, recuentos absolutos"),
        ("2", "Andalucía 2017, tasas por 10.000 habitantes"),
        ("3", "España 2017, tasas por 10.000 habitantes"),
        ("4", "Comparación de España 2017 y 2023"),
    )
    while True:
        print("\nPráctica 6 — AHPSort")
        for numero, nombre in ejercicios:
            print(f"  {numero}. {nombre}")
        print("  0. Salir")
        try:
            opcion = input("Elige un ejercicio (1-4) o 0 para salir: ").strip()
        except EOFError:
            return
        if opcion == "0":
            return
        if opcion not in {"1", "2", "3", "4"}:
            print("Opción no válida. Elige 1, 2, 3, 4 o 0.")
            continue
        opciones.ejercicio = opcion
        try:
            ejecutar_ejercicio(opciones)
        except (OSError, ValueError) as error:
            print(f"Error: {error}", file=sys.stderr)


if __name__ == "__main__":
    opciones = leer_opciones()
    if opciones.menu_interactivo:
        mostrar_menu(opciones)
    else:
        try:
            ejecutar_ejercicio(opciones)
        except (OSError, ValueError) as error:
            print(f"Error: {error}", file=sys.stderr)
            sys.exit(1)

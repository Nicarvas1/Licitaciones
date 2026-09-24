#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor SIN IA para ofertas de Mercado Publico (experimento).

Objetivo: medir cuanto se puede extraer solo con codigo y compararlo con los
resultados de la IA (extraccion_ia.json). No modifica checkpoints de la IA.

Tecnicas, de mas a menos confiable:
  1. Tablas con encabezados reconocibles (Excel, DOCX, CSV y PDF con tablas con
     bordes): se identifica que columna es descripcion, marca, modelo, cantidad,
     precio unitario y total, y se leen las filas.
  2. Filas "cuadradas" en texto libre: lineas donde aparecen cantidad, precio
     unitario y total con cantidad x unitario = total (tolerancia 1%). Sirve para
     tablas sin bordes y cotizaciones en texto.
  3. Catalogo de marcas y familias de modelos (HP ProBook 440 G11, Lenovo
     ThinkPad E14, Dell Latitude 3540, ...) para completar marca y modelo, tambien
     desde anexos tecnicos del mismo proveedor.

Los productos pasan por la MISMA validacion que usa 3_extraer_ia.py.

Salidas:
  - <licitacion>/extraccion_reglas.json
  - Excel con hojas Productos, Resumen, Archivos y (si hay resultados IA) Comparacion.

Uso:
  python 3_extraer_reglas.py --dir ".\\lotes\\2025-11\\ofertas" --metadata-csv ".\\lotes\\2025-11\\para_scrapear.csv"
"""

import argparse
import csv
import importlib.util
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent


def _cargar_extractor_ia():
    spec = importlib.util.spec_from_file_location("extractor_ia", RAIZ_PROYECTO / "3_extraer_ia.py")
    modulo = importlib.util.module_from_spec(spec)
    sys.path.insert(0, str(RAIZ_PROYECTO))
    spec.loader.exec_module(modulo)
    return modulo


ia = _cargar_extractor_ia()
pdfplumber = ia.pdfplumber
Document = ia.Document
load_workbook = ia.load_workbook
Workbook = ia.Workbook
Font = ia.Font
PatternFill = ia.PatternFill


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def norm(texto):
    """minusculas, sin tildes, espacios simples."""
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join(texto.lower().split())


def celda_texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return " ".join(str(valor).split())


# ---------------------------------------------------------------------------
# Catalogo de marcas y familias de modelos
# ---------------------------------------------------------------------------
# (patron de familia, marca, categoria). El modelo se arma con la familia y los
# tokens siguientes que contienen digitos (ej. "ProBook 440 G11").
FAMILIAS = [
    # HP
    (r"probook|elitebook|zbook|chromebook", "HP", "equipo"),
    (r"prodesk|elitedesk|eliteone|proone|elite\s*mini|pro\s*mini", "HP", "equipo"),
    (r"pavilion|victus|omen|envy|spectre", "HP", "equipo"),
    (r"laserjet|officejet|deskjet|smart\s*tank|designjet|pagewide|neverstop", "HP", "impresora"),
    # Lenovo
    (r"thinkpad|thinkbook|ideapad|legion|yoga", "Lenovo", "equipo"),
    (r"thinkcentre|ideacentre|thinkstation", "Lenovo", "equipo"),
    (r"thinkvision", "Lenovo", "monitor"),
    # Dell
    (r"latitude|vostro|inspiron|precision|xps|alienware", "Dell", "equipo"),
    (r"optiplex", "Dell", "equipo"),
    # Asus
    (r"vivobook|zenbook|expertbook|expertcenter|proart|\brog\b|\btuf\b", "Asus", "equipo"),
    # Acer
    (r"aspire|travelmate|veriton|extensa|swift|nitro|predator", "Acer", "equipo"),
    # Apple
    (r"macbook|imac|mac\s*mini|mac\s*studio", "Apple", "equipo"),
    # Impresoras
    (r"ecotank|workforce", "Epson", "impresora"),
    (r"imageclass|pixma|imagerunner|maxify|i-sensys", "Canon", "impresora"),
    (r"bizhub", "Konica Minolta", "impresora"),
    (r"ecosys|taskalfa", "Kyocera", "impresora"),
    (r"versalink|altalink|workcentre", "Xerox", "impresora"),
]
FAMILIAS = [(re.compile(rf"\b(?:{patron})\b", re.I), marca, categoria) for patron, marca, categoria in FAMILIAS]

MARCAS_PATRONES = [
    (r"\bhp\b|hewlett", "HP"), (r"\blenovo\b", "Lenovo"), (r"\bdell\b", "Dell"),
    (r"\basus\b", "Asus"), (r"\bacer\b", "Acer"), (r"\bapple\b", "Apple"),
    (r"\bsamsung\b", "Samsung"), (r"\blg\b", "LG"), (r"\bepson\b", "Epson"),
    (r"\bcanon\b", "Canon"), (r"\bbrother\b", "Brother"), (r"\blexmark\b", "Lexmark"),
    (r"\bxerox\b", "Xerox"), (r"\bricoh\b", "Ricoh"), (r"\bkyocera\b", "Kyocera"),
    (r"\bmicrosoft\b|\bsurface\b", "Microsoft"), (r"\bmsi\b", "MSI"), (r"\bhuawei\b", "Huawei"),
    (r"\bviewsonic\b", "ViewSonic"), (r"\baoc\b", "AOC"), (r"\bbenq\b", "BenQ"),
    (r"\bphilips\b", "Philips"), (r"\bgigabyte\b", "Gigabyte"), (r"\bpantum\b", "Pantum"),
    (r"\bkonica\b|\bminolta\b", "Konica Minolta"), (r"\bsharp\b", "Sharp"),
    (r"\btoshiba\b|\bdynabook\b", "Dynabook"),
]
MARCAS_PATRONES = [(re.compile(patron, re.I), marca) for patron, marca in MARCAS_PATRONES]

# Modelos sin familia con nombre: "HP 240 G9", "HP P24 G5", monitores Dell "P2422H".
MODELOS_SUELTOS = [
    (re.compile(r"\bHP\s+(\d{3}\s+G\d+)\b", re.I), "HP", "equipo"),
    (re.compile(r"\bHP\s+([PEVMZ]\d{2}[a-z]{0,2}(?:\s+G\d+)?)\b", re.I), "HP", "monitor"),
    (re.compile(r"\b([PSEU]\d{4}H[A-Z]?)\b"), "Dell", "monitor"),
    (re.compile(r"\bLenovo\s+(V\d{2}\s+G\d+(?:\s+[A-Z]{2,3})?)\b", re.I), "Lenovo", "equipo"),
]
TOKEN_MODELO = re.compile(r"^(?:[A-Za-z]{0,3}\d[\w\-]*|G\d+|Gen\s?\d+|Pro|Plus|Mini|SFF|Tower|Micro|AIO|Flip|x360)$", re.I)
# Especificaciones que NO son parte del modelo: i5, 16GB, 512GB, 2.4GHz, 23.8", 14"...
TOKEN_ESPECIFICACION = re.compile(r"^(?:[ir][3579](?:-\w+)?|ryzen|core|intel|amd|celeron|pentium|\d+(?:[.,]\d+)?(?:gb|tb|mb|ghz|mhz|w|hz|\"|''|pulgadas)|\d+[.,]\d+)$", re.I)


def detectar_modelo(texto):
    """Devuelve (marca, modelo, categoria) o (None, None, None)."""
    texto = celda_texto(texto)
    for patron, marca, categoria in FAMILIAS:
        coincidencia = patron.search(texto)
        if not coincidencia:
            continue
        # "Gen 5" se une en un solo token para no cortar el modelo.
        resto = re.sub(r"\bGen\s+(\d+)", r"Gen\1", texto[coincidencia.end():], flags=re.I)
        tokens = resto.split()
        extra = []
        for token in tokens[:5]:
            token = token.strip(",;:()")
            if TOKEN_ESPECIFICACION.match(token) or not TOKEN_MODELO.match(token):
                break
            extra.append(token)
        modelo = " ".join([coincidencia.group(0)] + extra)
        return marca, modelo, categoria
    for patron, marca, categoria in MODELOS_SUELTOS:
        coincidencia = patron.search(texto)
        if coincidencia:
            return marca, coincidencia.group(1), categoria
    return None, None, None


def detectar_marca(texto):
    for patron, marca in MARCAS_PATRONES:
        if patron.search(str(texto or "")):
            return marca
    return None


def modelo_tras_marca(texto):
    """Respaldo para marcas sin familias con nombre (LG 24MK430H, Samsung LS24C310):
    el primer codigo alfanumerico (letras y digitos) despues de la marca."""
    for patron, _ in MARCAS_PATRONES:
        coincidencia = patron.search(str(texto or ""))
        if not coincidencia:
            continue
        for token in str(texto)[coincidencia.end():].split()[:3]:
            token = token.strip(",;:()")
            if len(token) >= 4 and re.search(r"[A-Za-z]", token) and re.search(r"\d", token) \
                    and not TOKEN_ESPECIFICACION.match(token):
                return token
        return None
    return None


# ---------------------------------------------------------------------------
# Encabezados de tablas
# ---------------------------------------------------------------------------
PATRON_IVA_INCLUIDO = re.compile(r"con\s*iva|c/\s*iva|iva\s*incl|bruto")
PATRON_NETO = re.compile(r"neto|sin\s*iva|s/\s*iva")


def clasificar_encabezado(texto):
    """Devuelve (tipo, variante_iva) o None."""
    t = norm(texto)
    if not t or len(t) > 80:
        return None
    iva = "con_iva" if PATRON_IVA_INCLUIDO.search(t) else ("neto" if PATRON_NETO.search(t) else None)
    if re.search(r"unitari|\bunit\b|p\.?\s*unit|por\s+(equipo|unidad|item)|\bc/u\b", t):
        return "unitario", iva
    if re.search(r"\btotal|subtotal|\bmonto\s+(neto\s+)?ofertado|valor\s+(neto\s+)?ofertado", t):
        return "total", iva
    if re.search(r"^cant|cantidad|unidades|\bqty\b|n.?\s*de\s*(equipos|unidades)", t):
        return "cantidad", None
    if re.search(r"\bmarca\b", t) and re.search(r"\bmodelo\b", t):
        return "marca_modelo", None
    if re.search(r"\bmarca\b", t):
        return "marca", None
    if re.search(r"\bmodelo\b|part\s*number|numero\s+de\s+parte|n.?\s*de\s*parte|\bsku\b|\bp/?n\b", t):
        return "modelo", None
    if re.search(r"descrip|producto|detalle|\bbien|equipo|especificac|glosa|articulo|nombre|caracteristic", t):
        return "descripcion", None
    if re.fullmatch(r"(item|n|n.|no\.?|nro\.?|numero|linea|#)", t):
        return "item", None
    if re.search(r"precio|valor|monto|costo", t):
        return "precio", iva
    return None


TIPOS_IDENTIDAD = {"descripcion", "marca", "modelo", "marca_modelo"}
TIPOS_PRECIO = {"unitario", "total", "precio"}


def evaluar_encabezado(celdas):
    tipos = {}
    for indice, celda in enumerate(celdas):
        clasificacion = clasificar_encabezado(celda)
        if clasificacion:
            tipos.setdefault(clasificacion[0], []).append((indice, clasificacion[1], norm(celda)))
    valido = bool(TIPOS_IDENTIDAD & set(tipos)) and bool(TIPOS_PRECIO & set(tipos))
    return (len(tipos) if valido else 0), tipos


def elegir_columna(candidatos, preferir=None):
    """candidatos: [(indice, iva, texto)]. Prefiere neto; luego el patron dado; luego el primero."""
    if not candidatos:
        return None
    netos = [c for c in candidatos if c[1] == "neto"]
    if netos:
        candidatos = netos
    elif any(c[1] is None for c in candidatos):
        candidatos = [c for c in candidatos if c[1] is None] + [c for c in candidatos if c[1] is not None]
    if preferir:
        preferidos = [c for c in candidatos if re.search(preferir, c[2])]
        if preferidos:
            return preferidos[0][0]
    return candidatos[0][0]


def construir_mapeo(tipos):
    mapeo = {}
    # Formularios con "descripcion solicitada" y "descripcion ofertada": se usa la ofertada.
    descripciones = tipos.get("descripcion", [])
    if descripciones:
        ofertadas = [c for c in descripciones if re.search(r"ofert|propuest|proveedor", c[2])]
        mapeo["descripcion"] = (ofertadas or descripciones[-1:])[0][0]
    for tipo in ("marca", "modelo", "marca_modelo", "cantidad", "item"):
        if tipos.get(tipo):
            mapeo[tipo] = tipos[tipo][0][0]
    mapeo["unitario"] = elegir_columna(tipos.get("unitario", []))
    mapeo["total"] = elegir_columna(tipos.get("total", []), preferir=r"linea|item|equipo")
    if mapeo["unitario"] is None and tipos.get("precio"):
        mapeo["unitario"] = elegir_columna(tipos["precio"])  # "Precio" a secas = precio por unidad
    mapeo["con_iva"] = any(
        c[1] == "con_iva" for tipo in ("unitario", "total", "precio") for c in tipos.get(tipo, [])
        if c[0] in (mapeo.get("unitario"), mapeo.get("total"))
    )
    return {clave: valor for clave, valor in mapeo.items() if valor is not None}


def buscar_encabezado(filas, max_filas=40):
    """Busca la fila (o par de filas) de encabezado. Devuelve (inicio_datos, mapeo) o None."""
    mejor = None
    for indice in range(min(len(filas), max_filas)):
        opciones = [(evaluar_encabezado(filas[indice]), indice + 1)]
        if indice + 1 < len(filas) and not fila_tiene_monto(filas[indice + 1]):
            ancho = max(len(filas[indice]), len(filas[indice + 1]))
            combinada = [
                f"{celda_texto(filas[indice][c]) if c < len(filas[indice]) else ''} "
                f"{celda_texto(filas[indice + 1][c]) if c < len(filas[indice + 1]) else ''}"
                for c in range(ancho)
            ]
            opciones.append((evaluar_encabezado(combinada), indice + 2))
        for (puntaje, tipos), inicio in opciones:
            if puntaje and (mejor is None or puntaje > mejor[0]):
                mejor = (puntaje, inicio, tipos)
    if not mejor:
        return None
    return mejor[1], construir_mapeo(mejor[2])


# ---------------------------------------------------------------------------
# Numeros y filas
# ---------------------------------------------------------------------------
PATRON_NO_OFERTA = re.compile(
    r"presupuesto\s+(?:m[aá]ximo|disponible|referencial|estimado)|monto\s+(?:m[aá]ximo|disponible|referencial)|"
    r"disponibilidad\s+presupuestaria", re.I)
PRECIO_MINIMO_CLP = 10000  # ningun equipo del alcance cuesta menos; evita tomar n. de item como precio
PATRON_MONTO_CELDA = re.compile(r"^\s*(?:US\$|\$|CLP|USD)?\s*\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?\s*$|^\s*\$\s*\d+\s*$", re.I)
PATRON_RESUMEN = re.compile(
    r"^(sub\s*total|total|neto|iva|valor\s+total|monto\s+total|total\s+general|total\s+neto|"
    r"total\s+oferta|descuento|19\s*%)", re.I
)


PATRON_TOKEN_NUMERO = re.compile(r"\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?(?!\d)|\d+(?:[.,]\d{1,2})?(?!\d)")
# Palabras que pueden acompanar un monto sin cambiar su significado.
PATRON_PALABRAS_MONTO = re.compile(
    r"\b(clp|usd|us|uf|utm|pesos?|netos?|iva|c/u|cu|mas|incluido|incl|incluye|unitario|total|valor|precio|"
    r"aprox|sin|con|c/iva|s/iva|x)\b|[$+.\-/%()]", re.I)


def numero(valor):
    """Convierte una celda en numero. Aisla el monto de textos como '$ 650.000 c/u',
    '650.000 + IVA' o '650.000.-'. Devuelve None si la celda mezcla especificaciones
    (ej. 'Core i5 16GB') o trae mas de un numero."""
    if isinstance(valor, (int, float)):
        return valor
    texto = celda_texto(valor)
    if not texto or not re.search(r"\d", texto):
        return None
    t = norm(texto)
    cantidad = re.match(r"^\s*(\d+)\s*(unidades?|equipos?|u\.?|un\.?|uds?\.?|unid\.?)\s*$", t)
    if cantidad:
        return int(cantidad.group(1))
    tokens = PATRON_TOKEN_NUMERO.findall(t)
    resto = PATRON_PALABRAS_MONTO.sub(" ", PATRON_TOKEN_NUMERO.sub(" ", t))
    if len(tokens) != 1 or re.search(r"[a-z]", resto):
        return None
    return ia.normalizar_numero(tokens[0])


def fila_tiene_monto(fila):
    return any(
        (isinstance(c, (int, float)) and c >= 1000) or (isinstance(c, str) and PATRON_MONTO_CELDA.match(c))
        for c in fila
    )


def moneda_de(texto):
    t = norm(texto)
    if re.search(r"us\$|usd|dolar", t):
        return "USD"
    if re.search(r"\butm\b", t):
        return "UTM"
    if re.search(r"\buf\b", t):
        return "UF"
    return "CLP"


def es_fila_resumen(fila, mapeo):
    textos = [celda_texto(c) for c in fila if celda_texto(c)]
    if not textos:
        return False
    descripcion = celda_texto(fila[mapeo["descripcion"]]) if "descripcion" in mapeo and mapeo["descripcion"] < len(fila) else ""
    return bool(PATRON_RESUMEN.match(norm(descripcion or textos[0])))


def valor_columna(fila, mapeo, tipo):
    indice = mapeo.get(tipo)
    if indice is None or indice >= len(fila):
        return None
    return fila[indice]


def ubicacion_base(contexto):
    return {clave: contexto[clave] for clave in ("tipo_fuente", "pagina", "hoja", "tabla") if contexto.get(clave) is not None}


def bbox_de(contexto, indice):
    bboxes = contexto.get("bboxes_filas") or []
    return list(bboxes[indice]) if indice < len(bboxes) and bboxes[indice] else None


def filas_desde_tabla(filas, mapeo, contexto):
    """Convierte las filas de una tabla ya mapeada en productos crudos.
    contexto["desfase"] = indice (0-based) de la primera fila de `filas` dentro de la grilla."""
    productos = []
    celdas = contexto.get("bboxes_celdas") or []
    fila_encabezado = contexto.get("fila_encabezado")
    for numero_fila, fila in enumerate(filas, 1):
        indice_grilla = contexto["desfase"] + numero_fila - 1
        if not any(celda_texto(c) for c in fila):
            continue  # las plantillas separan productos con filas vacias: se sigue leyendo
        if not fila_tiene_monto(fila):
            # Plantillas con un bloque (y un encabezado) por producto: se adopta el nuevo encabezado.
            puntaje, tipos = evaluar_encabezado(fila)
            if puntaje >= 3:
                mapeo = construir_mapeo(tipos)
                fila_encabezado = indice_grilla
                continue
        if es_fila_resumen(fila, mapeo):
            continue
        if PATRON_NO_OFERTA.search(" ".join(celda_texto(c) for c in fila)):
            continue
        descripcion = celda_texto(valor_columna(fila, mapeo, "descripcion"))
        marca = celda_texto(valor_columna(fila, mapeo, "marca"))
        modelo = celda_texto(valor_columna(fila, mapeo, "modelo"))
        marca_modelo = celda_texto(valor_columna(fila, mapeo, "marca_modelo"))
        cantidad = numero(valor_columna(fila, mapeo, "cantidad"))
        unitario = numero(valor_columna(fila, mapeo, "unitario"))
        total = numero(valor_columna(fila, mapeo, "total"))
        if moneda_de(" ".join(celda_texto(c) for c in fila)) == "CLP":
            unitario = unitario if isinstance(unitario, (int, float)) and unitario >= PRECIO_MINIMO_CLP else None
            total = total if isinstance(total, (int, float)) and total >= PRECIO_MINIMO_CLP else None
        tiene_precio = any(isinstance(v, (int, float)) and v > 0 for v in (unitario, total))

        if not tiene_precio:
            # Descripcion partida en varias filas: se agrega a la fila anterior.
            if productos and descripcion and not cantidad and not celda_texto(valor_columna(fila, mapeo, "item")):
                productos[-1]["producto"] = f"{productos[-1]['producto']} {descripcion}".strip()
                productos[-1]["ubicacion"].setdefault("filas_extra", []).append(indice_grilla)
                if bbox_de(contexto, indice_grilla):
                    productos[-1]["ubicacion"].setdefault("bbox_extra", []).append(bbox_de(contexto, indice_grilla))
            continue
        if cantidad is not None and not (0 < cantidad <= 100000):
            cantidad = None
        nota_iva = None
        if cantidad and unitario and total:
            neto = cantidad * unitario
            if abs(total - neto * 1.19) <= max(2, total * 0.005):
                # unitario neto y total con IVA: se deja el total neto coherente
                total, nota_iva = round(neto, 2), "total_con_iva_convertido_a_neto"
            elif abs(total * 1.19 - neto) <= max(2, neto * 0.005):
                unitario, nota_iva = round(total / cantidad, 2), "unitario_con_iva_convertido_a_neto"
        total_calculado = False
        if cantidad and unitario and not total:
            total, total_calculado = round(cantidad * unitario, 2), True
        texto_identidad = " ".join(filter(None, [descripcion, marca_modelo, marca, modelo]))
        if not texto_identidad:
            continue
        ubicacion = {
            **ubicacion_base(contexto),
            "fila": indice_grilla,
            "fila_encabezado": fila_encabezado,
            "columnas": {tipo: mapeo[tipo] for tipo in ("descripcion", "marca", "modelo", "marca_modelo",
                                                         "cantidad", "unitario", "total") if tipo in mapeo},
        }
        if bbox_de(contexto, indice_grilla):
            ubicacion["bbox_fila"] = bbox_de(contexto, indice_grilla)
            if contexto.get("bbox_encabezado"):
                ubicacion["bbox_encabezado"] = contexto["bbox_encabezado"]
            if indice_grilla < len(celdas):
                ubicacion["bbox_celdas"] = {
                    tipo: list(celdas[indice_grilla][columna])
                    for tipo, columna in ubicacion["columnas"].items()
                    if columna < len(celdas[indice_grilla]) and celdas[indice_grilla][columna]
                }
        productos.append({
            "item": celda_texto(valor_columna(fila, mapeo, "item")) or None,
            "producto": texto_identidad,
            "marca": marca or None,
            "modelo": modelo or None,
            "cantidad": cantidad,
            "cantidad_fuente": "explicita" if cantidad is not None else None,
            "precio_unitario": unitario if unitario and unitario > 0 else None,
            "precio_total": total if total and total > 0 else None,
            "precio_total_tipo": "linea" if total else None,
            "total_calculado": total_calculado,
            "moneda": moneda_de(" ".join(celda_texto(c) for c in fila)),
            "pagina": contexto.get("pagina"),
            "fila_fuente": f"{contexto['ubicacion']} fila {contexto['desfase'] + numero_fila}",
            "evidencia": " | ".join(celda_texto(c) for c in fila if celda_texto(c))[:300],
            "confianza": "alta",
            "con_iva": mapeo.get("con_iva", False),
            "nota_iva": nota_iva,
            "ubicacion": ubicacion,
        })
    return productos


# Filas "cuadradas" en texto libre: cantidad x unitario = total.
PATRON_NUMERO_TEXTO = re.compile(r"(?<![\w.,])(?:US\$|\$)?\s?\d{1,3}(?:\.\d{3})+(?:,\d+)?(?![\w])|(?<![\w.,$])\d+(?![\w.,])")


def fila_cuadrada(linea):
    """Busca en una linea tres numeros seguidos con cantidad x unitario = total.
    Devuelve (descripcion, cantidad, unitario, total) o None."""
    numeros = []
    for coincidencia in PATRON_NUMERO_TEXTO.finditer(linea):
        valor = ia.normalizar_numero(coincidencia.group(0))
        if isinstance(valor, (int, float)):
            numeros.append((coincidencia.start(), valor))
    for i in range(len(numeros) - 2):
        (inicio_a, a), (_, b), (_, c) = numeros[i], numeros[i + 1], numeros[i + 2]
        for cantidad, unitario, inicio in ((a, b, inicio_a), (b, a, inicio_a)):
            if not (isinstance(cantidad, int) and 0 < cantidad <= 10000 and unitario >= 1000 and c >= unitario):
                continue
            if abs(cantidad * unitario - c) <= max(2, c * 0.01):
                descripcion = re.sub(r"^\s*\d{1,3}[.)\-]?\s+", "", linea[:inicio]).strip(" |:-\t")
                return descripcion, cantidad, unitario, c
    return None


def menciona_producto(texto):
    return bool(
        ia.PATRON_EQUIPO_RELEVANTE.search(texto) or ia.PATRON_MONITOR.search(texto)
        or ia.PATRON_IMPRESORA.search(texto) or detectar_modelo(texto)[1]
    )


PATRON_PALABRA_ENCABEZADO = re.compile(r"descrip|\bcant|precio|unitari|\btotal|\bitem|valor|detalle|monto|producto", re.I)


def es_linea_encabezado(linea):
    return len(set(m.group(0).lower() for m in PATRON_PALABRA_ENCABEZADO.finditer(linea))) >= 2


def bloque_descripcion(lineas, indice, bboxes, max_lineas=6):
    """Lineas de descripcion encima de una fila con montos: se detiene en un
    encabezado, en otra linea con montos, en un total o en un salto vertical grande."""
    bloque = []
    j = indice - 1
    while j >= 0 and len(bloque) < max_lineas:
        linea = lineas[j].strip()
        if not linea or fila_cuadrada(linea) or re.search(r"\$\s*\d", linea) \
                or es_linea_encabezado(linea) or PATRON_RESUMEN.match(norm(linea)):
            break
        if bboxes and j + 1 < len(bboxes) and bboxes[j] and bboxes[j + 1]:
            alto = max(1.0, bboxes[j][3] - bboxes[j][1])
            if bboxes[j + 1][1] - bboxes[j][3] > 1.2 * alto:  # parrafo separado
                break
        bloque.insert(0, linea)
        j -= 1
    return bloque, j + 1


def filas_desde_texto(lineas, contexto):
    productos = []
    bboxes = contexto.get("bboxes_filas") or []
    for numero_linea, linea in enumerate(lineas):
        resultado = fila_cuadrada(linea)
        if not resultado:
            continue
        descripcion, cantidad, unitario, total = resultado
        ubicacion = {**ubicacion_base(contexto), "fila": numero_linea, "linea_texto": True}
        if bbox_de(contexto, numero_linea):
            ubicacion["bbox_fila"] = bbox_de(contexto, numero_linea)
        # Descripcion en varias lineas (la linea con montos suele ser la ultima):
        # se antepone el bloque de lineas sin montos que esta justo encima.
        bloque, inicio_bloque = bloque_descripcion(lineas, numero_linea, bboxes)
        if bloque and (menciona_producto(" ".join(bloque)) or not menciona_producto(descripcion)):
            descripcion = " ".join(bloque + [descripcion]).strip()
            ubicacion["filas_extra"] = list(range(inicio_bloque, numero_linea))
            extra = [bbox_de(contexto, k) for k in range(inicio_bloque, numero_linea)]
            if all(extra):
                ubicacion["bbox_extra"] = extra
        if not re.search(r"[A-Za-z]{3,}", descripcion) or PATRON_RESUMEN.match(norm(descripcion)) \
                or es_linea_encabezado(descripcion) or PATRON_NO_OFERTA.search(f"{descripcion} {linea}"):
            continue
        productos.append({
            "item": None,
            "producto": descripcion,
            "marca": None,
            "modelo": None,
            "cantidad": cantidad,
            "cantidad_fuente": "explicita",
            "precio_unitario": unitario,
            "precio_total": total,
            "precio_total_tipo": "linea",
            "moneda": moneda_de(linea),
            "pagina": contexto.get("pagina"),
            "fila_fuente": f"{contexto['ubicacion']} linea {numero_linea + 1}",
            "evidencia": linea.strip()[:300],
            "confianza": "media",
            "con_iva": False,
            "ubicacion": ubicacion,
        })
    return productos


# ---------------------------------------------------------------------------
# Lectura por tipo de archivo
# ---------------------------------------------------------------------------
def extraer_de_grillas(grillas):
    """grillas: [(filas, contexto)]. Aplica mapeo de encabezados y, si no hay,
    filas cuadradas. Soporta tablas que continuan en la pagina siguiente."""
    productos = []
    metodos = set()
    mapeo_anterior = None
    for filas, contexto in grillas:
        filas = [list(fila) for fila in filas]
        encontrado = buscar_encabezado(filas)
        if encontrado:
            inicio, mapeo = encontrado
            mapeo_anterior = (mapeo, max(len(f) for f in filas))
            fila_encabezado = inicio - 1
            extra = {"desfase": inicio, "fila_encabezado": fila_encabezado}
            bboxes = contexto.get("bboxes_filas") or []
            if fila_encabezado < len(bboxes) and bboxes[fila_encabezado]:
                extra["bbox_encabezado"] = list(bboxes[fila_encabezado])
            productos += filas_desde_tabla(filas[inicio:], mapeo, {**contexto, **extra})
            metodos.add("tabla_encabezados")
            continue
        if mapeo_anterior and filas and max(len(f) for f in filas) == mapeo_anterior[1]:
            continuados = filas_desde_tabla(filas, mapeo_anterior[0], {**contexto, "desfase": 0})
            for producto in continuados:
                producto["ubicacion"]["continuacion"] = True
            productos += continuados
            metodos.add("tabla_encabezados")
            continue
        lineas = [" ".join(celda_texto(c) for c in fila if celda_texto(c)) for fila in filas]
        encontrados = filas_desde_texto(lineas, contexto)
        if encontrados:
            metodos.add("filas_cuadradas")
            productos += encontrados
    return productos, metodos


def leer_excel(path, max_filas=500, max_columnas=40):
    libro = load_workbook(path, read_only=True, data_only=True)
    grillas = []
    textos = []
    try:
        for hoja in libro.worksheets:
            filas = []
            for fila in hoja.iter_rows(max_row=max_filas, max_col=max_columnas, values_only=True):
                filas.append(list(fila))
            if filas:
                grillas.append((filas, {"ubicacion": f"hoja {hoja.title}", "tipo_fuente": "excel", "hoja": hoja.title}))
                textos += [" ".join(celda_texto(c) for c in fila if celda_texto(c)) for fila in filas]
    finally:
        libro.close()
    return grillas, "\n".join(textos)


def leer_docx(path):
    documento = Document(path)
    grillas = []
    for numero_tabla, tabla in enumerate(documento.tables, 1):
        filas = [[celda.text for celda in fila.cells] for fila in tabla.rows]
        grillas.append((filas, {"ubicacion": f"tabla {numero_tabla}", "tipo_fuente": "docx", "tabla": numero_tabla}))
    parrafos = [p.text for p in documento.paragraphs if p.text.strip()]
    grillas.append(([[p] for p in parrafos], {"ubicacion": "texto", "tipo_fuente": "docx_texto"}))
    texto = "\n".join(parrafos + [" ".join(f) for g, _ in grillas[:-1] for f in g])
    return grillas, texto


def leer_csv(path):
    contenido = None
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            contenido = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    if contenido is None:
        return [], ""
    try:
        dialecto = csv.Sniffer().sniff(contenido[:4096], delimiters=",;\t")
        filas = list(csv.reader(contenido.splitlines(), dialecto))
    except csv.Error:
        filas = [[linea] for linea in contenido.splitlines()]
    return [(filas, {"ubicacion": "csv", "tipo_fuente": "csv"})], contenido


def leer_pdf(path, max_paginas=20):
    """Tablas con bordes por pagina; si una pagina no tiene tabla util, su texto
    se revisa linea a linea buscando filas cuadradas."""
    grillas = []
    textos = []
    paginas_sin_texto = 0
    with pdfplumber.open(path) as pdf:
        for indice, pagina in enumerate(pdf.pages[:max_paginas]):
            texto = pagina.extract_text() or ""
            textos.append(texto)
            if len(texto.strip()) <= 40:
                paginas_sin_texto += 1
            try:
                tablas = [(t, t.extract() or []) for t in pagina.find_tables()]
            except Exception:
                tablas = []
            tablas = [(t, filas) for t, filas in tablas if ia.tabla_estructurada(filas)]
            for numero_tabla, (tabla, filas) in enumerate(tablas, 1):
                grillas.append((filas, {
                    "ubicacion": f"pagina {indice + 1} tabla {numero_tabla}", "tipo_fuente": "pdf",
                    "pagina": indice + 1, "tabla": numero_tabla,
                    "bboxes_filas": [fila.bbox for fila in tabla.rows],
                    "bboxes_celdas": [fila.cells for fila in tabla.rows],
                }))
            if not tablas and texto.strip():
                try:
                    lineas = pagina.extract_text_lines()
                except Exception:
                    lineas = [{"text": linea} for linea in texto.splitlines()]
                grillas.append(([[linea["text"]] for linea in lineas], {
                    "ubicacion": f"pagina {indice + 1} texto", "tipo_fuente": "pdf_texto", "pagina": indice + 1,
                    "bboxes_filas": [
                        (linea["x0"], linea["top"], linea["x1"], linea["bottom"]) if "x0" in linea else None
                        for linea in lineas
                    ],
                }))
        total_paginas = min(len(pdf.pages), max_paginas)
    escaneado = total_paginas > 0 and paginas_sin_texto == total_paginas
    return grillas, "\n".join(textos), escaneado


def leer_archivo(path, args):
    """Devuelve (grillas, texto_completo, estado)."""
    extension = path.suffix.lower()
    try:
        if extension in (".xlsx", ".xlsm"):
            grillas, texto = leer_excel(path)
            return grillas, texto, "ok"
        if extension == ".docx":
            grillas, texto = leer_docx(path)
            return grillas, texto, "ok"
        if extension in (".csv", ".tsv", ".txt"):
            grillas, texto = leer_csv(path)
            return grillas, texto, "ok"
        if extension == ".pdf":
            grillas, texto, escaneado = leer_pdf(path, args.max_paginas)
            return grillas, texto, "escaneado_requiere_ocr_o_ia" if escaneado else "ok"
        if extension in ia.EXTENSIONES_IMAGEN:
            return [], "", "imagen_requiere_ia"
        return [], "", "no_soportado"
    except Exception as exc:
        return [], "", f"error_lectura: {exc}"


# ---------------------------------------------------------------------------
# Proveedor
# ---------------------------------------------------------------------------
def completar_identidad(producto):
    """Marca, modelo y categoria desde las columnas o la descripcion."""
    texto = " ".join(filter(None, [producto.get("producto"), producto.get("marca"), producto.get("modelo")]))
    marca_catalogo, modelo_catalogo, categoria = detectar_modelo(texto)
    marca = producto.get("marca") or marca_catalogo or detectar_marca(texto)
    if marca:
        marca = detectar_marca(marca) or marca_catalogo or marca
    producto["marca"] = marca
    producto["modelo"] = producto.get("modelo") or modelo_catalogo or modelo_tras_marca(texto)
    if categoria and not producto.get("categoria"):
        producto["categoria"] = categoria
    return producto


def categoria_propia(producto):
    """Categoria segun el texto de la propia fila (sin marca/modelo prestados)."""
    return ia.clasificar_producto({"producto": producto.get("producto"), "modelo": producto.get("modelo"),
                                   "categoria": producto.get("categoria")})


def completar_desde_otros_documentos(productos, textos_por_archivo):
    """Si a un EQUIPO le falta marca/modelo y en los documentos del proveedor aparece
    exactamente un modelo del catalogo de su misma categoria, se asigna.
    Nunca se asigna a filas que por su propio texto no son equipos del alcance
    (discos, UPS, despacho, licencias...): eso las hacia pasar el filtro."""
    modelos = {}
    for archivo, texto in textos_por_archivo.items():
        for linea in texto.splitlines():
            marca, modelo, categoria = detectar_modelo(linea)
            if modelo:
                clave = (marca, ia.texto_normalizado(modelo), categoria)
                modelos.setdefault(clave, (marca, modelo, categoria, archivo))
    for producto in productos:
        if producto.get("modelo") and producto.get("marca"):
            continue
        categoria = categoria_propia(producto)
        if not categoria:
            continue
        candidatos = [m for m in modelos.values() if m[2] == categoria]
        if len({(m[0], ia.texto_normalizado(m[1])) for m in candidatos}) == 1:
            marca, modelo, _, archivo = candidatos[0]
            producto["marca"] = producto.get("marca") or marca
            producto["modelo"] = producto.get("modelo") or modelo
            producto["fuente_producto"] = archivo
            producto["marca_modelo_desde_otro_documento"] = True
    return productos


def procesar_proveedor(carpeta, args, seven_zip):
    info = ia.cargar_json(carpeta / "oferta.json")
    proveedor = info.get("proveedor") or carpeta.name
    rut = info.get("rut") or carpeta.name.split("__", 1)[0]
    total_oferta = info.get("total") or info.get("total_oferta") or ""
    ia.descomprimir(carpeta, seven_zip)

    crudos = []
    textos = {}
    archivos = []
    for path in sorted(ia.listar_archivos_oferta(carpeta), key=ia.prioridad):
        relativo = str(path.relative_to(carpeta))
        grillas, texto, estado = leer_archivo(path, args)
        textos[relativo] = texto
        productos, metodos = extraer_de_grillas(grillas) if estado == "ok" else ([], set())
        tipo_doc = ia.tipo_documental(path.name)
        for producto in productos:
            normalizado = ia.normalizar_producto(producto, proveedor, rut, relativo, tipo_doc)
            if normalizado:
                normalizado["metodo"] = "filas_cuadradas" if producto["confianza"] == "media" else "tabla_encabezados"
                normalizado["con_iva"] = producto.get("con_iva", False)
                normalizado["nota_iva"] = producto.get("nota_iva")
                normalizado["total_calculado"] = producto.get("total_calculado", False)
                normalizado["ubicacion"] = producto.get("ubicacion", {})
                crudos.append(completar_identidad(normalizado))
        archivos.append({
            "archivo": relativo,
            "tipo": path.suffix.lower(),
            "tipo_documental": tipo_doc,
            "estado_lectura": estado,
            "metodos": sorted(metodos),
            "filas_extraidas": len(productos),
        })

    crudos = [p for p in crudos if categoria_propia(p)]  # alcance segun el texto propio de la fila
    crudos = completar_desde_otros_documentos(crudos, textos)
    relevantes = ia.filtrar_productos_relevantes([dict(p) for p in crudos])
    for producto in relevantes:
        producto["subcategoria"] = ia.subcategoria_producto(producto)
    productos = ia.validar_productos([dict(p) for p in relevantes], relevantes, total_oferta)
    ok = sum(p.get("estado_validacion") == "ok" for p in productos)
    problemas = sum(p.get("estado_validacion") in ("incompleto", "inconsistente") for p in productos)
    if productos and ok and not problemas:
        estado = "resuelto"
    elif productos:
        estado = "parcial"
    else:
        estado = "sin_resultado"
    return {
        "proveedor": proveedor,
        "rut": rut,
        "total_oferta": total_oferta,
        "estado_reglas": estado,
        "productos": productos,
        "resultados_archivos": archivos,
    }


# ---------------------------------------------------------------------------
# Comparacion con la IA
# ---------------------------------------------------------------------------
def precios(productos):
    return sorted(float(p["precio_unitario"]) for p in productos if isinstance(p.get("precio_unitario"), (int, float)))


def mismos_precios(a, b):
    return len(a) == len(b) and all(ia.numeros_iguales(x, y) for x, y in zip(a, b))


def comparar(resultado_reglas, resultado_ia):
    reglas = resultado_reglas.get("productos", []) if resultado_reglas else []
    productos_ia = []
    if resultado_ia:
        productos_ia = ia.validar_productos(
            [dict(p) for p in resultado_ia.get("productos", [])],
            [dict(p) for p in resultado_ia.get("productos_parciales", [])],
            resultado_ia.get("total_oferta"),
        )
    precios_reglas, precios_ia = precios(reglas), precios(productos_ia)
    if not precios_reglas and not precios_ia:
        coincidencia = "ninguno_tiene_precio"
    elif not precios_reglas:
        coincidencia = "solo_ia"
    elif not precios_ia:
        coincidencia = "solo_reglas"
    elif mismos_precios(precios_reglas, precios_ia):
        coincidencia = "mismos_precios"
    elif set(round(p) for p in precios_reglas) & set(round(p) for p in precios_ia):
        coincidencia = "precios_parcialmente_iguales"
    else:
        coincidencia = "precios_distintos"

    marcas_iguales = None
    if coincidencia == "mismos_precios":
        pares = zip(sorted(reglas, key=lambda p: p.get("precio_unitario") or 0),
                    sorted(productos_ia, key=lambda p: p.get("precio_unitario") or 0))
        marcas_iguales = all(
            (detectar_marca(a.get("marca")) or a.get("marca")) == (detectar_marca(b.get("marca")) or b.get("marca"))
            for a, b in pares
        )
    return {
        "coincidencia": coincidencia,
        "marcas_iguales": marcas_iguales,
        "precios_reglas": ", ".join(f"{p:,.0f}" for p in precios_reglas),
        "precios_ia": ", ".join(f"{p:,.0f}" for p in precios_ia),
        "productos_ia": len(productos_ia),
        "productos_ia_ok": sum(p.get("estado_validacion") == "ok" for p in productos_ia),
    }


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def escribir_hoja(libro, nombre, columnas, filas):
    hoja = libro.create_sheet(nombre)
    for numero_columna, columna in enumerate(columnas, 1):
        celda = hoja.cell(1, numero_columna, columna)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    for numero_fila, fila in enumerate(filas, 2):
        for numero_columna, columna in enumerate(columnas, 1):
            valor = fila.get(columna)
            if isinstance(valor, (list, dict)):
                valor = ", ".join(map(str, valor)) if isinstance(valor, list) else json.dumps(valor, ensure_ascii=False)
            hoja.cell(numero_fila, numero_columna, valor)
    hoja.freeze_panes = "A2"
    if filas:
        hoja.auto_filter.ref = hoja.dimensions


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Extraccion de ofertas solo con reglas (sin IA).")
    parser.add_argument("--dir", default="ofertas")
    parser.add_argument("--metadata-csv", default=str(RAIZ_PROYECTO / "para_scrapear.csv"))
    parser.add_argument("--excel", help="Excel de salida (por defecto <dir>/resultado_reglas.xlsx)")
    parser.add_argument("--limite-licitaciones", type=int)
    parser.add_argument("--proveedor")
    parser.add_argument("--max-paginas", type=int, default=20)
    parser.add_argument("--seven-zip")
    args = parser.parse_args()

    raiz = Path(args.dir)
    if not raiz.exists():
        sys.exit(f"No existe: {raiz}")
    licitaciones = ia.detectar_licitaciones(raiz)
    if args.limite_licitaciones:
        licitaciones = licitaciones[:args.limite_licitaciones]
    metadata = ia.cargar_metadata_licitaciones(args.metadata_csv)
    seven_zip = ia.encontrar_7zip(args.seven_zip)
    ruta_excel = Path(args.excel) if args.excel else raiz / "resultado_reglas.xlsx"

    filas_productos, filas_resumen, filas_archivos, filas_comparacion = [], [], [], []
    for licitacion in licitaciones:
        carpetas = sorted(c for c in licitacion.iterdir() if c.is_dir() and not c.name.startswith("_"))
        if args.proveedor:
            carpetas = [c for c in carpetas if c.name == args.proveedor]
        if not carpetas:
            continue
        resultados_ia = {}
        ruta_ia = licitacion / "extraccion_ia.json"
        if ruta_ia.is_file():
            try:
                resultados_ia = {str(r.get("rut") or "").strip(): r for r in json.loads(ruta_ia.read_text(encoding="utf-8"))}
            except (OSError, json.JSONDecodeError):
                resultados_ia = {}
        contexto = {
            "codigo": licitacion.name,
            "nombre_licitacion": metadata.get(licitacion.name, {}).get("nombre_licitacion", ""),
        }
        resultados = []
        for carpeta in carpetas:
            resultado = procesar_proveedor(carpeta, args, seven_zip)
            resultados.append(resultado)
            base = {**contexto, "proveedor": resultado["proveedor"], "rut": resultado["rut"]}
            for numero, producto in enumerate(resultado["productos"], 1):
                filas_productos.append({**contexto, **producto,
                                        "id_revision": ia.id_revision(licitacion.name, resultado["rut"], producto)})
            for archivo in resultado["resultados_archivos"]:
                filas_archivos.append({**base, **archivo})
            fila_resumen = {
                **base,
                "estado_reglas": resultado["estado_reglas"],
                "productos": len(resultado["productos"]),
                "productos_ok": sum(p.get("estado_validacion") == "ok" for p in resultado["productos"]),
            }
            filas_resumen.append(fila_resumen)
            if resultados_ia:
                comparacion = comparar(resultado, resultados_ia.get(str(resultado["rut"]).strip()))
                filas_comparacion.append({**base, "estado_reglas": resultado["estado_reglas"], **comparacion})
            print(f"  {licitacion.name} | {resultado['proveedor'][:34]:34} | {resultado['estado_reglas']:13} "
                  f"| productos={len(resultado['productos'])}")
        ia.escribir_json_atomico(licitacion / "extraccion_reglas.json", resultados)

    libro = Workbook()
    libro.remove(libro.active)
    escribir_hoja(libro, "Productos", [
        "id_revision", "codigo", "nombre_licitacion", "proveedor", "rut", "estado_validacion", "metodo", "producto",
        "marca", "modelo", "categoria", "subcategoria", "cantidad", "precio_unitario", "precio_total", "moneda",
        "con_iva", "nota_iva", "total_calculado", "alertas", "archivo_fuente", "fuente_producto", "marca_modelo_desde_otro_documento",
        "pagina", "fila_fuente", "evidencia",
    ], filas_productos)
    escribir_hoja(libro, "Resumen", [
        "codigo", "nombre_licitacion", "proveedor", "rut", "estado_reglas", "productos", "productos_ok",
    ], filas_resumen)
    escribir_hoja(libro, "Archivos", [
        "codigo", "proveedor", "rut", "archivo", "tipo", "tipo_documental", "estado_lectura",
        "metodos", "filas_extraidas",
    ], filas_archivos)
    if filas_comparacion:
        escribir_hoja(libro, "Comparacion", [
            "codigo", "nombre_licitacion", "proveedor", "rut", "estado_reglas", "coincidencia",
            "marcas_iguales", "precios_reglas", "precios_ia", "productos_ia", "productos_ia_ok",
        ], filas_comparacion)
    ruta_excel.parent.mkdir(parents=True, exist_ok=True)
    libro.save(ruta_excel)

    # ---------------- Resumen en consola ----------------
    total = len(filas_resumen)

    def pct(n, d):
        return f"{n} ({(100 * n / d):.0f}%)" if d else "0"

    estados = Counter(f["estado_reglas"] for f in filas_resumen)
    print("\n" + "=" * 74)
    print("EXTRACCION SOLO CON REGLAS")
    print(f"Proveedores:                   {total}")
    print(f"  resueltos (todo valida ok):  {pct(estados['resuelto'], total)}")
    print(f"  parciales (algo encontrado): {pct(estados['parcial'], total)}")
    print(f"  sin resultado:               {pct(estados['sin_resultado'], total)}")
    lecturas = Counter(a["estado_lectura"].split(":")[0] for a in filas_archivos)
    metodos = Counter(m for a in filas_archivos for m in a["metodos"])
    print(f"Archivos: {len(filas_archivos)} | " + ", ".join(f"{k}={v}" for k, v in lecturas.most_common()))
    print("Metodos que produjeron filas: " + (", ".join(f"{k}={v}" for k, v in metodos.most_common()) or "ninguno"))
    if filas_comparacion:
        coincidencias = Counter(f["coincidencia"] for f in filas_comparacion)
        ambos = sum(coincidencias[k] for k in ("mismos_precios", "precios_parcialmente_iguales", "precios_distintos"))
        print("\nCOMPARACION CON LA IA (la IA no es verdad absoluta: revisar diferencias)")
        for clave in ("mismos_precios", "precios_parcialmente_iguales", "precios_distintos",
                      "solo_ia", "solo_reglas", "ninguno_tiene_precio"):
            print(f"  {clave:30} {coincidencias[clave]}")
        if ambos:
            print(f"  Cuando ambos encuentran precios, coinciden en todos: {pct(coincidencias['mismos_precios'], ambos)}")
        resueltos = [f for f in filas_comparacion if f["estado_reglas"] == "resuelto"]
        if resueltos:
            iguales = sum(f["coincidencia"] == "mismos_precios" for f in resueltos)
            print(f"  De los 'resueltos' por reglas, mismos precios que la IA: {pct(iguales, len(resueltos))}")
    print(f"\nExcel: {ruta_excel}")
    print("=" * 74)


if __name__ == "__main__":
    main()

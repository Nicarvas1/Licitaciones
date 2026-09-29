#!/usr/bin/env python3
"""Prueba de extraccion directa: documentos completos -> una llamada a LM Studio.

Ejemplos:
  python prueba_extraccion_simple.py --archivo cotizacion.pdf --modelo "qwen/qwen3.5-9b"
  python prueba_extraccion_simple.py --carpeta "ofertas/CODIGO/RUT__PROVEEDOR" \
      --metadata-csv para_scrapear.csv --modelo "qwen/qwen3.5-9b"
"""

import argparse
import csv
import json
import os
import re
import time
from pathlib import Path

import pdfplumber
import requests
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


LMSTUDIO_URL = os.environ.get(
    "LMSTUDIO_URL", "http://127.0.0.1:1234/v1/chat/completions"
)
EXTENSIONES = {".pdf", ".docx", ".xlsx", ".xlsm", ".csv", ".tsv", ".txt"}

PROMPT = """Analiza los documentos completos de este proveedor.

Obtén todos los computadores, notebooks, desktops, estaciones de trabajo,
all-in-one, monitores e impresoras/multifuncionales ofertados.

Las tablas pueden estar desordenadas o no tener tabulaciones o separaciones
claras. Reconstruye correctamente cada fila usando sus encabezados, número de
ítem, descripción, cantidad, precio unitario y precio total.

Devuelve ÚNICAMENTE JSON válido con esta estructura:
{{
  "productos": [
    {{
      "producto": "descripción del producto ofertado",
      "marca": null,
      "cantidad": null,
      "precio_unitario": null,
      "precio_total": null,
      "fuente": "nombre exacto del archivo y página u hoja"
    }}
  ]
}}

Reglas:
- Devuelve una fila por producto ofertado y no omitas productos del alcance.
- No declares procesador, RAM, disco, sistema operativo, accesorios, licencias,
  garantías, servicios, despacho, televisores, tablets, servidores ni redes como
  productos independientes.
- Si el mismo producto aparece en un documento técnico y uno económico, devuelve
  una sola fila: usa la descripción y marca técnicas, y los valores económicos.
- precio_unitario corresponde a una unidad y precio_total a esa línea.
- Si una fila tiene cantidad y total pero no unitario, calcula total / cantidad.
- No confundas IVA, subtotal o total general de la oferta con el precio de un producto.
- Si existen precios netos y con IVA, usa los netos.
- Devuelve cantidades y montos como números, sin símbolos ni separadores de miles.
- No inventes información. Si un dato no aparece, usa null.
- En fuente indica el nombre exacto del archivo y la página u hoja utilizada.

LICITACIÓN: {licitacion}
FECHA: {fecha}
PROVEEDOR: {proveedor}
TOTAL DE LA OFERTA EN EL PORTAL: {total_oferta}

DOCUMENTOS:
{documentos}
"""


def limpiar(valor):
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(valor or ""))


def leer_pdf(ruta, max_paginas):
    partes = []
    with pdfplumber.open(ruta) as pdf:
        for indice, pagina in enumerate(pdf.pages[:max_paginas], 1):
            partes.append(f"[PÁGINA {indice}]")
            try:
                tablas = pagina.find_tables()
            except Exception:
                tablas = []
            # Evita enviar dos veces las mismas filas: una como texto plano y otra
            # como tabla. Esta duplicación es una causa frecuente de productos repetidos.
            if tablas:
                cajas = [tabla.bbox for tabla in tablas]

                def fuera_de_tablas(objeto):
                    x0 = objeto.get("x0")
                    x1 = objeto.get("x1")
                    top = objeto.get("top")
                    bottom = objeto.get("bottom")
                    if None in (x0, x1, top, bottom):
                        return True
                    return not any(
                        x0 >= caja[0] and x1 <= caja[2] and top >= caja[1] and bottom <= caja[3]
                        for caja in cajas
                    )

                texto = pagina.filter(fuera_de_tablas).extract_text(layout=True) or ""
            else:
                texto = pagina.extract_text(layout=True) or pagina.extract_text() or ""
            if texto.strip():
                partes.append(limpiar(texto).strip())
            for numero, tabla_objeto in enumerate(tablas, 1):
                partes.append(f"[TABLA {numero} - PÁGINA {indice}]")
                for fila_numero, fila in enumerate(tabla_objeto.extract() or [], 1):
                    if not fila:
                        continue
                    valores = [limpiar(celda).replace("\n", " ").strip() for celda in fila]
                    if any(valores):
                        partes.append(f"FILA {fila_numero}: " + " || ".join(valores))
    return "\n".join(partes).strip()


def leer_docx(ruta):
    documento = Document(ruta)
    partes = [limpiar(p.text).strip() for p in documento.paragraphs if p.text.strip()]
    for numero, tabla in enumerate(documento.tables, 1):
        partes.append(f"[TABLA {numero}]")
        for fila_numero, fila in enumerate(tabla.rows, 1):
            valores = [limpiar(celda.text).replace("\n", " ").strip() for celda in fila.cells]
            if any(valores):
                partes.append(f"FILA {fila_numero}: " + " || ".join(valores))
    return "\n".join(partes).strip()


def leer_excel(ruta, max_filas, max_columnas):
    libro = load_workbook(
        ruta, read_only=True, data_only=True, keep_vba=ruta.suffix.lower() == ".xlsm"
    )
    partes = []
    try:
        for hoja in libro.worksheets:
            partes.append(f"[HOJA {hoja.title}]")
            for numero, fila in enumerate(hoja.iter_rows(values_only=True), 1):
                if numero > max_filas:
                    break
                valores = [limpiar(valor).strip() for valor in fila[:max_columnas]]
                if any(valores):
                    partes.append(f"FILA {numero}: " + " || ".join(valores))
    finally:
        libro.close()
    return "\n".join(partes).strip()


def leer_texto(ruta):
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            return limpiar(ruta.read_text(encoding=encoding)).strip()
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"No se pudo leer {ruta}")


def leer_archivo(ruta, args):
    extension = ruta.suffix.lower()
    if extension == ".pdf":
        return leer_pdf(ruta, args.max_paginas)
    if extension == ".docx":
        return leer_docx(ruta)
    if extension in {".xlsx", ".xlsm"}:
        return leer_excel(ruta, args.max_filas_excel, args.max_columnas_excel)
    if extension in {".csv", ".tsv", ".txt"}:
        return leer_texto(ruta)
    return ""


def obtener_archivos(args):
    archivos = [Path(ruta).resolve() for ruta in (args.archivo or [])]
    if args.carpeta:
        carpeta = Path(args.carpeta).resolve()
        if not carpeta.is_dir():
            raise SystemExit(f"No existe la carpeta: {carpeta}")
        archivos.extend(
            ruta for ruta in sorted(carpeta.iterdir())
            if ruta.is_file() and ruta.suffix.lower() in EXTENSIONES
        )
    unicos = []
    for archivo in archivos:
        if archivo not in unicos:
            unicos.append(archivo)
    faltantes = [str(ruta) for ruta in unicos if not ruta.is_file()]
    if faltantes:
        raise SystemExit(f"No existen estos archivos: {', '.join(faltantes)}")
    if not unicos:
        raise SystemExit("Indica --archivo o --carpeta con documentos compatibles.")
    return unicos


def contexto_oferta(args):
    licitacion = args.licitacion or "no informada"
    fecha = args.fecha or "no informada"
    proveedor = args.proveedor or "no informado"
    total = None
    if args.carpeta:
        carpeta = Path(args.carpeta).resolve()
        oferta = carpeta / "oferta.json"
        if oferta.is_file():
            datos = json.loads(oferta.read_text(encoding="utf-8"))
            proveedor = args.proveedor or datos.get("proveedor") or carpeta.name
            total = datos.get("total")
        if not args.licitacion:
            licitacion = carpeta.parent.name
    if args.metadata_csv and Path(args.metadata_csv).is_file():
        with Path(args.metadata_csv).open(encoding="utf-8-sig", newline="") as archivo:
            for fila in csv.DictReader(archivo):
                if (fila.get("codigo") or "").strip() == licitacion:
                    fecha = args.fecha or fila.get("fecha_publicacion") or fecha
                    break
    return licitacion, fecha, proveedor, total


def parsear_json(texto):
    limpio = texto.strip()
    limpio = re.sub(r"^```(?:json)?\s*", "", limpio, flags=re.I)
    limpio = re.sub(r"\s*```$", "", limpio)
    try:
        return json.loads(limpio)
    except json.JSONDecodeError:
        inicio = limpio.find("{")
        if inicio >= 0:
            objeto, _ = json.JSONDecoder().raw_decode(limpio[inicio:])
            return objeto
        raise


def consultar_lmstudio(prompt, args):
    payload = {
        "model": args.modelo,
        "messages": [
            {"role": "system", "content": "Devuelve exclusivamente JSON válido."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.0,
        "max_tokens": args.max_tokens,
        "stream": False,
        "chat_template_kwargs": {"enable_thinking": False},
    }
    inicio = time.perf_counter()
    respuesta = requests.post(args.url, json=payload, timeout=args.timeout)
    respuesta.raise_for_status()
    datos = respuesta.json()
    contenido = ((datos.get("choices") or [{}])[0].get("message") or {}).get("content", "")
    return parsear_json(contenido), datos.get("usage") or {}, round(time.perf_counter() - inicio, 1)


def normalizar_numero(valor):
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return valor
    texto = str(valor).strip().replace("$", "").replace(" ", "")
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", texto):
        texto = texto.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:,\d{3})+", texto):
        texto = texto.replace(",", "")
    else:
        texto = texto.replace(",", ".")
    texto = re.sub(r"[^\d.]", "", texto)
    try:
        numero = float(texto)
        return int(numero) if numero.is_integer() else numero
    except ValueError:
        return None


def guardar_excel(ruta, productos, contexto):
    licitacion, fecha, proveedor, _ = contexto
    columnas = [
        "Licitación", "Fecha licitación", "Proveedor", "Producto ofertado",
        "Marca", "Cantidad", "Precio unitario", "Precio total", "Fuente",
    ]
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Productos"
    for columna, titulo in enumerate(columnas, 1):
        celda = hoja.cell(1, columna, titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    for fila, producto in enumerate(productos, 2):
        valores = [
            licitacion, fecha, proveedor, producto.get("producto"), producto.get("marca"),
            normalizar_numero(producto.get("cantidad")),
            normalizar_numero(producto.get("precio_unitario")),
            normalizar_numero(producto.get("precio_total")), producto.get("fuente"),
        ]
        for columna, valor in enumerate(valores, 1):
            hoja.cell(fila, columna, valor)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    hoja.column_dimensions["D"].width = 55
    hoja.column_dimensions["I"].width = 55
    libro.save(ruta)


def main():
    parser = argparse.ArgumentParser(
        description="Prueba simple: documentos completos, una llamada a LM Studio y un Excel corto."
    )
    entrada = parser.add_mutually_exclusive_group(required=True)
    entrada.add_argument("--archivo", nargs="+", help="Uno o más documentos del mismo proveedor")
    entrada.add_argument("--carpeta", help="Carpeta de un proveedor que contiene oferta.json y anexos")
    parser.add_argument("--modelo", required=True, help="Identificador exacto del modelo cargado en LM Studio")
    parser.add_argument("--url", default=LMSTUDIO_URL)
    parser.add_argument("--licitacion")
    parser.add_argument("--fecha")
    parser.add_argument("--proveedor")
    parser.add_argument("--metadata-csv", default="para_scrapear.csv")
    parser.add_argument("--excel", default="prueba_extraccion_simple.xlsx")
    parser.add_argument("--json", dest="salida_json", default="prueba_extraccion_simple.json")
    parser.add_argument("--max-caracteres", type=int, default=100000)
    parser.add_argument("--max-paginas", type=int, default=50)
    parser.add_argument("--max-filas-excel", type=int, default=1000)
    parser.add_argument("--max-columnas-excel", type=int, default=50)
    parser.add_argument("--max-tokens", type=int, default=4096)
    parser.add_argument("--timeout", type=int, default=600)
    args = parser.parse_args()

    archivos = obtener_archivos(args)
    bloques = []
    for archivo in archivos:
        print(f"Leyendo: {archivo.name}")
        texto = leer_archivo(archivo, args)
        if texto:
            bloques.append(f"\n===== ARCHIVO: {archivo.name} =====\n{texto}")
        else:
            print(f"  Aviso: no se obtuvo texto de {archivo.name}")
    documentos = "\n".join(bloques).strip()
    if not documentos:
        raise SystemExit("No se obtuvo texto de ningún documento. Puede tratarse de un PDF escaneado.")
    if len(documentos) > args.max_caracteres:
        raise SystemExit(
            f"Los documentos suman {len(documentos):,} caracteres y exceden el límite de "
            f"{args.max_caracteres:,}. Esta prueba no fragmenta ni recorta información."
        )

    contexto = contexto_oferta(args)
    prompt = PROMPT.format(
        licitacion=contexto[0], fecha=contexto[1], proveedor=contexto[2],
        total_oferta=contexto[3] or "no informado", documentos=documentos,
    )
    print(f"Enviando una sola llamada | archivos={len(bloques)} | caracteres={len(prompt):,}")
    resultado, uso, segundos = consultar_lmstudio(prompt, args)
    productos = resultado.get("productos", []) if isinstance(resultado, dict) else []
    productos = [producto for producto in productos if isinstance(producto, dict)]

    salida = {
        "licitacion": contexto[0], "fecha_licitacion": contexto[1],
        "proveedor": contexto[2], "productos": productos,
        "diagnostico": {"segundos": segundos, "uso": uso, "archivos": [p.name for p in archivos]},
    }
    Path(args.salida_json).write_text(json.dumps(salida, ensure_ascii=False, indent=2), encoding="utf-8")
    guardar_excel(Path(args.excel), productos, contexto)
    print(f"Listo | productos={len(productos)} | tiempo={segundos}s")
    print(f"JSON: {Path(args.salida_json).resolve()}")
    print(f"Excel: {Path(args.excel).resolve()}")


if __name__ == "__main__":
    main()

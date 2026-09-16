#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor de prueba OpenAI para ofertas de Mercado Publico.

Flujo:
  1. Extrae el texto de CADA archivo por separado.
    2. Hace UNA llamada pequena a OpenAI por archivo.
  3. Guarda todos los hallazgos parciales, sin early stopping.
    4. Opcionalmente consolida los hallazgos por proveedor.
    5. Genera un Excel con productos, resumen por licitacion y consumo de tokens.

Formatos: PDF, DOCX, XLSX, XLSM, TXT, CSV, TSV, ZIP, RAR y 7Z.
Los PDF sin texto quedan registrados en necesita_ocr.

Instalacion:
    pip install pymupdf requests python-docx openpyxl openai

Ejemplos:
    python 3_extraer_openai_prueba.py --dir ofertas/3572-22-LE25 --rehacer --debug
    python 3_extraer_openai_prueba.py --dir ofertas --limite-proveedores 0 --excel resultados_openai.xlsx
"""

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

try:
    import fitz
    import pdfplumber
    import requests
    from openai import OpenAI
    from docx import Document
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
except ImportError as exc:
    sys.exit(
        f"Falta dependencia: {exc}. Instala: "
        "pip install pymupdf pdfplumber requests python-docx openpyxl openai"
    )

def cargar_env_local():
    ruta = Path(__file__).resolve().with_name(".env")
    if not ruta.is_file():
        return
    for linea in ruta.read_text(encoding="utf-8-sig").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or "=" not in linea:
            continue
        clave, valor = linea.split("=", 1)
        clave = clave.strip()
        if clave:
            os.environ.setdefault(clave, valor.strip().strip("\"'"))


cargar_env_local()

OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
DEBUG = False

MARCAS = [
    "HP", "Lenovo", "Dell", "Asus", "Acer", "Apple", "Samsung", "Lexmark",
    "Epson", "Canon", "Brother", "MSI", "Huawei", "Kingston", "Logitech",
    "LG", "ThinkCentre", "ThinkPad", "ProOne", "ProDesk", "OptiPlex",
    "Latitude", "Pavilion"
]

PROMPT_ARCHIVO = r'''Eres un extractor de datos de ofertas de licitaciones publicas chilenas.
Analiza UN archivo de oferta. Puede ser economico o tecnico.

PROVEEDOR: {proveedor}
RUT: {rut}
ARCHIVO: {archivo}
TIPO DOCUMENTAL: {tipo_documental}
TOTAL MOSTRADO EN EL CUADRO DE OFERTAS: {total_oferta}

Devuelve SOLO JSON valido:
{{
  "productos": [
    {{
      "item": "numero o identificador de item, o null",
      "producto": "descripcion concreta del bien ofertado",
      "marca": "marca o null",
      "modelo": "modelo o null",
      "cantidad": null,
      "precio_unitario": null,
      "precio_total": null,
      "moneda": "CLP|USD|UTM|null",
      "confianza": "alta|media|baja"
    }}
  ],
  "observaciones": "motivo breve si no existe informacion suficiente, o null"
}}

REGLAS:
- Producto es obligatorio cuando el archivo identifica algun bien, incluso sin marca.
- Extrae una fila por item o producto.
- En archivos tecnicos conserva producto, marca y modelo aunque no haya precio.
- En archivos economicos conserva cantidad y precios aunque la descripcion sea generica.
- Busca expresamente campos como "precio por equipo", "monto por equipo",
    "precio unitario", "oferta por equipo", "precio total", "monto total" y
    "total equipos". Si aparecen precio unitario y total de la misma oferta,
    devuelve ambos aunque la cantidad no este escrita.
- Si existe cantidad y precio total de linea, calcula precio_unitario = total/cantidad.
- Si la cantidad no aparece pero precio_total/precio_unitario da una division
    entera positiva, deja que el programa infiera la cantidad.
- No uses IVA, subtotal ni total general como producto o precio unitario.
- No inventes. Usa null cuando el dato no aparece.
- Si no hay productos, devuelve {{"productos": [], "observaciones": "motivo"}}.
- El alcance comercial es SOLO: notebook/laptop/portatil, desktop/escritorio/PC,
    all-in-one/AIO, workstation, monitor, impresora/multifuncional y sus tintas o
    toner. Puedes incluir teclado, mouse u otro complemento SOLO si acompana a un
    equipo computacional relevante del mismo archivo/oferta.
- No devuelvas como productos independientes las especificaciones de un equipo:
    procesador, RAM, SSD, HDD, disco, puertos, conectividad, sistema operativo,
    fuente de poder, garantia o servicios. Tampoco devuelvas servidores, storage,
    switches, routers, redes, cables, racks ni servicios de instalacion.
- En cada producto agrega "categoria": "equipo|monitor|impresora|consumible|complemento".

PISTAS DE PRODUCTOS:
{pistas_producto}

PISTAS DE PRECIOS:
{pistas_precio}

TEXTO DEL ARCHIVO:
{texto}
'''

PROMPT_CONSOLIDAR = r'''Eres un reconciliador de resultados extraidos de documentos de una misma oferta.
Une registros que representan el mismo item o producto. Los datos tecnicos pueden
estar en un archivo y el precio/cantidad en otro.

PROVEEDOR: {proveedor}
RUT: {rut}
TOTAL DEL CUADRO DE OFERTAS: {total_oferta}

Devuelve SOLO JSON valido:
{{
  "productos": [
    {{
      "item": "item o null",
      "producto": "descripcion concreta",
      "marca": "marca o null",
      "modelo": "modelo o null",
      "cantidad": null,
      "precio_unitario": null,
      "precio_total": null,
      "moneda": "CLP|USD|UTM|null",
      "fuente_producto": "archivo o null",
      "fuente_precio": "archivo o null",
      "confianza": "alta|media|baja"
    }}
  ],
  "observaciones": "texto breve o null"
}}

REGLAS:
- Une solamente cuando item, orden, descripcion, cantidad o modelo permitan una correspondencia razonable.
- No confundas el total general del proveedor con un precio unitario.
- Si no puedes unir un precio con seguridad, conserva el producto con precio null.
- Si hay total de linea y cantidad, calcula precio_unitario = total/cantidad.
- La cantidad puede estar en el anexo tecnico, el precio unitario y total en el
    economico, y la marca/modelo en otro documento del mismo proveedor: combina
    esas fuentes cuando describan el mismo equipo.
- Distingue equipo principal de complementos. Conserva un complemento solo si
    tiene precio propio o ayuda a explicar el paquete; no lo mezcles con el precio
    del equipo si el documento no dice que esta incluido.
- Si el total y el precio unitario corresponden a la misma oferta, infiere una
    cantidad entera solo cuando la division sea consistente y marca su origen.
- No inventes productos ni precios.

HALLAZGOS PARCIALES:
{parciales}
'''

from prompts_extraccion import (
    PROMPT_ARCHIVO as PROMPT_ARCHIVO_COMPARTIDO,
    PROMPT_CONSOLIDAR as PROMPT_CONSOLIDAR_COMPARTIDO,
)

PROMPT_ARCHIVO = PROMPT_ARCHIVO_COMPARTIDO
PROMPT_CONSOLIDAR = PROMPT_CONSOLIDAR_COMPARTIDO

PATRON_PRODUCTO = re.compile(
    r"notebook|computador|laptop|all.?in.?one|\baio\b|desktop|monitor|impresora|"
    r"workstation|servidor|pc\b|equipo|modelo|procesador|ryzen|core\s*i[3579]|"
    r"\bram\b|ssd|hdd|pulgadas|\bgb\b", re.I
)
PATRON_PRECIO = re.compile(
    r"precio|valor|unitario|cantidad|subtotal|total|neto|iva|\$\s*[\d.]|"
    r"\b\d{1,3}(?:\.\d{3})+(?:,\d+)?\b", re.I
)


def limpiar_control(texto):
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(texto))


def cargar_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def cargar_metadata_licitaciones(ruta):
    ruta = Path(ruta)
    if not ruta.is_file():
        print(f"AVISO: no existe el CSV de metadatos: {ruta}")
        return {}
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            with ruta.open(encoding=encoding, newline="") as archivo:
                filas = list(csv.DictReader(archivo))
            break
        except UnicodeDecodeError:
            continue
    else:
        print(f"AVISO: no se pudo leer el CSV de metadatos: {ruta}")
        return {}
    return {
        (fila.get("codigo") or "").strip(): {
            "nombre_licitacion": (fila.get("nombre") or "").strip(),
            "fecha_publicacion": (fila.get("fecha_publicacion") or "").strip(),
            "estado_licitacion": (fila.get("estado") or "").strip(),
            "organismo": (fila.get("organismo") or "").strip()
        }
        for fila in filas
        if (fila.get("codigo") or "").strip()
    }


def extraer_pdf(path, max_paginas):
    # pdfplumber detecta tablas por coordenadas y evita que filas/columnas
    # se mezclen al aplanar el PDF a texto lineal (bug de get_text("text")).
    try:
        partes = []
        with pdfplumber.open(path) as pdf:
            for index, page in enumerate(pdf.pages):
                if index >= max_paginas:
                    break
                texto_pagina = (page.extract_text() or "").strip()
                if texto_pagina:
                    partes.append(texto_pagina)
                try:
                    tablas = page.find_tables()
                except Exception:
                    tablas = []
                for num_tabla, tabla in enumerate(tablas, 1):
                    filas = tabla.extract()
                    if not filas:
                        continue
                    partes.append(f"[TABLA {num_tabla} - pagina {index + 1}, una fila por producto]")
                    for num_fila, fila in enumerate(filas, 1):
                        valores = [(celda or "").strip().replace("\n", " ") for celda in fila]
                        if any(valores):
                            partes.append(f"FILA {num_fila}: " + " || ".join(valores))
        texto = limpiar_control("\n".join(partes)).strip()
        if len(texto) > 40:
            return texto, "texto"
    except Exception:
        texto = ""

    # Respaldo: pdfplumber no extrajo nada usable (ej. PDF escaneado sin capa de texto normal).
    try:
        doc = fitz.open(path)
        partes = []
        for index, page in enumerate(doc):
            if index >= max_paginas:
                break
            partes.append(page.get_text("text"))
        doc.close()
        texto = limpiar_control("\n".join(partes)).strip()
        return texto, "texto" if len(texto) > 40 else "necesita_ocr"
    except Exception as exc:
        return "", f"error_pdf: {exc}"


def extraer_docx(path):
    try:
        doc = Document(path)
        partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for numero_tabla, tabla in enumerate(doc.tables, 1):
            partes.append(f"[TABLA {numero_tabla}]")
            for numero_fila, fila in enumerate(tabla.rows, 1):
                valores = [limpiar_control(celda.text).strip().replace("\n", " | ") for celda in fila.cells]
                if any(valores):
                    partes.append(f"FILA {numero_fila}: " + " || ".join(valores))
        texto = "\n".join(partes).strip()
        return texto, "texto" if len(texto) > 20 else "vacio"
    except Exception as exc:
        return "", f"error_docx: {exc}"


def extraer_excel(path, max_filas, max_columnas):
    try:
        wb = load_workbook(
            path,
            data_only=True,
            read_only=True,
            keep_vba=path.suffix.lower() == ".xlsm"
        )
        partes = []
        for hoja in wb.worksheets:
            partes.append(f"[HOJA {hoja.title}]")
            for numero_fila, fila in enumerate(hoja.iter_rows(values_only=True), 1):
                if numero_fila > max_filas:
                    break
                valores = []
                for valor in fila[:max_columnas]:
                    valores.append("" if valor is None else limpiar_control(valor).strip())
                if any(valores):
                    partes.append(f"FILA {numero_fila}: " + " || ".join(valores))
        wb.close()
        texto = "\n".join(partes).strip()
        return texto, "texto" if len(texto) > 20 else "vacio"
    except Exception as exc:
        return "", f"error_excel: {exc}"


def extraer_texto_plano(path):
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            texto = limpiar_control(path.read_text(encoding=encoding)).strip()
            return texto, "texto" if len(texto) > 20 else "vacio"
        except Exception:
            pass
    return "", "error_texto"


def encontrar_7zip(ruta_explicitada=None):
    candidatos = [ruta_explicitada, shutil.which("7z"), shutil.which("7z.exe")]
    if os.name == "nt":
        candidatos.extend([
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe"
        ])
    return next((str(x) for x in candidatos if x and Path(x).exists()), None)


def descomprimir(carpeta, seven_zip):
    registros = []
    raiz_extraidos = carpeta / "_extraidos"
    archivos = list(carpeta.glob("*.zip")) + list(carpeta.glob("*.rar")) + list(carpeta.glob("*.7z"))
    for archivo in archivos:
        destino = raiz_extraidos / re.sub(r"[^\w.-]", "_", archivo.stem)
        marcador = destino / ".extraido_ok"
        if marcador.exists():
            registros.append({"archivo": archivo.name, "estado": "ya_extraido"})
            continue
        destino.mkdir(parents=True, exist_ok=True)
        try:
            if archivo.suffix.lower() == ".zip":
                with zipfile.ZipFile(archivo) as compacto:
                    compacto.extractall(destino)
            elif seven_zip:
                proceso = subprocess.run(
                    [seven_zip, "x", "-y", f"-o{destino}", str(archivo)],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if proceso.returncode != 0:
                    raise RuntimeError((proceso.stderr or proceso.stdout)[-500:])
            else:
                raise RuntimeError("7-Zip no encontrado")
            marcador.write_text("ok", encoding="utf-8")
            registros.append({"archivo": archivo.name, "estado": "ok"})
        except Exception as exc:
            registros.append({"archivo": archivo.name, "estado": f"error: {exc}"})
    return registros


def extraer_archivo(path, args):
    extension = path.suffix.lower()
    if extension == ".pdf":
        return extraer_pdf(path, args.max_paginas)
    if extension == ".docx":
        return extraer_docx(path)
    if extension in (".xlsx", ".xlsm"):
        return extraer_excel(path, args.max_filas_excel, args.max_columnas_excel)
    if extension in (".txt", ".csv", ".tsv"):
        return extraer_texto_plano(path)
    if extension == ".xls":
        return "", "xls_legacy_no_soportado"
    return "", "no_soportado"


def tipo_documental(nombre):
    texto = nombre.lower()
    if any(x in texto for x in ("econom", "cotiza", "presupuesto", "precio")):
        return "economico"
    if any(x in texto for x in ("tecnic", "ficha", "especific", "producto")):
        return "tecnico"
    return "otro"


def prioridad(path):
    tipo = tipo_documental(path.name)
    orden = {"economico": 0, "tecnico": 1, "otro": 2}[tipo]
    return orden, path.name.lower()


def pistas(texto):
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]
    productos = [linea for linea in lineas if PATRON_PRODUCTO.search(linea) and len(linea) < 350]
    precios = [linea for linea in lineas if PATRON_PRECIO.search(linea) and len(linea) < 350]
    return productos[:40], precios[:40]


def compactar_texto(texto, limite):
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]
    importantes = [linea for linea in lineas if PATRON_PRODUCTO.search(linea) or PATRON_PRECIO.search(linea)]
    seleccionadas = []
    vistas = set()
    for linea in importantes + lineas:
        if linea not in vistas:
            seleccionadas.append(linea)
            vistas.add(linea)
        if sum(len(x) + 1 for x in seleccionadas) >= limite:
            break
    return "\n".join(seleccionadas)[:limite]


def parsear_json(respuesta):
    texto = respuesta.strip()
    texto = re.sub(r"^```(?:json)?\s*|\s*```$", "", texto, flags=re.I | re.S)
    try:
        return json.loads(texto), "ok"
    except Exception:
        pass
    coincidencia = re.search(r"\{.*\}", texto, re.S)
    if not coincidencia:
        return None, "sin_json"
    reparado = re.sub(r",\s*([}\]])", r"\1", coincidencia.group(0))
    try:
        return json.loads(reparado), "ok_reparado"
    except Exception as exc:
        return None, f"json_invalido: {exc}"


def consultar_openai(prompt, modelo, timeout, num_ctx=None, max_tokens=2048):
    """Llama a OpenAI con JSON mode. La clave se lee de OPENAI_API_KEY."""
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        return None, "error_openai: Falta OPENAI_API_KEY en .env o en el entorno", "", {}
    try:
        cliente = OpenAI(api_key=api_key, timeout=timeout, max_retries=2)
        respuesta = cliente.chat.completions.create(
            model=modelo,
            messages=[
                {"role": "system", "content": "Devuelve exclusivamente un objeto JSON valido."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0,
            max_tokens=max_tokens
        )
        cruda = respuesta.choices[0].message.content or ""
        datos, estado = parsear_json(cruda)
        uso = getattr(respuesta, "usage", None)
        detalle_prompt = getattr(uso, "prompt_tokens_details", None)
        detalle_completion = getattr(uso, "completion_tokens_details", None)
        consumo = {
            "modelo": getattr(respuesta, "model", modelo),
            "prompt_tokens": getattr(uso, "prompt_tokens", None),
            "completion_tokens": getattr(uso, "completion_tokens", None),
            "total_tokens": getattr(uso, "total_tokens", None),
            "cached_tokens": getattr(detalle_prompt, "cached_tokens", None),
            "reasoning_tokens": getattr(detalle_completion, "reasoning_tokens", None)
        }
        return datos, estado, cruda, consumo
    except Exception as exc:
        return None, f"error_openai: {type(exc).__name__}: {exc}", "", {}

def normalizar_numero(valor):
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return valor
    texto = str(valor).strip().replace("$", "").replace(" ", "")
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", texto):
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto and "." not in texto:
        texto = texto.replace(",", ".")
    texto = re.sub(r"[^\d.]", "", texto)
    try:
        numero = float(texto)
        return int(numero) if numero.is_integer() else numero
    except Exception:
        return None


def normalizar_producto(producto, proveedor, rut, archivo, tipo_doc):
    descripcion = str(producto.get("producto") or "").strip()
    if not descripcion:
        return None
    cantidad = normalizar_numero(producto.get("cantidad"))
    precio_unitario = normalizar_numero(producto.get("precio_unitario"))
    precio_total = normalizar_numero(producto.get("precio_total"))
    cantidad_fuente = producto.get("cantidad_fuente")
    if precio_unitario is None and precio_total is not None and cantidad not in (None, 0):
        precio_unitario = round(float(precio_total) / float(cantidad), 2)
    if cantidad is None and precio_unitario not in (None, 0) and precio_total not in (None, 0):
        proporcion = float(precio_total) / float(precio_unitario)
        cantidad_entera = round(proporcion)
        if cantidad_entera > 0 and abs(proporcion - cantidad_entera) < 0.01:
            cantidad = cantidad_entera
            cantidad_fuente = "inferida_total_dividido_unitario"
    return {
        "proveedor": proveedor,
        "rut": rut,
        "item": producto.get("item"),
        "producto": descripcion,
        "marca": producto.get("marca"),
        "modelo": producto.get("modelo"),
        "categoria": producto.get("categoria"),
        "cantidad": cantidad,
        "cantidad_fuente": cantidad_fuente,
        "precio_unitario": precio_unitario,
        "precio_total": precio_total,
        "moneda": producto.get("moneda"),
        "confianza": producto.get("confianza"),
        "archivo_fuente": archivo,
        "tipo_documental": tipo_doc
    }


PATRON_EQUIPO_RELEVANTE = re.compile(
    r"notebook|laptop|port[aá]til|ultrabook|chromebook|all.?in.?one|\baio\b|"
    r"desktop|escritorio|\bpc\b|computador(?:a)?|workstation|thinkcentre|"
    r"thinkpad|prodesk|optiplex|latitude|pavilion", re.I
)
PATRON_MONITOR = re.compile(r"monitor|display|pantalla", re.I)
PATRON_IMPRESORA = re.compile(r"impresora|multifuncional|plotter", re.I)
PATRON_CONSUMIBLE = re.compile(r"tinta|toner|t[oó]ner|cartucho|tambor", re.I)
PATRON_COMPLEMENTO = re.compile(
    r"teclado|mouse|rat[oó]n|docking|dock|base de expansi[oó]n", re.I
)
PATRON_ESPECIFICACION = re.compile(
    r"procesador|cpu|\bram\b|memoria(?:\s+(?:ram|ddr))?|\bssd\b|\bhdd\b|"
    r"disco(?:\s+duro)?|puertos?|conectividad|sistema operativo|windows|"
    r"fuente de poder|garant[ií]a|servicio|instalaci[oó]n|configuraci[oó]n|"
    r"servidor|storage|switch|router|access point|\bred\b|cable|rack", re.I
)


def clasificar_producto(producto):
    texto = " ".join(str(producto.get(c) or "") for c in ("producto", "modelo", "categoria"))
    if PATRON_ESPECIFICACION.search(texto) and not any(
        patron.search(texto) for patron in (PATRON_EQUIPO_RELEVANTE, PATRON_MONITOR, PATRON_IMPRESORA)
    ):
        return None
    if PATRON_EQUIPO_RELEVANTE.search(texto):
        return "equipo"
    if PATRON_MONITOR.search(texto):
        return "monitor"
    if PATRON_IMPRESORA.search(texto):
        return "impresora"
    if PATRON_CONSUMIBLE.search(texto):
        return "consumible"
    if PATRON_COMPLEMENTO.search(texto):
        return "complemento"
    return None


def filtrar_productos_relevantes(productos):
    clasificados = []
    for producto in productos:
        categoria = clasificar_producto(producto)
        if categoria:
            producto["categoria"] = categoria
            clasificados.append(producto)
    tiene_principal = any(
        producto["categoria"] in {"equipo", "monitor", "impresora", "consumible"}
        for producto in clasificados
    )
    return [
        producto for producto in clasificados
        if producto["categoria"] != "complemento" or tiene_principal
    ]


def extraer_parciales_archivo(path, texto, proveedor, rut, total_oferta, args):
    productos_pista, precios_pista = pistas(texto)
    texto_reducido = compactar_texto(texto, args.max_chars_archivo)
    relativo = str(path)
    tipo_doc = tipo_documental(path.name)
    prompt = PROMPT_ARCHIVO.format(
        proveedor=proveedor,
        rut=rut,
        archivo=relativo,
        tipo_documental=tipo_doc,
        total_oferta=total_oferta or "no informado",
        pistas_producto="\n".join(productos_pista) or "(ninguna)",
        pistas_precio="\n".join(precios_pista) or "(ninguna)",
        texto=texto_reducido
    )
    datos, estado, cruda, consumo = consultar_openai(
        prompt, args.modelo, args.timeout, args.num_ctx, args.max_tokens
    )
    productos = []
    if isinstance(datos, dict) and isinstance(datos.get("productos"), list):
        for producto in datos["productos"]:
            if isinstance(producto, dict):
                limpio = normalizar_producto(producto, proveedor, rut, relativo, tipo_doc)
                if limpio:
                    productos.append(limpio)
    return productos, estado, cruda, prompt, datos.get("observaciones") if isinstance(datos, dict) else None, consumo


def consolidar_parciales(parciales, proveedor, rut, total_oferta, args):
    if not parciales or args.sin_consolidar:
        return parciales, "omitida", "", "", {}
    resumen = json.dumps(parciales, ensure_ascii=False, separators=(",", ":"))
    if len(resumen) > args.max_chars_consolidacion:
        resumen = resumen[:args.max_chars_consolidacion]
    prompt = PROMPT_CONSOLIDAR.format(
        proveedor=proveedor,
        rut=rut,
        total_oferta=total_oferta or "no informado",
        parciales=resumen
    )
    datos, estado, cruda, consumo = consultar_openai(
        prompt, args.modelo, args.timeout, args.num_ctx, args.max_tokens
    )
    consolidados = []
    if isinstance(datos, dict) and isinstance(datos.get("productos"), list):
        for producto in datos["productos"]:
            if not isinstance(producto, dict):
                continue
            limpio = normalizar_producto(producto, proveedor, rut, "consolidacion", "consolidado")
            if limpio:
                limpio["fuente_producto"] = producto.get("fuente_producto")
                limpio["fuente_precio"] = producto.get("fuente_precio")
                consolidados.append(limpio)
    return filtrar_productos_relevantes(consolidados or parciales), estado, cruda, prompt, consumo


def procesar_oferta(carpeta, args, seven_zip, solo_tecnicos=False):
    info = cargar_json(carpeta / "oferta.json")
    proveedor = info.get("proveedor") or carpeta.name
    rut = info.get("rut") or carpeta.name.split("__", 1)[0]
    total_oferta = info.get("total") or info.get("total_oferta") or ""
    archivos_comprimidos = descomprimir(carpeta, seven_zip)

    archivos = [
        path for path in carpeta.rglob("*")
        if path.is_file()
        and path.name not in {"oferta.json", "extraccion_ia.json", "extraccion_openai.json", ".extraido_ok"}
        and "_debug_ia" not in path.parts
    ]
    if solo_tecnicos:
        archivos = [
            path for path in archivos
            if tipo_documental(path.name) == "tecnico"
            or "tecnico" in str(path).lower()
        ]

    parciales = []
    resultados_archivos = []
    necesita_ocr = []
    no_soportados = []
    marcas_texto = set()

    debug_dir = carpeta / "_debug_ia"
    if args.debug:
        debug_dir.mkdir(exist_ok=True)

    for indice, path in enumerate(sorted(archivos, key=prioridad), 1):
        texto, estado_lectura = extraer_archivo(path, args)
        relativo = str(path.relative_to(carpeta))
        for marca in MARCAS:
            if re.search(rf"\b{re.escape(marca)}\b", texto, re.I):
                marcas_texto.add(marca)

        registro = {
            "archivo": relativo,
            "tipo": path.suffix.lower(),
            "tipo_documental": tipo_documental(path.name),
            "estado_lectura": estado_lectura,
            "caracteres": len(texto),
            "estado_ia": "no_ejecutada",
            "productos_encontrados": 0,
            "observaciones": None
        }

        if estado_lectura == "necesita_ocr":
            necesita_ocr.append(relativo)
        elif estado_lectura in ("no_soportado", "xls_legacy_no_soportado"):
            no_soportados.append(relativo)
        elif estado_lectura == "texto" and texto and not args.solo_texto:
            productos, estado_ia, cruda, prompt, observaciones, consumo = extraer_parciales_archivo(
                Path(relativo), texto, proveedor, rut, total_oferta, args
            )
            parciales.extend(productos)
            registro["estado_ia"] = estado_ia
            registro["productos_encontrados"] = len(productos)
            registro["observaciones"] = observaciones
            registro["consumo_openai"] = consumo
            if args.debug:
                base = f"{indice:02d}__{re.sub(r'[^\w.-]', '_', path.stem)[:70]}"
                (debug_dir / f"{base}__prompt.txt").write_text(prompt, encoding="utf-8")
                (debug_dir / f"{base}__respuesta.txt").write_text(cruda, encoding="utf-8")
        elif args.solo_texto and estado_lectura == "texto":
            registro["estado_ia"] = "solo_texto"

        resultados_archivos.append(registro)

    productos_finales = []
    estado_consolidacion = "no_ejecutada"
    consumo_consolidacion = {}
    if not args.solo_texto:
        (
            productos_finales,
            estado_consolidacion,
            cruda_consolidacion,
            prompt_consolidacion,
            consumo_consolidacion
        ) = consolidar_parciales(parciales, proveedor, rut, total_oferta, args)
        if args.debug and prompt_consolidacion:
            (debug_dir / "99__consolidacion__prompt.txt").write_text(prompt_consolidacion, encoding="utf-8")
            (debug_dir / "99__consolidacion__respuesta.txt").write_text(cruda_consolidacion, encoding="utf-8")

    return {
        "proveedor": proveedor,
        "rut": rut,
        "total_oferta": total_oferta,
        "productos": productos_finales,
        "productos_parciales": filtrar_productos_relevantes(parciales),
        "resultados_archivos": resultados_archivos,
        "estado_consolidacion": estado_consolidacion,
        "consumo_consolidacion": consumo_consolidacion,
        "archivos_comprimidos": archivos_comprimidos,
        "necesita_ocr": necesita_ocr,
        "no_soportados": no_soportados,
        "marcas_detectadas_texto": sorted(marcas_texto)
    }


def detectar_licitaciones(raiz):
    if any(hijo.is_dir() and (hijo / "oferta.json").exists() for hijo in raiz.iterdir()):
        return [raiz]
    return sorted((hijo for hijo in raiz.iterdir() if hijo.is_dir()), key=lambda path: path.name)


def registrar_resultado(codigo, metadata, resultado, productos, consumos, proveedores):
    contexto = {
        "codigo": codigo,
        "nombre_licitacion": metadata.get("nombre_licitacion", ""),
        "fecha_publicacion": metadata.get("fecha_publicacion", ""),
        "estado_licitacion": metadata.get("estado_licitacion", ""),
        "organismo": metadata.get("organismo", "")
    }
    for producto in filtrar_productos_relevantes(resultado.get("productos", [])):
        productos.append({**contexto, **producto})

    llamadas = 0
    errores_api = 0
    for archivo in resultado.get("resultados_archivos", []):
        estado = archivo.get("estado_ia", "no_ejecutada")
        if estado in ("no_ejecutada", "solo_texto"):
            continue
        llamadas += 1
        if str(estado).startswith("error") or str(estado).startswith("http_"):
            errores_api += 1
        consumo = archivo.get("consumo_openai") or {}
        consumos.append({
            **contexto,
            "proveedor": resultado.get("proveedor", ""),
            "rut": resultado.get("rut", ""),
            "tipo_llamada": "archivo",
            "archivo": archivo.get("archivo", ""),
            "estado_ia": estado,
            "productos_encontrados": archivo.get("productos_encontrados", 0),
            **consumo
        })

    estado_consolidacion = resultado.get("estado_consolidacion", "no_ejecutada")
    if estado_consolidacion not in ("no_ejecutada", "omitida"):
        llamadas += 1
        if str(estado_consolidacion).startswith("error") or str(estado_consolidacion).startswith("http_"):
            errores_api += 1
        consumos.append({
            **contexto,
            "proveedor": resultado.get("proveedor", ""),
            "rut": resultado.get("rut", ""),
            "tipo_llamada": "consolidacion",
            "archivo": "",
            "estado_ia": estado_consolidacion,
            "productos_encontrados": len(resultado.get("productos", [])),
            **(resultado.get("consumo_consolidacion") or {})
        })

    proveedores.append({
        **contexto,
        "proveedor": resultado.get("proveedor", ""),
        "rut": resultado.get("rut", ""),
        "productos": len(resultado.get("productos", [])),
        "llamadas_api": llamadas,
        "errores_api": errores_api,
        "archivos_ocr": len(resultado.get("necesita_ocr", [])),
        "archivos_no_soportados": len(resultado.get("no_soportados", []))
    })


def valor_token(valor):
    try:
        return int(valor or 0)
    except (TypeError, ValueError):
        return 0


def producto_tiene_precio(producto):
    return (
        producto.get("precio_unitario") not in (None, "")
        or producto.get("precio_total") not in (None, "")
    )


def generar_excel(productos, consumos, proveedores, ruta):
    """Genera el Excel y evita perder la corrida si el archivo esta abierto."""
    from datetime import datetime

    columnas_productos = [
        "codigo", "nombre_licitacion", "fecha_publicacion", "estado_licitacion", "organismo",
        "proveedor", "rut", "item", "producto", "categoria", "marca", "modelo",
        "cantidad", "cantidad_fuente", "precio_unitario", "precio_total", "moneda",
        "fuente_producto", "fuente_precio", "archivo_fuente", "confianza"
    ]
    columnas_consumo = [
        "codigo", "nombre_licitacion", "fecha_publicacion", "proveedor", "rut",
        "tipo_llamada", "archivo", "estado_ia", "productos_encontrados", "modelo",
        "prompt_tokens", "completion_tokens", "total_tokens", "cached_tokens", "reasoning_tokens"
    ]

    resumen_por_codigo = {}
    for proveedor in proveedores:
        codigo = proveedor["codigo"]
        resumen = resumen_por_codigo.setdefault(codigo, {
            "codigo": codigo,
            "nombre_licitacion": proveedor["nombre_licitacion"],
            "fecha_publicacion": proveedor["fecha_publicacion"],
            "estado_licitacion": proveedor["estado_licitacion"],
            "organismo": proveedor["organismo"],
            "proveedores_procesados": 0,
            "productos": 0,
            "llamadas_api": 0,
            "errores_api": 0,
            "archivos_ocr": 0,
            "archivos_no_soportados": 0,
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
            "cached_tokens": 0,
            "reasoning_tokens": 0
        })
        resumen["proveedores_procesados"] += 1
        for campo in ("productos", "llamadas_api", "errores_api", "archivos_ocr", "archivos_no_soportados"):
            resumen[campo] += proveedor[campo]
    for consumo in consumos:
        resumen = resumen_por_codigo.get(consumo["codigo"])
        if not resumen:
            continue
        for campo in ("prompt_tokens", "completion_tokens", "total_tokens", "cached_tokens", "reasoning_tokens"):
            resumen[campo] += valor_token(consumo.get(campo))

    resumenes = list(resumen_por_codigo.values())
    if resumenes:
        campos_sumables = [
            "proveedores_procesados", "productos", "llamadas_api", "errores_api",
            "archivos_ocr", "archivos_no_soportados", "prompt_tokens",
            "completion_tokens", "total_tokens", "cached_tokens", "reasoning_tokens"
        ]
        total = {campo: "" for campo in resumenes[0]}
        total.update({"codigo": "TOTAL", "nombre_licitacion": "Todas las licitaciones"})
        for campo in campos_sumables:
            total[campo] = sum(fila[campo] for fila in resumenes)
        resumenes.append(total)

    columnas_resumen = [
        "codigo", "nombre_licitacion", "fecha_publicacion", "estado_licitacion", "organismo",
        "proveedores_procesados", "productos", "llamadas_api", "errores_api",
        "archivos_ocr", "archivos_no_soportados", "prompt_tokens", "completion_tokens",
        "total_tokens", "cached_tokens", "reasoning_tokens"
    ]

    wb = Workbook()
    wb.remove(wb.active)

    def agregar_hoja(titulo, columnas, filas):
        ws = wb.create_sheet(titulo)
        for columna, nombre in enumerate(columnas, 1):
            celda = ws.cell(1, columna, nombre)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="1F4E78")
        for numero_fila, fila in enumerate(filas, 2):
            for columna, nombre in enumerate(columnas, 1):
                ws.cell(numero_fila, columna, fila.get(nombre))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    agregar_hoja("Productos", columnas_productos, productos)
    agregar_hoja("Resumen", columnas_resumen, resumenes)
    agregar_hoja("Consumo", columnas_consumo, consumos)

    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(ruta)
        return ruta
    except PermissionError:
        # En Windows normalmente significa que el XLSX esta abierto en Excel.
        sello = datetime.now().strftime("%Y%m%d_%H%M%S")
        alternativa = ruta.with_name(f"{ruta.stem}_{sello}{ruta.suffix}")
        wb.save(alternativa)
        print(f"AVISO: no se pudo sobrescribir {ruta.name}; probablemente esta abierto en Excel.")
        print(f"Se guardo una copia nueva en: {alternativa}")
        return alternativa


def main():
    global DEBUG
    parser = argparse.ArgumentParser(description="Extraccion por archivo y consolidacion ligera por proveedor")
    parser.add_argument("--dir", default="ofertas")
    parser.add_argument("--modelo", default="gpt-4.1-mini")
    parser.add_argument("--solo-texto", action="store_true")
    parser.add_argument("--sin-consolidar", action="store_true")
    parser.add_argument("--excel")
    parser.add_argument("--rehacer", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--seven-zip")
    parser.add_argument("--max-paginas", type=int, default=20)
    parser.add_argument("--max-filas-excel", type=int, default=500)
    parser.add_argument("--max-columnas-excel", type=int, default=40)
    parser.add_argument("--max-chars-archivo", type=int, default=7000)
    parser.add_argument("--max-chars-consolidacion", type=int, default=12000)
    parser.add_argument("--num-ctx", type=int, default=8192)
    parser.add_argument("--timeout", type=int, default=300)
    parser.add_argument("--max-tokens", type=int, default=2048)
    parser.add_argument("--metadata-csv", default=str(Path(__file__).resolve().with_name("para_scrapear.csv")),
                        help="CSV con codigo, nombre y fecha_publicacion")
    parser.add_argument("--proveedor", help="Texto contenido en nombre/RUT de un proveedor")
    parser.add_argument("--limite-proveedores", type=int, default=1,
                        help="Para esta prueba, procesa solo N proveedores (default 1)")
    parser.add_argument("--solo-tecnicos-fallback", action="store_true",
                        help="Procesa tecnicos y los combina con el resultado economico previo")
    args = parser.parse_args()
    DEBUG = args.debug

    raiz = Path(args.dir)
    if not raiz.exists():
        sys.exit(f"No existe: {raiz}")

    seven_zip = encontrar_7zip(args.seven_zip)
    licitaciones = detectar_licitaciones(raiz)
    metadata_licitaciones = cargar_metadata_licitaciones(args.metadata_csv)
    todos = []
    consumos = []
    proveedores_procesados = []

    print("=" * 74)
    print(f"EXTRACCION OPENAI PRUEBA | modelo={args.modelo} | licitaciones={len(licitaciones)}")
    print(f"7-Zip: {seven_zip or 'NO encontrado'} | solo_texto={args.solo_texto}")
    print("Logica: una llamada OpenAI por archivo + consolidacion opcional")
    print("=" * 74)

    detener = False
    for licitacion in licitaciones:
        oferentes = [
            carpeta for carpeta in licitacion.iterdir()
            if carpeta.is_dir() and not carpeta.name.startswith("_")
        ]
        if args.proveedor:
            aguja = args.proveedor.lower()
            oferentes = [c for c in oferentes if aguja in c.name.lower() or
                         aguja in str(cargar_json(c / "oferta.json").get("proveedor", "")).lower()]
        metadata = metadata_licitaciones.get(licitacion.name, {})
        salida = licitacion / "extraccion_openai.json"
        resultados = []
        resultados_previos = []
        if salida.exists():
            try:
                resultados_previos = json.loads(salida.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                resultados_previos = []
        if salida.exists() and not args.rehacer and not args.solo_tecnicos_fallback:
            resultados = resultados_previos
            for oferta in resultados:
                registrar_resultado(
                    licitacion.name, metadata, oferta, todos, consumos, proveedores_procesados
                )
            ruts_procesados = {str(oferta.get("rut") or "").strip() for oferta in resultados}
            oferentes = [
                carpeta for carpeta in oferentes
                if str(cargar_json(carpeta / "oferta.json").get("rut") or "").strip() not in ruts_procesados
            ]
            if not oferentes:
                print(f"{licitacion.name}: ya procesada ({len(resultados)} proveedores)")
                continue
            print(f"{licitacion.name}: reanudando; faltan {len(oferentes)} proveedores")

        if args.solo_tecnicos_fallback:
            resultados = []
            ruts_sin_precio = {
                str(oferta.get("rut") or "").strip()
                for oferta in resultados_previos
                if not any(
                    producto_tiene_precio(producto)
                    for producto in oferta.get("productos", [])
                    if isinstance(producto, dict)
                )
            }
            oferentes = [
                carpeta for carpeta in oferentes
                if str(cargar_json(carpeta / "oferta.json").get("rut") or "").strip()
                in ruts_sin_precio
            ]

        if args.limite_proveedores:
            oferentes = oferentes[:args.limite_proveedores]
        if not oferentes:
            continue

        titulo = metadata.get("nombre_licitacion") or "sin nombre en CSV"
        fecha = metadata.get("fecha_publicacion") or "sin fecha en CSV"
        print(f"\n{licitacion.name}: {titulo} | {fecha} | {len(oferentes)} proveedores")
        for carpeta in oferentes:
            try:
                resultado = procesar_oferta(
                    carpeta, args, seven_zip,
                    solo_tecnicos=args.solo_tecnicos_fallback
                )
                if args.solo_tecnicos_fallback:
                    rut_resultado = str(resultado.get("rut") or "").strip()
                    previo = next(
                        (oferta for oferta in resultados_previos
                         if str(oferta.get("rut") or "").strip() == rut_resultado),
                        None
                    )
                    if previo:
                        parciales = previo.get("productos_parciales", []) + resultado.get(
                            "productos_parciales", []
                        )
                        consolidados, estado, cruda, prompt, consumo = consolidar_parciales(
                            parciales, resultado["proveedor"], resultado["rut"],
                            resultado["total_oferta"], args
                        )
                        resultado["productos_parciales"] = parciales
                        resultado["productos"] = consolidados
                        resultado["estado_consolidacion"] = estado
                        resultado["consumo_consolidacion"] = consumo
                    resultados = [
                        oferta for oferta in resultados_previos
                        if str(oferta.get("rut") or "").strip() != rut_resultado
                    ]
            except RuntimeError as exc:
                print(f"\nERROR FATAL: {exc}")
                print("Se detiene la corrida porque La API dejo de responder.")
                detener = True
                break
            resultado.update({
                "codigo": licitacion.name,
                "nombre_licitacion": metadata.get("nombre_licitacion", ""),
                "fecha_publicacion": metadata.get("fecha_publicacion", ""),
                "estado_licitacion": metadata.get("estado_licitacion", ""),
                "organismo": metadata.get("organismo", "")
            })
            resultados.append(resultado)
            registrar_resultado(
                licitacion.name, metadata, resultado, todos, consumos, proveedores_procesados
            )
            tokens_proveedor = sum(
                valor_token(fila.get("total_tokens"))
                for fila in consumos
                if fila["codigo"] == licitacion.name and fila["rut"] == resultado["rut"]
            )
            print(
                f"  {resultado['proveedor'][:34]:34} "
                f"parciales={len(resultado['productos_parciales']):2} "
                f"finales={len(resultado['productos']):2} "
                f"ocr={len(resultado['necesita_ocr']):2} "
                f"consol={resultado['estado_consolidacion']} "
                f"tokens={tokens_proveedor}"
            )
            # Diagnostico visible por archivo: antes estos estados quedaban solo en JSON.
            for archivo in resultado["resultados_archivos"]:
                estado_ia = archivo.get("estado_ia", "")
                encontrados = archivo.get("productos_encontrados", 0)
                if estado_ia not in ("no_ejecutada", "solo_texto", "ok", "ok_reparado") or encontrados == 0:
                    print(
                        f"      [{archivo.get('tipo_documental','otro')}] "
                        f"{archivo.get('archivo','')[:47]} | lectura={archivo.get('estado_lectura')} "
                        f"| ia={estado_ia} | productos={encontrados}"
                    )
                    if archivo.get("observaciones"):
                        print(f"        observacion: {str(archivo['observaciones'])[:180]}")
            for producto in resultado["productos"]:
                print(f"    - {producto['producto'][:55]} | unitario={producto['precio_unitario']}")

            # checkpoint por proveedor para no perder el avance si Ollama cae
            salida.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")

        if resultados:
            salida.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")
        if detener:
            break

    ruta_excel = Path(args.excel) if args.excel else raiz / "resultado_openai.xlsx"
    ruta_excel_real = None
    if todos or consumos or proveedores_procesados:
        ruta_excel_real = generar_excel(todos, consumos, proveedores_procesados, ruta_excel)

    print("\n" + "=" * 74)
    print(f"Productos consolidados: {len(todos)}")
    print(f"Llamadas API registradas: {len(consumos)}")
    print(f"Prompt tokens: {sum(valor_token(fila.get('prompt_tokens')) for fila in consumos)}")
    print(f"Completion tokens: {sum(valor_token(fila.get('completion_tokens')) for fila in consumos)}")
    print(f"Tokens totales: {sum(valor_token(fila.get('total_tokens')) for fila in consumos)}")
    print(f"Excel: {ruta_excel_real if ruta_excel_real else 'no generado'}")
    print("=" * 74)


if __name__ == "__main__":
    main()

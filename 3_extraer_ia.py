#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor IA local para ofertas de Mercado Publico.

Flujo:
  1. Extrae el texto de CADA archivo por separado.
    2. Preserva tablas y divide documentos largos sin perder filas.
    3. Consulta Ollama o LM Studio secuencialmente.
    4. Consolida hallazgos por proveedor sin truncar registros.
    5. Valida alcance, evidencia y coherencia matematica.
    6. Guarda checkpoints, alertas y un log JSONL para reanudar corridas.

Formatos: PDF, DOCX, XLSX, XLSM, TXT, CSV, TSV, ZIP, RAR y 7Z.
Los PDF sin texto quedan registrados en necesita_ocr o se procesan con
Tesseract cuando se usa --ocr.

Instalacion:
    pip install -r requirements.txt

Ejemplos:
    python 3_extraer_ia.py --dir ofertas/3572-22-LE25 --backend lmstudio --modelo openai/gpt-oss-20b --rehacer --debug
    python 3_extraer_ia.py --dir ofertas/3572-22-LE25 --solo-texto --rehacer
"""

import argparse
import csv
import importlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path

try:
    import pymupdf as fitz
    import pdfplumber
    import requests
    from docx import Document
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
except ImportError as exc:
    sys.exit(
        f"Falta dependencia: {exc}. Instala: "
        "pip install pymupdf pdfplumber requests python-docx openpyxl"
    )

try:
    pytesseract = importlib.import_module("pytesseract")
    Image = importlib.import_module("PIL.Image")
except ModuleNotFoundError:
    pytesseract = None
    Image = None

OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
LMSTUDIO_URL = "http://127.0.0.1:1234/v1/chat/completions"
DEBUG = False
ESTADOS_IA_OK = {"ok", "ok_reparado", "ok_extraido"}

MARCAS = [
    "HP", "Lenovo", "Dell", "Asus", "Acer", "Apple", "Samsung", "Lexmark",
    "Epson", "Canon", "Brother", "MSI", "Huawei", "Kingston", "Logitech",
    "LG", "ThinkCentre", "ThinkPad", "ProOne", "ProDesk", "OptiPlex",
    "Latitude", "Pavilion"
]

from prompts_extraccion import (
        PROMPT_ARCHIVO,
        PROMPT_CONSOLIDAR,
)

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


def escribir_json_atomico(path, datos):
    temporal = path.with_suffix(path.suffix + ".tmp")
    temporal.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporal, path)


def cargar_metadata_licitaciones(ruta):
    ruta = Path(ruta)
    if not ruta.is_file():
        return {}
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            with ruta.open(encoding=encoding, newline="") as archivo:
                filas = list(csv.DictReader(archivo))
            break
        except UnicodeDecodeError:
            continue
    else:
        return {}
    return {
        (fila.get("codigo") or "").strip(): {
            "nombre_licitacion": (fila.get("nombre") or "").strip(),
            "fecha_publicacion": (fila.get("fecha_publicacion") or "").strip(),
            "estado_licitacion": (fila.get("estado") or "").strip(),
            "organismo": (fila.get("organismo") or "").strip(),
        }
        for fila in filas
        if (fila.get("codigo") or "").strip()
    }


def extraer_pdf_ocr(
    path,
    max_paginas,
    tesseract_cmd=None,
    idioma="spa+eng",
    dpi=180,
    paginas=None,
):
    if pytesseract is None or Image is None:
        return "", "ocr_dependencia_no_disponible"
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        doc = fitz.open(path)
        partes = []
        escala = dpi / 72
        for index, page in enumerate(doc):
            if index >= max_paginas:
                break
            if paginas is not None and index not in paginas:
                continue
            pixmap = page.get_pixmap(matrix=fitz.Matrix(escala, escala), alpha=False)
            with Image.open(io.BytesIO(pixmap.tobytes("png"))) as imagen:
                try:
                    texto = pytesseract.image_to_string(imagen, lang=idioma)
                except pytesseract.TesseractError:
                    texto = pytesseract.image_to_string(imagen, lang="eng")
            if texto.strip():
                partes.append(f"[PAGINA {index + 1} - OCR]\n{texto.strip()}")
        doc.close()
        texto = limpiar_control("\n\n".join(partes)).strip()
        return texto, "texto_ocr" if len(texto) > 40 else "necesita_ocr"
    except Exception as exc:
        return "", f"error_ocr: {exc}"


def tabla_estructurada(filas):
    filas_con_datos = [fila for fila in filas if any(str(celda or "").strip() for celda in fila)]
    if not filas_con_datos:
        return False
    filas_multicolumna = sum(
        sum(bool(str(celda or "").strip()) for celda in fila) >= 2
        for fila in filas_con_datos
    )
    return filas_multicolumna / len(filas_con_datos) >= 0.5


def objeto_fuera_de_tablas(objeto, cajas):
    objeto_izquierda = objeto.get("x0", 0)
    objeto_derecha = objeto.get("x1", 0)
    objeto_arriba = objeto.get("top", 0)
    objeto_abajo = objeto.get("bottom", 0)
    return not any(
        objeto_izquierda < derecha
        and objeto_derecha > izquierda
        and objeto_arriba < abajo
        and objeto_abajo > arriba
        for izquierda, arriba, derecha, abajo in cajas
    )


def extraer_pdf(path, max_paginas, usar_ocr=False, tesseract_cmd=None, idioma_ocr="spa+eng", dpi_ocr=180):
    try:
        partes = []
        paginas_sin_texto = []
        with pdfplumber.open(path) as pdf:
            for index, page in enumerate(pdf.pages):
                if index >= max_paginas:
                    break
                try:
                    tablas = page.find_tables()
                except Exception:
                    tablas = []
                tablas_validas = []
                tablas_colapsadas = 0
                for tabla in tablas:
                    filas = tabla.extract() or []
                    if tabla_estructurada(filas):
                        tablas_validas.append((tabla, filas))
                    else:
                        tablas_colapsadas += 1

                if tablas_validas:
                    cajas = [tabla.bbox for tabla, _ in tablas_validas]
                    pagina_sin_tablas = page.filter(lambda objeto: objeto_fuera_de_tablas(objeto, cajas))
                    texto_pagina = (pagina_sin_tablas.extract_text() or "").strip()
                else:
                    texto_pagina = (page.extract_text(layout=True) or page.extract_text() or "").strip()

                partes.append(f"[PAGINA {index + 1}]")
                if texto_pagina:
                    partes.append(texto_pagina)
                if tablas_colapsadas:
                    partes.append(f"[AVISO: {tablas_colapsadas} tabla(s) sin columnas recuperables]")
                for num_tabla, (_, filas) in enumerate(tablas_validas, 1):
                    partes.append(f"[TABLA {num_tabla} - pagina {index + 1}]")
                    for num_fila, fila in enumerate(filas, 1):
                        valores = [(celda or "").strip().replace("\n", " ") for celda in fila]
                        if any(valores):
                            partes.append(f"FILA {num_fila}: " + " || ".join(valores))
                if len(texto_pagina) <= 40 and not tablas_validas:
                    paginas_sin_texto.append(index)
        texto = limpiar_control("\n".join(partes)).strip()
    except Exception:
        texto = ""
        paginas_sin_texto = []

    if usar_ocr and paginas_sin_texto:
        texto_ocr, estado_ocr = extraer_pdf_ocr(
            path,
            max_paginas,
            tesseract_cmd,
            idioma_ocr,
            dpi_ocr,
            set(paginas_sin_texto),
        )
        if texto_ocr:
            texto = limpiar_control(f"{texto}\n\n{texto_ocr}").strip()
            return texto, "texto_ocr"
        if len(texto) > 40:
            return texto, f"texto_con_paginas_ocr_pendiente:{estado_ocr}"
        return "", estado_ocr

    if len(texto) > 40:
        estado = "texto_con_paginas_ocr_pendiente" if paginas_sin_texto else "texto"
        return texto, estado

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
        if len(texto) > 40:
            return texto, "texto"
    except Exception as exc:
        if not usar_ocr:
            return "", f"error_pdf: {exc}"
    if usar_ocr:
        return extraer_pdf_ocr(path, max_paginas, tesseract_cmd, idioma_ocr, dpi_ocr)
    return "", "necesita_ocr"


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
        return extraer_pdf(
            path,
            args.max_paginas,
            usar_ocr=args.ocr,
            tesseract_cmd=args.tesseract_cmd,
            idioma_ocr=args.idioma_ocr,
            dpi_ocr=args.dpi_ocr,
        )
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


def dividir_texto(texto, limite, lineas_solapadas=3):
    if len(texto) <= limite:
        return [texto]
    lineas = [linea for linea in texto.splitlines() if linea.strip()]
    fragmentos = []
    actual = []
    caracteres = 0
    for linea in lineas:
        trozos = [linea[indice:indice + limite] for indice in range(0, len(linea), limite)] or [""]
        for trozo in trozos:
            largo = len(trozo) + 1
            if actual and caracteres + largo > limite:
                fragmentos.append("\n".join(actual))
                actual = actual[-lineas_solapadas:]
                caracteres = sum(len(valor) + 1 for valor in actual)
            actual.append(trozo)
            caracteres += largo
    if actual:
        fragmentos.append("\n".join(actual))
    return fragmentos


def combinar_consumos(consumos, modelo):
    return {
        "modelo": next((fila.get("modelo") for fila in consumos if fila.get("modelo")), modelo),
        "prompt_tokens": sum(fila.get("prompt_tokens") or 0 for fila in consumos),
        "completion_tokens": sum(fila.get("completion_tokens") or 0 for fila in consumos),
        "total_tokens": sum(fila.get("total_tokens") or 0 for fila in consumos),
        "llamadas": len(consumos),
    }


def parsear_json(respuesta):
    texto = respuesta.strip()
    texto = re.sub(r"^```(?:json)?\s*|\s*```$", "", texto, flags=re.I | re.S)
    try:
        return json.loads(texto), "ok"
    except Exception:
        pass
    decoder = json.JSONDecoder()
    errores = []
    candidatos = []
    for coincidencia in re.finditer(r"\{", texto):
        fragmento = texto[coincidencia.start():]
        try:
            candidato, _ = decoder.raw_decode(fragmento)
            if isinstance(candidato, dict):
                candidatos.append(candidato)
        except json.JSONDecodeError as exc:
            errores.append(str(exc))
    for candidato in candidatos:
        if isinstance(candidato.get("productos"), list):
            return candidato, "ok_extraido"
    if candidatos:
        return candidatos[0], "ok_extraido"

    coincidencia = re.search(r"\{.*\}", texto, re.S)
    if coincidencia:
        reparado = re.sub(r",\s*([}\]])", r"\1", coincidencia.group(0))
        try:
            return json.loads(reparado), "ok_reparado"
        except Exception as exc:
            errores.append(str(exc))
    return None, f"json_invalido: {errores[-1]}" if errores else "sin_json"


def consultar_modelo(
    prompt,
    modelo,
    timeout,
    num_ctx,
    backend="ollama",
    max_tokens=4096,
    reintentos=2,
    espera_reintento=5,
):
    if backend == "lmstudio":
        payload = {
            "model": modelo,
            "messages": [
                {"role": "system", "content": "Devuelve exclusivamente un objeto JSON valido."},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.0,
            "max_tokens": max_tokens,
            "reasoning_effort": "none",
            "stream": False,
            "response_format": {"type": "text"},
            "chat_template_kwargs": {"enable_thinking": False},
        }
        url = LMSTUDIO_URL
    else:
        payload = {
            "model": modelo,
            "prompt": prompt,
            "stream": False,
            "format": "json",
            "options": {
                "temperature": 0.1,
                "num_ctx": num_ctx,
                "num_predict": max_tokens
            }
        }
        url = OLLAMA_URL
    payload = {
        **payload,
    }
    ultimo_error = None
    for intento in range(reintentos + 1):
        try:
            respuesta = requests.post(url, json=payload, timeout=timeout)
            if respuesta.status_code >= 400:
                estado = f"http_{respuesta.status_code}: {respuesta.text[:500]}"
                if respuesta.status_code not in (429, 500, 502, 503, 504) or intento >= reintentos:
                    return None, estado, "", {}
                ultimo_error = estado
            else:
                datos_respuesta = respuesta.json()
                if backend == "lmstudio":
                    cruda = ((datos_respuesta.get("choices") or [{}])[0]
                             .get("message", {}).get("content") or "")
                else:
                    cruda = datos_respuesta.get("response", "")
                datos, estado = parsear_json(cruda)
                if backend == "lmstudio":
                    uso = datos_respuesta.get("usage") or {}
                    consumo = {
                        "modelo": datos_respuesta.get("model", modelo),
                        "prompt_tokens": uso.get("prompt_tokens"),
                        "completion_tokens": uso.get("completion_tokens"),
                        "total_tokens": uso.get("total_tokens"),
                    }
                else:
                    consumo = {
                        "modelo": modelo,
                        "prompt_tokens": datos_respuesta.get("prompt_eval_count"),
                        "completion_tokens": datos_respuesta.get("eval_count"),
                        "total_tokens": (
                            (datos_respuesta.get("prompt_eval_count") or 0)
                            + (datos_respuesta.get("eval_count") or 0)
                        ),
                    }
                return datos, estado, cruda, consumo
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as exc:
            ultimo_error = str(exc)
            if intento >= reintentos:
                raise RuntimeError(f"{backend} dejo de responder tras {reintentos + 1} intentos: {exc}")
        except Exception as exc:
            return None, f"error_{backend}: {exc}", "", {}
        time.sleep(espera_reintento)
    return None, f"error_{backend}: {ultimo_error}", "", {}


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
    precio_total_tipo = str(producto.get("precio_total_tipo") or "").strip().lower() or None
    if precio_total_tipo not in (None, "linea", "oferta"):
        precio_total_tipo = None
    cantidad_fuente = producto.get("cantidad_fuente")
    if (
        precio_unitario is None
        and precio_total is not None
        and cantidad not in (None, 0)
        and precio_total_tipo == "linea"
    ):
        precio_unitario = round(float(precio_total) / float(cantidad), 2)
    if (
        cantidad is None
        and precio_unitario not in (None, 0)
        and precio_total not in (None, 0)
        and precio_total_tipo == "linea"
    ):
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
        "precio_total_tipo": precio_total_tipo,
        "moneda": producto.get("moneda"),
        "confianza": producto.get("confianza"),
        "pagina": producto.get("pagina"),
        "fila_fuente": producto.get("fila_fuente"),
        "evidencia": producto.get("evidencia"),
        "archivo_fuente": archivo,
        "tipo_documental": tipo_doc
    }


PATRON_EQUIPO_RELEVANTE = re.compile(
    r"notebook|laptop|port[aá]til|ultrabook|chromebook|all.?in.?one|\baio\b|"
    r"desktop|escritorio|\bpc\b|computador(?:a)?|workstation|thinkcentre|"
    r"thinkpad|prodesk|optiplex|latitude|pavilion", re.I
)
PATRON_TV = re.compile(r"\b(?:smart\s*)?t\.?v\.?\b|televisor", re.I)
PATRON_EXCLUIDO = re.compile(
    r"proyector|tablet|celular|smartphone|servidor|storage|switch|router|"
    r"access point|\bred\b|rack|cable|"
    r"teclado|mouse|rat[oó]n|docking|dock|base de expansi[oó]n|"
    r"tinta|t[oó]ner|cartucho|tambor|repuesto|licencia|instalaci[oó]n|servicio",
    re.I
)
PATRON_MONITOR = re.compile(r"monitor|display", re.I)
PATRON_IMPRESORA = re.compile(r"impresora|multifuncional|plotter", re.I)
PATRON_ESPECIFICACION = re.compile(
    r"procesador|cpu|\bram\b|memoria(?:\s+(?:ram|ddr))?|\bssd\b|\bhdd\b|"
    r"disco(?:\s+duro)?|puertos?|conectividad|sistema operativo|windows|"
    r"fuente de poder|garant[ií]a|servicio|instalaci[oó]n|configuraci[oó]n|"
    r"servidor|storage|switch|router|access point|\bred\b|cable|rack", re.I
)


def clasificar_producto(producto):
    texto = " ".join(str(producto.get(c) or "") for c in ("producto", "modelo", "categoria"))
    if PATRON_TV.search(texto):
        return None
    if PATRON_EQUIPO_RELEVANTE.search(texto):
        return "equipo"
    if PATRON_EXCLUIDO.search(texto):
        return None
    if PATRON_ESPECIFICACION.search(texto) and not any(
        patron.search(texto) for patron in (PATRON_EQUIPO_RELEVANTE, PATRON_MONITOR, PATRON_IMPRESORA)
    ):
        return None
    if PATRON_MONITOR.search(texto):
        return "monitor"
    if PATRON_IMPRESORA.search(texto):
        return "impresora"
    categoria_declarada = str(producto.get("categoria") or "").strip().lower()
    if categoria_declarada in {"equipo", "monitor", "impresora"}:
        return categoria_declarada
    return None


def filtrar_productos_relevantes(productos):
    clasificados = []
    for producto in productos:
        categoria = clasificar_producto(producto)
        if categoria:
            producto["categoria"] = categoria
            clasificados.append(producto)
    return clasificados


def texto_normalizado(valor):
    texto = re.sub(r"[^a-z0-9]+", " ", str(valor or "").lower())
    return " ".join(texto.split())


def numeros_iguales(valor_a, valor_b, tolerancia=0.01):
    if valor_a in (None, "") or valor_b in (None, ""):
        return True
    valor_a = float(valor_a)
    valor_b = float(valor_b)
    return abs(valor_a - valor_b) <= max(1.0, abs(valor_b) * tolerancia)


def productos_equivalentes(producto_a, producto_b):
    if producto_a.get("categoria") != producto_b.get("categoria"):
        return False
    item_a = texto_normalizado(producto_a.get("item"))
    item_b = texto_normalizado(producto_b.get("item"))
    if item_a and item_b and item_a != item_b:
        return False
    mismo_modelo = (
        len(texto_normalizado(producto_a.get("modelo"))) >= 4
        and texto_normalizado(producto_a.get("modelo"))
        == texto_normalizado(producto_b.get("modelo"))
    )
    misma_descripcion = (
        texto_normalizado(producto_a.get("producto"))
        == texto_normalizado(producto_b.get("producto"))
    )
    mismos_numeros = all(
        numeros_iguales(producto_a.get(campo), producto_b.get(campo))
        for campo in ("cantidad", "precio_unitario", "precio_total")
    )
    return mismos_numeros and (mismo_modelo or misma_descripcion)


def deduplicar_productos(productos):
    unicos = []
    for producto in productos:
        existente = next(
            (candidato for candidato in unicos if productos_equivalentes(candidato, producto)),
            None,
        )
        if existente is None:
            unicos.append(producto)
            continue
        fuentes = set(existente.get("fuentes_respaldo") or [])
        fuentes.update(producto.get("fuentes_respaldo") or [])
        for campo in ("archivo_fuente", "fuente_producto", "fuente_precio"):
            if existente.get(campo):
                fuentes.add(str(existente[campo]))
            if producto.get(campo):
                fuentes.add(str(producto[campo]))
        existente["fuentes_respaldo"] = sorted(fuentes)
        for campo in (
            "marca", "modelo", "item", "cantidad", "cantidad_fuente",
            "precio_unitario", "precio_total", "precio_total_tipo", "moneda", "evidencia",
            "pagina", "fila_fuente", "fuente_producto", "fuente_precio",
        ):
            if not existente.get(campo) and producto.get(campo):
                existente[campo] = producto[campo]
    return unicos


def coincide_con_parcial(producto, parcial):
    if producto.get("categoria") != parcial.get("categoria"):
        return False
    modelo = texto_normalizado(producto.get("modelo"))
    modelo_parcial = texto_normalizado(parcial.get("modelo"))
    if len(modelo) >= 4 and modelo == modelo_parcial:
        return True
    tokens = set(texto_normalizado(producto.get("producto")).split())
    tokens_parcial = set(texto_normalizado(parcial.get("producto")).split())
    if not tokens or not tokens_parcial:
        return False
    return len(tokens & tokens_parcial) / min(len(tokens), len(tokens_parcial)) >= 0.5


def misma_identidad_producto(producto_a, producto_b):
    if producto_a.get("categoria") != producto_b.get("categoria"):
        return False
    modelo_a = texto_normalizado(producto_a.get("modelo"))
    modelo_b = texto_normalizado(producto_b.get("modelo"))
    if len(modelo_a) >= 4 and modelo_a == modelo_b:
        return True
    item_a = texto_normalizado(producto_a.get("item"))
    item_b = texto_normalizado(producto_b.get("item"))
    if item_a and item_a == item_b:
        return True
    return texto_normalizado(producto_a.get("producto")) == texto_normalizado(producto_b.get("producto"))


def validar_productos(productos, parciales, total_oferta):
    productos = deduplicar_productos(filtrar_productos_relevantes(productos))
    parciales = filtrar_productos_relevantes(parciales)
    for producto in productos:
        fuentes_declaradas = {
            str(producto.get(campo))
            for campo in ("archivo_fuente", "fuente_producto", "fuente_precio")
            if producto.get(campo) and producto.get(campo) != "consolidacion"
        }
        fuentes = {
            str(parcial.get("archivo_fuente"))
            for parcial in parciales
            if parcial.get("archivo_fuente")
            and (
                str(parcial.get("archivo_fuente")) in fuentes_declaradas
                or coincide_con_parcial(producto, parcial)
            )
        }
        producto["fuentes_respaldo"] = sorted(fuentes)
        alertas = []
        cantidad = producto.get("cantidad")
        unitario = producto.get("precio_unitario")
        total = producto.get("precio_total")
        if not producto.get("marca"):
            alertas.append("marca_no_identificada")
        if cantidad in (None, ""):
            alertas.append("cantidad_no_identificada")
        elif float(cantidad) <= 0 or abs(float(cantidad) - round(float(cantidad))) > 0.001:
            alertas.append("cantidad_invalida")
        if unitario in (None, ""):
            alertas.append("precio_unitario_no_identificado")
        elif float(unitario) <= 0:
            alertas.append("precio_unitario_invalido")
        if total not in (None, "") and producto.get("precio_total_tipo") is None:
            alertas.append("origen_precio_total_no_identificado")
        elif total not in (None, "") and producto.get("precio_total_tipo") == "oferta":
            alertas.append("total_general_asignado_a_producto")
        if all(valor not in (None, "", 0) for valor in (cantidad, unitario, total)):
            calculado = float(cantidad) * float(unitario)
            if abs(calculado - float(total)) > max(2.0, abs(float(total)) * 0.01):
                alertas.append("cantidad_por_unitario_no_coincide_con_total")
        if not fuentes:
            alertas.append("sin_respaldo_en_hallazgos_parciales")
        if str(producto.get("confianza") or "").lower() == "baja":
            alertas.append("confianza_modelo_baja")
        producto["alertas"] = alertas

    for indice, producto in enumerate(productos):
        for otro in productos[indice + 1:]:
            if not misma_identidad_producto(producto, otro):
                continue
            campos_conflictivos = [
                campo for campo in ("cantidad", "precio_unitario", "precio_total")
                if producto.get(campo) not in (None, "")
                and otro.get(campo) not in (None, "")
                and not numeros_iguales(producto[campo], otro[campo])
            ]
            for campo in campos_conflictivos:
                alerta = f"conflicto_entre_fuentes_{campo}"
                producto["alertas"].append(alerta)
                otro["alertas"].append(alerta)

    total_referencia = normalizar_numero(total_oferta)
    totales = [producto.get("precio_total") for producto in productos]
    if total_referencia and totales and all(isinstance(total, (int, float)) for total in totales):
        suma = sum(totales)
        if suma > float(total_referencia) * 1.02:
            for producto in productos:
                producto["alertas"].append("suma_productos_supera_total_oferta")

    for producto in productos:
        alertas = producto["alertas"]
        if any(
            "invalida" in alerta
            or "no_coincide" in alerta
            or "supera_total" in alerta
            or "conflicto_entre_fuentes" in alerta
            or "total_general" in alerta
            for alerta in alertas
        ):
            producto["estado_validacion"] = "inconsistente"
        elif any("no_identificad" in alerta for alerta in alertas):
            producto["estado_validacion"] = "incompleto"
        elif alertas:
            producto["estado_validacion"] = "revisar"
        else:
            producto["estado_validacion"] = "ok"
    return productos


def extraer_parciales_archivo(path, texto, proveedor, rut, total_oferta, args):
    relativo = str(path)
    tipo_doc = tipo_documental(path.name)
    fragmentos = dividir_texto(texto, args.max_chars_archivo)
    productos = []
    estados = []
    respuestas = []
    prompts = []
    observaciones = []
    consumos = []
    for numero, fragmento in enumerate(fragmentos, 1):
        productos_pista, precios_pista = pistas(fragmento)
        texto_fragmento = f"[FRAGMENTO {numero} DE {len(fragmentos)}]\n{fragmento}"
        prompt = PROMPT_ARCHIVO.format(
            proveedor=proveedor,
            rut=rut,
            archivo=relativo,
            tipo_documental=tipo_doc,
            total_oferta=total_oferta or "no informado",
            pistas_producto="\n".join(productos_pista) or "(ninguna)",
            pistas_precio="\n".join(precios_pista) or "(ninguna)",
            texto=texto_fragmento
        )
        datos, estado, cruda, consumo = consultar_modelo(
            prompt, args.modelo, args.timeout, args.num_ctx, args.backend, args.max_tokens,
            args.reintentos_modelo, args.espera_reintento
        )
        estados.append(estado)
        respuestas.append(f"[FRAGMENTO {numero}]\n{cruda}")
        prompts.append(prompt)
        consumos.append(consumo)
        if isinstance(datos, dict):
            if datos.get("observaciones"):
                observaciones.append(str(datos["observaciones"]))
            for producto in datos.get("productos", []):
                if isinstance(producto, dict):
                    limpio = normalizar_producto(producto, proveedor, rut, relativo, tipo_doc)
                    if limpio:
                        productos.append(limpio)
        if args.pausa_archivo > 0:
            time.sleep(args.pausa_archivo)
    estado_final = "ok" if all(estado in ESTADOS_IA_OK for estado in estados) else "; ".join(estados)
    return (
        productos,
        estado_final,
        "\n\n".join(respuestas),
        "\n\n".join(prompts),
        "; ".join(observaciones) or None,
        combinar_consumos(consumos, args.modelo),
    )


def dividir_registros_json(registros, limite):
    lotes = []
    actual = []
    largo = 2
    for registro in registros:
        serializado = json.dumps(registro, ensure_ascii=False, separators=(",", ":"))
        adicional = len(serializado) + (1 if actual else 0)
        if actual and largo + adicional > limite:
            lotes.append(actual)
            actual = []
            largo = 2
        actual.append(registro)
        largo += adicional
    if actual:
        lotes.append(actual)
    return lotes


def normalizar_consolidados(datos, respaldo, proveedor, rut):
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
    return consolidados or respaldo


def consolidar_parciales(parciales, proveedor, rut, total_oferta, args):
    parciales = deduplicar_productos(filtrar_productos_relevantes(parciales))
    if not parciales or args.sin_consolidar:
        return validar_productos(parciales, parciales, total_oferta), "omitida", "", "", {}

    lotes = dividir_registros_json(parciales, args.max_chars_consolidacion)
    consolidados = []
    estados = []
    respuestas = []
    prompts = []
    consumos = []
    for numero, lote in enumerate(lotes, 1):
        resumen = json.dumps(lote, ensure_ascii=False, separators=(",", ":"))
        prompt = PROMPT_CONSOLIDAR.format(
            proveedor=proveedor,
            rut=rut,
            total_oferta=total_oferta or "no informado",
            parciales=f"LOTE {numero} DE {len(lotes)}\n{resumen}"
        )
        datos, estado, cruda, consumo = consultar_modelo(
            prompt, args.modelo, args.timeout, args.num_ctx, args.backend, args.max_tokens,
            args.reintentos_modelo, args.espera_reintento
        )
        consolidados.extend(normalizar_consolidados(datos, lote, proveedor, rut))
        estados.append(estado)
        respuestas.append(f"[LOTE {numero}]\n{cruda}")
        prompts.append(prompt)
        consumos.append(consumo)
        if args.pausa_archivo > 0:
            time.sleep(args.pausa_archivo)

    consolidados = deduplicar_productos(filtrar_productos_relevantes(consolidados))
    if len(lotes) > 1:
        resumen_final = json.dumps(consolidados, ensure_ascii=False, separators=(",", ":"))
        if len(resumen_final) <= args.max_chars_consolidacion:
            prompt = PROMPT_CONSOLIDAR.format(
                proveedor=proveedor,
                rut=rut,
                total_oferta=total_oferta or "no informado",
                parciales=f"CONSOLIDACION FINAL DE {len(lotes)} LOTES\n{resumen_final}"
            )
            datos, estado, cruda, consumo = consultar_modelo(
                prompt, args.modelo, args.timeout, args.num_ctx, args.backend, args.max_tokens,
                args.reintentos_modelo, args.espera_reintento
            )
            consolidados = normalizar_consolidados(datos, consolidados, proveedor, rut)
            estados.append(estado)
            respuestas.append(f"[CONSOLIDACION FINAL]\n{cruda}")
            prompts.append(prompt)
            consumos.append(consumo)

    estado_final = "ok" if all(estado in ESTADOS_IA_OK for estado in estados) else "; ".join(estados)
    return (
        validar_productos(consolidados, parciales, total_oferta),
        estado_final,
        "\n\n".join(respuestas),
        "\n\n".join(prompts),
        combinar_consumos(consumos, args.modelo),
    )


def procesar_oferta(carpeta, args, seven_zip):
    info = cargar_json(carpeta / "oferta.json")
    proveedor = info.get("proveedor") or carpeta.name
    rut = info.get("rut") or carpeta.name.split("__", 1)[0]
    total_oferta = info.get("total") or info.get("total_oferta") or ""
    archivos_comprimidos = descomprimir(carpeta, seven_zip)

    archivos = [
        path for path in carpeta.rglob("*")
        if path.is_file()
        and path.name not in {"oferta.json", "extraccion_ia.json", ".extraido_ok"}
        and "_debug_ia" not in path.parts
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

        if "ocr_pendiente" in estado_lectura:
            necesita_ocr.append(relativo)
        if estado_lectura in ("necesita_ocr", "ocr_dependencia_no_disponible") or estado_lectura.startswith("error_ocr"):
            if relativo not in necesita_ocr:
                necesita_ocr.append(relativo)
        elif estado_lectura in ("no_soportado", "xls_legacy_no_soportado"):
            no_soportados.append(relativo)
        elif estado_lectura.startswith("texto") and texto and not args.solo_texto:
            productos, estado_ia, cruda, prompt, observaciones, consumo = extraer_parciales_archivo(
                Path(relativo), texto, proveedor, rut, total_oferta, args
            )
            parciales.extend(productos)
            registro["estado_ia"] = estado_ia
            registro["productos_encontrados"] = len(productos)
            registro["observaciones"] = observaciones
            registro["consumo_local"] = consumo
            registro["fragmentos"] = consumo.get("llamadas", 1)
            if args.debug:
                base = f"{indice:02d}__{re.sub(r'[^\w.-]', '_', path.stem)[:70]}"
                (debug_dir / f"{base}__prompt.txt").write_text(prompt, encoding="utf-8")
                (debug_dir / f"{base}__respuesta.txt").write_text(cruda, encoding="utf-8")
        elif args.solo_texto and estado_lectura.startswith("texto"):
            registro["estado_ia"] = "solo_texto"

        resultados_archivos.append(registro)

    productos_finales = []
    estado_consolidacion = "no_ejecutada"
    consumo_consolidacion = {}
    if not args.solo_texto:
        productos_finales, estado_consolidacion, cruda_consolidacion, prompt_consolidacion, consumo_consolidacion = consolidar_parciales(
            parciales, proveedor, rut, total_oferta, args
        )
        if args.debug and prompt_consolidacion:
            (debug_dir / "99__consolidacion__prompt.txt").write_text(prompt_consolidacion, encoding="utf-8")
            (debug_dir / "99__consolidacion__respuesta.txt").write_text(cruda_consolidacion, encoding="utf-8")

    return {
        "proveedor": proveedor,
        "rut": rut,
        "modelo": args.modelo,
        "backend": args.backend,
        "estado_oferta": info.get("estado", ""),
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


def escribir_log(ruta, evento, **datos):
    registro = {
        "fecha": datetime.now().isoformat(timespec="seconds"),
        "evento": evento,
        **datos,
    }
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("a", encoding="utf-8") as archivo:
        archivo.write(json.dumps(registro, ensure_ascii=False) + "\n")


def resultado_requiere_reproceso(resultado, usar_ocr=False, modelo=None, backend=None):
    estados = [
        str(archivo.get("estado_ia") or "")
        for archivo in resultado.get("resultados_archivos", [])
        if archivo.get("estado_ia") not in ("no_ejecutada", "solo_texto")
    ]
    estado_consolidacion = str(resultado.get("estado_consolidacion") or "")
    fallo_modelo = any(estado not in ESTADOS_IA_OK for estado in estados)
    fallo_consolidacion = estado_consolidacion not in ESTADOS_IA_OK | {"", "omitida", "no_ejecutada"}
    cambio_modelo = bool(
        modelo
        and resultado.get("modelo")
        and str(resultado.get("modelo")) != str(modelo)
    )
    cambio_backend = bool(
        backend
        and resultado.get("backend")
        and str(resultado.get("backend")) != str(backend)
    )
    return (
        fallo_modelo
        or fallo_consolidacion
        or cambio_modelo
        or cambio_backend
        or (usar_ocr and bool(resultado.get("necesita_ocr")))
    )


def registrar_resultado(codigo, metadata, resultado, productos, consumos, resumenes, archivos_ocr):
    resultado["productos"] = validar_productos(
        resultado.get("productos", []),
        resultado.get("productos_parciales", []),
        resultado.get("total_oferta"),
    )
    contexto = {
        "codigo": codigo,
        "nombre_licitacion": metadata.get("nombre_licitacion", ""),
        "fecha_publicacion": metadata.get("fecha_publicacion", ""),
        "estado_licitacion": metadata.get("estado_licitacion", ""),
        "organismo": metadata.get("organismo", ""),
    }
    for producto in resultado["productos"]:
        productos.append({**contexto, **producto})

    llamadas = 0
    errores_modelo = 0
    for archivo in resultado.get("resultados_archivos", []):
        estado = str(archivo.get("estado_ia") or "no_ejecutada")
        if estado in ("no_ejecutada", "solo_texto"):
            continue
        consumo = archivo.get("consumo_local") or {}
        llamadas_archivo = int(consumo.get("llamadas") or 1)
        llamadas += llamadas_archivo
        if estado not in ESTADOS_IA_OK:
            errores_modelo += 1
        consumos.append({
            **contexto,
            "proveedor": resultado.get("proveedor", ""),
            "rut": resultado.get("rut", ""),
            "tipo_llamada": "archivo",
            "archivo": archivo.get("archivo", ""),
            "estado": estado,
            **consumo,
        })

    estado_consolidacion = str(resultado.get("estado_consolidacion") or "no_ejecutada")
    if estado_consolidacion not in ("no_ejecutada", "omitida"):
        consumo = resultado.get("consumo_consolidacion") or {}
        llamadas_consolidacion = int(consumo.get("llamadas") or 1)
        llamadas += llamadas_consolidacion
        if estado_consolidacion not in ESTADOS_IA_OK:
            errores_modelo += 1
        consumos.append({
            **contexto,
            "proveedor": resultado.get("proveedor", ""),
            "rut": resultado.get("rut", ""),
            "tipo_llamada": "consolidacion",
            "archivo": "",
            "estado": estado_consolidacion,
            **consumo,
        })

    pendientes_ocr = resultado.get("necesita_ocr", [])
    archivos_ocr.extend(
        {
            "licitacion": codigo,
            "proveedor": resultado.get("proveedor", ""),
            "archivo": archivo,
        }
        for archivo in pendientes_ocr
    )
    productos_validos = sum(
        producto.get("estado_validacion") == "ok" for producto in resultado["productos"]
    )
    resumenes.append({
        **contexto,
        "proveedor": resultado.get("proveedor", ""),
        "rut": resultado.get("rut", ""),
        "estado_oferta": resultado.get("estado_oferta", ""),
        "productos": len(resultado["productos"]),
        "productos_validos": productos_validos,
        "productos_revisar": len(resultado["productos"]) - productos_validos,
        "llamadas": llamadas,
        "errores_modelo": errores_modelo,
        "ocr": len(pendientes_ocr),
        "no_soportados": len(resultado.get("no_soportados", [])),
    })


def generar_excel(productos, consumos, resumenes, ruta):
    columnas = [
        "codigo", "nombre_licitacion", "fecha_publicacion", "estado_licitacion",
        "organismo", "proveedor", "rut", "item", "producto", "categoria", "marca", "modelo",
        "cantidad", "cantidad_fuente", "precio_unitario", "precio_total", "moneda",
        "precio_total_tipo",
        "fuente_producto", "fuente_precio", "archivo_fuente", "pagina", "fila_fuente",
        "evidencia", "fuentes_respaldo", "confianza", "estado_validacion", "alertas"
    ]
    consumo_columnas = [
        "codigo", "proveedor", "rut", "tipo_llamada", "archivo", "estado",
        "modelo", "prompt_tokens", "completion_tokens", "total_tokens", "llamadas"
    ]
    resumen_columnas = [
        "codigo", "nombre_licitacion", "fecha_publicacion", "organismo", "proveedor", "rut",
        "estado_oferta", "productos", "productos_validos", "productos_revisar",
        "llamadas", "errores_modelo", "ocr", "no_soportados"
    ]
    alerta_columnas = [
        "codigo", "proveedor", "rut", "tipo", "detalle", "producto",
        "cantidad", "precio_unitario", "precio_total", "archivo_fuente"
    ]

    wb = Workbook()
    wb.remove(wb.active)

    def crear_hoja(nombre, campos, filas):
        ws = wb.create_sheet(nombre)
        for columna, campo in enumerate(campos, 1):
            celda = ws.cell(1, columna, campo)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="1F4E78")
        for fila, datos in enumerate(filas, 2):
            for columna, campo in enumerate(campos, 1):
                valor = datos.get(campo)
                if isinstance(valor, list):
                    valor = "; ".join(str(elemento) for elemento in valor)
                ws.cell(fila, columna, valor)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    alertas = []
    for producto in productos:
        for alerta in producto.get("alertas") or []:
            alertas.append({
                "codigo": producto.get("codigo"),
                "proveedor": producto.get("proveedor"),
                "rut": producto.get("rut"),
                "tipo": "producto",
                "detalle": alerta,
                "producto": producto.get("producto"),
                "cantidad": producto.get("cantidad"),
                "precio_unitario": producto.get("precio_unitario"),
                "precio_total": producto.get("precio_total"),
                "archivo_fuente": producto.get("archivo_fuente"),
            })
    for resumen in resumenes:
        if resumen.get("productos", 0) == 0 and resumen.get("estado_oferta") != "Rechazada":
            alertas.append({**resumen, "tipo": "proveedor", "detalle": "sin_productos_en_alcance"})
        if resumen.get("ocr", 0):
            alertas.append({**resumen, "tipo": "proveedor", "detalle": "archivos_pendientes_ocr"})
        if resumen.get("errores_modelo", 0):
            alertas.append({**resumen, "tipo": "proveedor", "detalle": "errores_modelo"})

    crear_hoja("Productos", columnas, productos)
    crear_hoja(
        "Productos_validos",
        columnas,
        [producto for producto in productos if producto.get("estado_validacion") == "ok"],
    )
    crear_hoja("Alertas", alerta_columnas, alertas)
    crear_hoja("Consumo", consumo_columnas, consumos)
    crear_hoja("Resumen", resumen_columnas, resumenes)
    wb.save(ruta)


def main():
    global DEBUG
    parser = argparse.ArgumentParser(description="Extraccion por archivo y consolidacion ligera por proveedor")
    parser.add_argument("--dir", default="ofertas")
    parser.add_argument("--modelo", default="llama3.2")
    parser.add_argument("--backend", choices=("ollama", "lmstudio"), default="ollama",
                        help="Motor local: ollama o lmstudio")
    parser.add_argument("--solo-texto", action="store_true")
    parser.add_argument("--sin-consolidar", action="store_true")
    parser.add_argument("--excel")
    parser.add_argument("--rehacer", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--seven-zip")
    parser.add_argument("--max-paginas", type=int, default=20)
    parser.add_argument("--max-filas-excel", type=int, default=500)
    parser.add_argument("--max-columnas-excel", type=int, default=40)
    parser.add_argument("--ocr", action="store_true",
                        help="Aplica Tesseract solo a PDF sin texto extraible")
    parser.add_argument("--tesseract-cmd",
                        help="Ruta a tesseract.exe si no esta disponible en PATH")
    parser.add_argument("--idioma-ocr", default="spa+eng")
    parser.add_argument("--dpi-ocr", type=int, default=180)
    parser.add_argument("--max-chars-archivo", type=int, default=7000)
    parser.add_argument("--max-chars-consolidacion", type=int, default=12000)
    parser.add_argument("--num-ctx", type=int, default=8192)
    parser.add_argument("--timeout", type=int, default=300)
    parser.add_argument("--reintentos-modelo", type=int, default=2)
    parser.add_argument("--espera-reintento", type=int, default=5)
    parser.add_argument("--log", help="Ruta del log JSONL; por defecto se guarda junto al Excel")
    parser.add_argument("--limite-licitaciones", type=int,
                        help="Procesa como maximo esta cantidad de licitaciones")
    parser.add_argument("--limite-proveedores", type=int,
                        help="Procesa como maximo esta cantidad de proveedores por licitacion")
    parser.add_argument("--pausa-licitacion", type=int, default=0,
                        help="Pausa en segundos despues de cada licitacion")
    parser.add_argument("--pausa-archivo", type=int, default=0,
                        help="Pausa en segundos entre llamadas al modelo")
    parser.add_argument("--max-tokens", type=int, default=4096,
                        help="Maximo de tokens generados por llamada LM Studio")
    parser.add_argument("--proveedor",
                        help="Procesa solo el proveedor cuyo nombre coincide con la carpeta")
    parser.add_argument("--metadata-csv", default=str(Path(__file__).resolve().with_name("para_scrapear.csv")),
                        help="CSV con codigo, nombre y fecha_publicacion")
    args = parser.parse_args()
    DEBUG = args.debug

    raiz = Path(args.dir)
    if not raiz.exists():
        sys.exit(f"No existe: {raiz}")

    seven_zip = encontrar_7zip(args.seven_zip)
    licitaciones = detectar_licitaciones(raiz)
    if args.limite_licitaciones is not None:
        if args.limite_licitaciones < 1:
            sys.exit("--limite-licitaciones debe ser mayor que cero")
        licitaciones = licitaciones[:args.limite_licitaciones]
    metadata_licitaciones = cargar_metadata_licitaciones(args.metadata_csv)
    todos = []
    consumos = []
    resumenes = []
    archivos_ocr = []
    ruta_excel = Path(args.excel) if args.excel else raiz / "resultado_productos.xlsx"
    ruta_log = Path(args.log) if args.log else ruta_excel.with_suffix(".log.jsonl")

    print("=" * 74)
    print(f"EXTRACCION IA v5 | backend={args.backend} | modelo={args.modelo} | licitaciones={len(licitaciones)}")
    print(f"7-Zip: {seven_zip or 'NO encontrado'} | solo_texto={args.solo_texto}")
    print("Logica: fragmentos secuenciales + consolidacion y validacion auditable")
    print("=" * 74)
    escribir_log(
        ruta_log,
        "inicio_corrida",
        directorio=str(raiz),
        backend=args.backend,
        modelo=args.modelo,
        licitaciones=len(licitaciones),
    )

    detener = False
    for licitacion in licitaciones:
        oferentes = [
            carpeta for carpeta in licitacion.iterdir()
            if carpeta.is_dir() and not carpeta.name.startswith("_")
        ]
        if args.proveedor:
            oferentes = [carpeta for carpeta in oferentes if carpeta.name == args.proveedor]
            if not oferentes:
                sys.exit(
                    f"No se encontro el proveedor '{args.proveedor}' en {licitacion}"
                )
        if not oferentes:
            continue

        salida = licitacion / "extraccion_ia.json"
        resultados = []
        if salida.exists() and not args.rehacer:
            try:
                resultados = json.loads(salida.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                resultados = []
            resultados_reintentar = [
                resultado for resultado in resultados
                if resultado_requiere_reproceso(resultado, args.ocr, args.modelo, args.backend)
            ]
            ruts_reintentar = {
                str(resultado.get("rut") or "").strip()
                for resultado in resultados_reintentar
            }
            resultados = [
                resultado for resultado in resultados
                if str(resultado.get("rut") or "").strip() not in ruts_reintentar
            ]
            ruts_procesados = {
                str(resultado.get("rut") or "").strip() for resultado in resultados
            }
            oferentes = [
                carpeta for carpeta in oferentes
                if str(cargar_json(carpeta / "oferta.json").get("rut") or "").strip()
                not in ruts_procesados
            ]
            for resultado in resultados:
                registrar_resultado(
                    licitacion.name,
                    metadata_licitaciones.get(licitacion.name, {}),
                    resultado,
                    todos,
                    consumos,
                    resumenes,
                    archivos_ocr,
                )
            if not oferentes:
                print(f"{licitacion.name}: ya procesada ({len(resultados)} proveedores)")
                continue
            print(f"{licitacion.name}: reanudando; faltan {len(oferentes)} proveedores")

        if args.limite_proveedores is not None:
            if args.limite_proveedores < 1:
                sys.exit("--limite-proveedores debe ser mayor que cero")
            oferentes = oferentes[:args.limite_proveedores]

        print(f"\n{licitacion.name}: {len(oferentes)} proveedores")
        for carpeta in oferentes:
            info_carpeta = cargar_json(carpeta / "oferta.json")
            escribir_log(
                ruta_log,
                "inicio_proveedor",
                licitacion=licitacion.name,
                proveedor=info_carpeta.get("proveedor", carpeta.name),
                rut=info_carpeta.get("rut", ""),
            )
            try:
                resultado = procesar_oferta(carpeta, args, seven_zip)
            except RuntimeError as exc:
                print(f"\nERROR FATAL: {exc}")
                print(f"Se detiene la corrida porque {args.backend} dejo de responder.")
                escribir_log(
                    ruta_log,
                    "error_proveedor",
                    licitacion=licitacion.name,
                    proveedor=info_carpeta.get("proveedor", carpeta.name),
                    rut=info_carpeta.get("rut", ""),
                    error=str(exc),
                )
                detener = True
                break
            except Exception as exc:
                print(f"\nERROR EN PROVEEDOR {carpeta.name}: {exc}")
                escribir_log(
                    ruta_log,
                    "error_proveedor",
                    licitacion=licitacion.name,
                    proveedor=info_carpeta.get("proveedor", carpeta.name),
                    rut=info_carpeta.get("rut", ""),
                    error=str(exc),
                )
                continue
            resultados.append(resultado)
            contexto = metadata_licitaciones.get(licitacion.name, {})
            registrar_resultado(
                licitacion.name,
                contexto,
                resultado,
                todos,
                consumos,
                resumenes,
                archivos_ocr,
            )
            resumen_actual = resumenes[-1]
            escribir_log(
                ruta_log,
                "fin_proveedor",
                licitacion=licitacion.name,
                proveedor=resultado["proveedor"],
                rut=resultado["rut"],
                productos=resumen_actual["productos"],
                productos_validos=resumen_actual["productos_validos"],
                errores_modelo=resumen_actual["errores_modelo"],
                ocr=resumen_actual["ocr"],
            )
            print(
                f"  {resultado['proveedor'][:34]:34} "
                f"parciales={len(resultado['productos_parciales']):2} "
                f"finales={len(resultado['productos']):2} "
                f"ocr={len(resultado['necesita_ocr']):2} "
                f"consol={resultado['estado_consolidacion']}"
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
            escribir_json_atomico(salida, resultados)

        if resultados:
            escribir_json_atomico(salida, resultados)
        if detener:
            break
        if args.pausa_licitacion > 0 and licitacion != licitaciones[-1]:
            print(f"Pausa de {args.pausa_licitacion} segundos para enfriar el equipo...")
            time.sleep(args.pausa_licitacion)

    if todos or resumenes:
        generar_excel(todos, consumos, resumenes, ruta_excel)

    print("\n" + "=" * 74)
    print(f"Productos consolidados: {len(todos)}")
    print(f"Productos validados OK: {sum(p.get('estado_validacion') == 'ok' for p in todos)}")
    print(f"Excel: {ruta_excel if todos or resumenes else 'no generado'}")
    print(f"Log: {ruta_log}")
    print(f"Archivos que requieren OCR: {len(archivos_ocr)}")
    for archivo_ocr in archivos_ocr:
        print(
            f"  - {archivo_ocr['licitacion']} | "
            f"{archivo_ocr['proveedor']} | {archivo_ocr['archivo']}"
        )
    print("=" * 74)
    escribir_log(
        ruta_log,
        "fin_corrida",
        productos=len(todos),
        productos_validos=sum(producto.get("estado_validacion") == "ok" for producto in todos),
        archivos_ocr=len(archivos_ocr),
        detenida=detener,
    )


if __name__ == "__main__":
    main()

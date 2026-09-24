#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extractor IA local para ofertas de Mercado Publico.

Flujo:
  1. Extrae el texto de CADA archivo por separado.
    2. Preserva tablas y divide documentos largos sin perder filas.
    3. Consulta Ollama o LM Studio; con --paralelo N procesa N proveedores a la vez.
       Con --vision, las paginas PDF con tablas dificiles o escaneadas se envian
       como imagen (+ su texto) a un modelo con vision.
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
import base64
import logging
import csv
import importlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import zipfile
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait

# pdfminer (usado por pdfplumber) avisa "Could not get FontBBox..." cuando una fuente del PDF
# no declara sus dimensiones; usa valores por defecto y sigue leyendo. Solo ensucia la consola.
logging.getLogger("pdfminer").setLevel(logging.ERROR)
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

# Se pueden cambiar con variables de entorno (ej. LM Studio en otro puerto o en otro equipo).
OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://127.0.0.1:11434/api/generate")
LMSTUDIO_URL = os.environ.get("LMSTUDIO_URL", "http://127.0.0.1:1234/v1/chat/completions")
DEBUG = False
ESTADOS_IA_OK = {"ok", "ok_reparado", "ok_extraido"}
# PyMuPDF no es seguro entre hilos: todo uso de fitz pasa por este lock.
# Solo protege abrir/renderizar PDFs (rapido); las llamadas al modelo quedan fuera.
FITZ_LOCK = threading.Lock()
LOG_LOCK = threading.Lock()  # el log JSONL se escribe desde varios hilos
# Senal de detencion para los hilos de trabajo (Ctrl+C o backend caido). Los hilos
# la revisan antes de cada llamada al modelo y antes de cada archivo.
DETENER = threading.Event()


class CorridaDetenida(Exception):
    """Un hilo abandono su proveedor porque la corrida se esta deteniendo."""


def verificar_detencion():
    if DETENER.is_set():
        raise CorridaDetenida("corrida detenida")
EXTENSIONES_IMAGEN = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}
PATRON_PALABRA_PRECIO = re.compile(r"precio|valor|unitario|total|neto|subtotal|monto", re.I)
PATRON_NUMERO = re.compile(r"\d+(?:[.,]\d+)*")
PATRON_EQUIPO_GENERICO = re.compile(r"\bequipos?\b", re.I)
PATRON_MONTO = re.compile(r"\$\s*[\d.]|\b\d{1,3}(?:\.\d{3})+(?:,\d+)?\b")

MARCAS = [
    "HP", "Lenovo", "Dell", "Asus", "Acer", "Apple", "Samsung", "Lexmark",
    "Epson", "Canon", "Brother", "MSI", "Huawei", "Kingston", "Logitech",
    "LG", "ThinkCentre", "ThinkPad", "ProOne", "ProDesk", "OptiPlex",
    "Latitude", "Pavilion"
]

from prompts_extraccion import (
        NOTA_IMAGENES_PROVEEDOR,
        PROMPT_ARCHIVO,
        PROMPT_CONSOLIDAR,
        PROMPT_PROVEEDOR,
        REGLAS_VISION,
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
        # Renderizar bajo lock (PyMuPDF no es thread-safe); el OCR corre fuera del lock.
        renders = []
        escala = dpi / 72
        with FITZ_LOCK:
            doc = fitz.open(path)
            try:
                for index, page in enumerate(doc):
                    if index >= max_paginas:
                        break
                    if paginas is not None and index not in paginas:
                        continue
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(escala, escala), alpha=False)
                    renders.append((index, pixmap.tobytes("png")))
            finally:
                doc.close()
        partes = []
        for index, png in renders:
            with Image.open(io.BytesIO(png)) as imagen:
                try:
                    texto = pytesseract.image_to_string(imagen, lang=idioma)
                except pytesseract.TesseractError:
                    texto = pytesseract.image_to_string(imagen, lang="eng")
            if texto.strip():
                partes.append(f"[PAGINA {index + 1} - OCR]\n{texto.strip()}")
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


def linea_tipo_fila(linea):
    """Linea que parece fila de tabla de precios: un monto y al menos otro numero
    (cantidad, item o segundo monto). Descarta lineas sueltas como 'Total $ 9.282.000'
    o 'Monto garantia $ 1.000.000'."""
    return bool(PATRON_MONTO.search(linea)) and len(PATRON_NUMERO.findall(linea)) >= 2


def pagina_en_alcance(texto_pagina):
    """La pagina menciona algun equipo del alcance comercial."""
    if PATRON_EQUIPO_RELEVANTE.search(texto_pagina) or PATRON_MONITOR.search(texto_pagina) \
            or PATRON_IMPRESORA.search(texto_pagina):
        return True
    # "equipo" generico cuenta solo si la pagina no habla de TV.
    return bool(PATRON_EQUIPO_GENERICO.search(texto_pagina)) and not PATRON_TV.search(texto_pagina)


def pagina_requiere_vision(texto_pagina, tablas_validas, tablas_colapsadas, min_filas=2):
    """Decide si una pagina conviene leerla como imagen en vez de solo texto.

    - sin_texto: pagina escaneada (sin imagen no hay nada que leer).
    - tabla_colapsada: pdfplumber detecto una tabla sin columnas recuperables, y la
      pagina menciona un equipo del alcance y precios.
    - tabla_sin_bordes: sin tablas detectadas, pero con al menos `min_filas` lineas
      que parecen filas de precios, una palabra de precio y un equipo del alcance.

    Las paginas que no califican NO se pierden: siguen por la ruta de texto.
    """
    if len(texto_pagina) <= 40 and not tablas_validas:
        return "sin_texto"
    if not pagina_en_alcance(texto_pagina):
        return None
    habla_de_precios = bool(PATRON_PALABRA_PRECIO.search(texto_pagina))
    if tablas_colapsadas:
        if habla_de_precios or PATRON_MONTO.search(texto_pagina):
            return "tabla_colapsada"
        return None
    if not tablas_validas and habla_de_precios:
        filas = sum(1 for linea in texto_pagina.splitlines() if linea_tipo_fila(linea))
        if filas >= min_filas:
            return "tabla_sin_bordes"
    return None


def extraer_pdf(path, max_paginas, usar_ocr=False, tesseract_cmd=None, idioma_ocr="spa+eng", dpi_ocr=180,
                detalle=None, min_filas_vision=2):
    """Extrae texto y tablas de un PDF.

    Si se entrega `detalle` (dict), se llena con informacion por pagina para el
    modo vision: texto_paginas {indice: texto} y paginas_vision {indice: motivo}.
    """
    try:
        partes = []
        paginas_sin_texto = []
        texto_paginas = {}
        paginas_vision = {}
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

                partes_pagina = [f"[PAGINA {index + 1}]"]
                if texto_pagina:
                    partes_pagina.append(texto_pagina)
                if tablas_colapsadas:
                    partes_pagina.append(f"[AVISO: {tablas_colapsadas} tabla(s) sin columnas recuperables]")
                for num_tabla, (_, filas) in enumerate(tablas_validas, 1):
                    partes_pagina.append(f"[TABLA {num_tabla} - pagina {index + 1}]")
                    for num_fila, fila in enumerate(filas, 1):
                        valores = [(celda or "").strip().replace("\n", " ") for celda in fila]
                        if any(valores):
                            partes_pagina.append(f"FILA {num_fila}: " + " || ".join(valores))
                partes.extend(partes_pagina)
                texto_paginas[index] = limpiar_control("\n".join(partes_pagina)).strip()
                motivo_vision = pagina_requiere_vision(
                    texto_pagina, tablas_validas, tablas_colapsadas, min_filas_vision
                )
                if motivo_vision:
                    paginas_vision[index] = motivo_vision
                if len(texto_pagina) <= 40 and not tablas_validas:
                    paginas_sin_texto.append(index)
        texto = limpiar_control("\n".join(partes)).strip()
    except Exception:
        texto = ""
        paginas_sin_texto = []
        texto_paginas = {}
        paginas_vision = {}

    if detalle is not None:
        detalle["texto_paginas"] = texto_paginas
        detalle["paginas_vision"] = paginas_vision
        if paginas_vision:
            # En modo vision no hace falta OCR ni respaldo: las paginas dificiles
            # se leen como imagen y el resto ya tiene texto estructurado.
            return texto, "texto_vision"

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
        partes = []
        with FITZ_LOCK:
            doc = fitz.open(path)
            try:
                for index, page in enumerate(doc):
                    if index >= max_paginas:
                        break
                    partes.append(page.get_text("text"))
            finally:
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


def extraer_archivo(path, args, detalle=None):
    extension = path.suffix.lower()
    if extension == ".pdf":
        return extraer_pdf(
            path,
            args.max_paginas,
            usar_ocr=args.ocr,
            tesseract_cmd=args.tesseract_cmd,
            idioma_ocr=args.idioma_ocr,
            dpi_ocr=args.dpi_ocr,
            detalle=detalle,
            min_filas_vision=getattr(args, "vision_min_filas", 2),
        )
    if extension in EXTENSIONES_IMAGEN:
        return "", "imagen" if getattr(args, "vision", False) else "no_soportado"
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


PATRON_ERROR_TRANSITORIO = re.compile(r"channel\s*error|model\s+(?:is\s+)?(?:loading|unloaded)|"
                                      r"no\s+models?\s+loaded|connection\s+(?:reset|closed)", re.I)


def es_error_transitorio(texto):
    """Errores de LM Studio que suelen resolverse reintentando (canal interno o LM Link
    recien conectado, modelo cargandose)."""
    return bool(PATRON_ERROR_TRANSITORIO.search(str(texto or "")))


def calentar_modelo(args, intentos=4, espera=10):
    """Peticion minima antes de la carga real: activa el enlace (LM Link) y confirma que
    el modelo responde. Devuelve los segundos que tardo o None si no respondio."""
    for intento in range(1, intentos + 1):
        inicio = time.perf_counter()
        try:
            _, estado, _, _ = consultar_modelo(
                'Responde solo {"ok": true}', args.modelo, args.timeout, args.num_ctx, args.backend,
                32, 0, 0)
        except RuntimeError as exc:
            estado = str(exc)
        if estado in ESTADOS_IA_OK:
            return round(time.perf_counter() - inicio, 1)
        print(f"  El modelo aun no responde ({estado[:120]}); reintento {intento}/{intentos} en {espera} s...")
        time.sleep(espera)
    return None


def consultar_modelo(
    prompt,
    modelo,
    timeout,
    num_ctx,
    backend="ollama",
    max_tokens=4096,
    reintentos=2,
    espera_reintento=5,
    imagenes=None,
):
    """imagenes: lista opcional de dicts {"mime": ..., "base64": ...} para modelos con vision."""
    verificar_detencion()
    if backend == "lmstudio":
        if imagenes:
            contenido_usuario = [
                {"type": "image_url", "image_url": {"url": f"data:{img['mime']};base64,{img['base64']}"}}
                for img in imagenes
            ] + [{"type": "text", "text": prompt}]
        else:
            contenido_usuario = prompt
        payload = {
            "model": modelo,
            "messages": [
                {"role": "system", "content": "Devuelve exclusivamente un objeto JSON valido."},
                {"role": "user", "content": contenido_usuario}
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
        if imagenes:
            payload["images"] = [img["base64"] for img in imagenes]
        url = OLLAMA_URL
    ultimo_error = None
    for intento in range(reintentos + 1):
        try:
            respuesta = requests.post(url, json=payload, timeout=timeout)
            if respuesta.status_code >= 400:
                estado = f"http_{respuesta.status_code}: {respuesta.text[:500]}"
                transitorio = respuesta.status_code in (429, 500, 502, 503, 504) or es_error_transitorio(respuesta.text)
                if not transitorio or intento >= reintentos:
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
# Accesorios y servicios que nombran un equipo pero no lo son:
# "Estuche de cuero para MacBook Pro", "FortiSIEM All-In-One Subscription License".
PATRON_ACCESORIO_INICIAL = re.compile(
    r"^\W*(?:n.?\s*\d+\s+)?(?:\d+\W+)?(?:estuche|funda|malet[ií]n|mochila|bolso|cargador|adaptador|protector|"
    r"l[aá]mina|soporte|base\s+refrigerante|candado|hub|kit|lic(?:encia|ense)|renovaci[oó]n|"
    r"extensi[oó]n\s+de\s+garant[ií]a|garant[ií]a)\b", re.I)
PATRON_SERVICIO_SOFTWARE = re.compile(r"subscription|suscripci[oó]n|\bsaas\b|per\s+device|por\s+dispositivo", re.I)
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
    descripcion = str(producto.get("producto") or "")
    if PATRON_ACCESORIO_INICIAL.search(descripcion) or PATRON_SERVICIO_SOFTWARE.search(descripcion):
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


PATRON_TIPO_AIO = re.compile(r"all.?in.?one|\baio\b|todo\s+en\s+uno|proone|eliteone|\bimac\b|pavilion\s+\d{2}-|\b\d{2}-cr\d", re.I)
PATRON_TIPO_WORKSTATION = re.compile(r"workstation|estaci[oó]n\s+de\s+trabajo|thinkstation|\bz[248]\s+g\d", re.I)
PATRON_TIPO_NOTEBOOK = re.compile(
    r"notebook|laptop|port[aá]til|ultrabook|chromebook|macbook|probook|elitebook|zbook|thinkpad|thinkbook|"
    r"ideapad|latitude|vostro|vivobook|zenbook|expertbook|travelmate|\b1[45]-[a-z]{2}\d", re.I)
PATRON_TIPO_DESKTOP = re.compile(
    r"desktop|escritorio|\btorre\b|\bpc\b|\bsff\b|mini\s*pc|prodesk|elitedesk|optiplex|thinkcentre|"
    r"veriton|expertcenter|mac\s*mini|computador", re.I)


def id_revision(licitacion, rut, producto):
    """Identificador estable de un producto para la revision: depende de su contenido,
    no de su posicion, asi las decisiones guardadas no se cruzan si cambia el orden."""
    import hashlib
    clave = "|".join(str(producto.get(c) or "") for c in ("producto", "precio_unitario", "cantidad", "archivo_fuente"))
    return f"{licitacion}__{rut}__{hashlib.sha1(clave.encode('utf-8')).hexdigest()[:8]}"


def subcategoria_producto(producto):
    """Tipo de equipo para analisis: notebook, all-in-one, desktop, workstation,
    monitor o impresora. 'equipo (sin especificar)' si no se puede distinguir."""
    categoria = producto.get("categoria") or clasificar_producto(producto)
    if categoria == "monitor":
        return "monitor"
    if categoria == "impresora":
        return "impresora"
    texto = " ".join(str(producto.get(c) or "") for c in ("producto", "modelo"))
    for tipo, patron in (("all-in-one", PATRON_TIPO_AIO), ("workstation", PATRON_TIPO_WORKSTATION),
                         ("notebook", PATRON_TIPO_NOTEBOOK), ("desktop", PATRON_TIPO_DESKTOP)):
        if patron.search(texto):
            return tipo
    return "equipo (sin especificar)" if categoria == "equipo" else None


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


# ---------------------------------------------------------------------------
# MODO VISION (--vision)
# ---------------------------------------------------------------------------
def imagen_desde_pixmap(pixmap):
    """JPEG calidad 85: una pagina escaneada pesa ~6 veces menos que en PNG y el modelo
    la lee igual. PNG solo si la version de PyMuPDF no soporta JPEG."""
    try:
        return {"mime": "image/jpeg", "base64": base64.b64encode(pixmap.tobytes("jpg", jpg_quality=85)).decode("ascii")}
    except Exception:
        return {"mime": "image/png", "base64": base64.b64encode(pixmap.tobytes("png")).decode("ascii")}


def renderizar_paginas_pdf(path, indices, dpi):
    """Devuelve {indice: {"mime", "base64"}} con las paginas pedidas como imagen JPEG."""
    imagenes = {}
    escala = dpi / 72
    with FITZ_LOCK:
        doc = fitz.open(path)
        try:
            for index in indices:
                if 0 <= index < len(doc):
                    pixmap = doc[index].get_pixmap(matrix=fitz.Matrix(escala, escala), alpha=False)
                    imagenes[index] = imagen_desde_pixmap(pixmap)
        finally:
            doc.close()
    return imagenes


def cargar_imagen_archivo(path, lado_maximo=1600):
    """Carga un JPG/PNG/WEBP adjunto. Reduce fotos muy grandes para no gastar contexto."""
    datos = path.read_bytes()
    mime = EXTENSIONES_IMAGEN[path.suffix.lower()]
    if Image is not None:
        try:
            with Image.open(io.BytesIO(datos)) as original:
                if max(original.size) > lado_maximo or mime == "image/webp":
                    imagen = original.convert("RGB")
                    imagen.thumbnail((lado_maximo, lado_maximo))
                    buffer = io.BytesIO()
                    imagen.save(buffer, format="JPEG", quality=85)
                    datos, mime = buffer.getvalue(), "image/jpeg"
        except Exception:
            pass
    return {"mime": mime, "base64": base64.b64encode(datos).decode("ascii")}


def sumar_consumos(consumos, modelo):
    """Como combinar_consumos, pero suma llamadas de consumos ya combinados."""
    consumos = [consumo for consumo in consumos if consumo]
    return {
        "modelo": next((fila.get("modelo") for fila in consumos if fila.get("modelo")), modelo),
        "prompt_tokens": sum(fila.get("prompt_tokens") or 0 for fila in consumos),
        "completion_tokens": sum(fila.get("completion_tokens") or 0 for fila in consumos),
        "total_tokens": sum(fila.get("total_tokens") or 0 for fila in consumos),
        "llamadas": sum(int(fila.get("llamadas") or 0) for fila in consumos),
    }


def extraer_parciales_vision(path, paginas, proveedor, rut, total_oferta, args):
    """Envia paginas como imagen (+ su texto) en grupos de --vision-paginas-por-llamada.

    paginas: lista de dicts {"pagina": numero_1_based, "imagen": {...}, "texto": str}
    Devuelve la misma tupla que extraer_parciales_archivo.
    """
    relativo = str(path)
    tipo_doc = tipo_documental(path.name)
    por_llamada = max(1, args.vision_paginas_por_llamada)
    grupos = [paginas[inicio:inicio + por_llamada] for inicio in range(0, len(paginas), por_llamada)]
    productos = []
    estados = []
    respuestas = []
    prompts = []
    observaciones = []
    consumos = []
    for numero, grupo in enumerate(grupos, 1):
        numeros = ", ".join(str(pagina["pagina"]) for pagina in grupo)
        textos = "\n\n".join(pagina["texto"] for pagina in grupo if pagina["texto"].strip())
        texto_prompt = (
            f"{REGLAS_VISION}\n\n"
            f"[GRUPO {numero} DE {len(grupos)}] IMAGENES ADJUNTAS, EN ESTE ORDEN: paginas {numeros}\n"
            "TEXTO EXTRAIDO AUTOMATICAMENTE DE ESAS PAGINAS "
            "(puede venir desordenado; usalo solo para confirmar cifras y nombres):\n"
            f"{textos or '(sin texto extraible: paginas escaneadas o imagen)'}"
        )
        prompt = PROMPT_ARCHIVO.format(
            proveedor=proveedor,
            rut=rut,
            archivo=relativo,
            tipo_documental=tipo_doc,
            total_oferta=total_oferta or "no informado",
            pistas_producto="(usar las imagenes adjuntas)",
            pistas_precio="(usar las imagenes adjuntas)",
            texto=texto_prompt,
        )
        datos, estado, cruda, consumo = consultar_modelo(
            prompt, args.modelo, args.timeout, args.num_ctx, args.backend, args.max_tokens,
            args.reintentos_modelo, args.espera_reintento,
            imagenes=[pagina["imagen"] for pagina in grupo],
        )
        estados.append(estado)
        respuestas.append(f"[VISION GRUPO {numero} - paginas {numeros}]\n{cruda}")
        prompts.append(f"[VISION - {len(grupo)} imagen(es): paginas {numeros}]\n{prompt}")
        consumos.append(consumo)
        if isinstance(datos, dict):
            if datos.get("observaciones"):
                observaciones.append(str(datos["observaciones"]))
            for producto in datos.get("productos", []):
                if isinstance(producto, dict):
                    limpio = normalizar_producto(producto, proveedor, rut, relativo, tipo_doc)
                    if limpio:
                        limpio["lectura"] = "vision"
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


def extraer_parciales_mixto(path, relativo, detalle, proveedor, rut, total_oferta, args):
    """Archivo con paginas para vision: texto normal para las paginas faciles
    y vision para las dificiles. Para imagenes sueltas (JPG/PNG) solo vision."""
    if path.suffix.lower() in EXTENSIONES_IMAGEN:
        paginas = [{"pagina": 1, "imagen": cargar_imagen_archivo(path), "texto": ""}]
        texto_resto = ""
    else:
        texto_paginas = detalle.get("texto_paginas", {})
        indices_vision = sorted(detalle.get("paginas_vision", {}))
        imagenes = renderizar_paginas_pdf(path, indices_vision, args.dpi_vision)
        paginas = [
            {"pagina": index + 1, "imagen": imagenes[index], "texto": texto_paginas.get(index, "")}
            for index in indices_vision if index in imagenes
        ]
        texto_resto = "\n".join(
            texto_paginas[index] for index in sorted(texto_paginas) if index not in set(indices_vision)
        ).strip()

    resultados = []
    if len(texto_resto) > 40:
        resultados.append(extraer_parciales_archivo(Path(relativo), texto_resto, proveedor, rut, total_oferta, args))
    if paginas:
        resultados.append(extraer_parciales_vision(Path(relativo), paginas, proveedor, rut, total_oferta, args))
    if not resultados:
        return [], "no_ejecutada", "", "", None, {}

    productos = [producto for resultado in resultados for producto in resultado[0]]
    estados = [resultado[1] for resultado in resultados]
    estado_final = "ok" if all(estado in ESTADOS_IA_OK for estado in estados) else "; ".join(estados)
    observaciones = "; ".join(resultado[4] for resultado in resultados if resultado[4]) or None
    return (
        productos,
        estado_final,
        "\n\n".join(resultado[2] for resultado in resultados),
        "\n\n".join(resultado[3] for resultado in resultados),
        observaciones,
        sumar_consumos([resultado[5] for resultado in resultados], args.modelo),
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


def reemplazar_resultado(resultados, nuevo):
    """Reemplaza en la lista el resultado del mismo RUT (o lo agrega si no existe)."""
    rut = str(nuevo.get("rut") or "").strip()
    for indice, resultado in enumerate(resultados):
        if str(resultado.get("rut") or "").strip() == rut:
            resultados[indice] = nuevo
            return
    resultados.append(nuevo)


def listar_archivos_oferta(carpeta):
    return [
        path for path in carpeta.rglob("*")
        if path.is_file()
        and path.name not in {"oferta.json", "extraccion_ia.json", ".extraido_ok"}
        and "_debug_ia" not in path.parts
    ]


# ---------------------------------------------------------------------------
# REPROCESO SELECTIVO: solo los archivos que lo necesitan
# ---------------------------------------------------------------------------
def archivo_fallo_modelo(registro):
    estado = str(registro.get("estado_ia") or "")
    return estado not in ("no_ejecutada", "solo_texto") and estado not in ESTADOS_IA_OK


def archivo_pendiente_ocr(registro, previo):
    return registro.get("archivo") in set(previo.get("necesita_ocr") or [])


def revision_necesaria(resultado, carpeta, args):
    """Chequeo rapido (sin leer PDFs) de si un proveedor ya procesado puede
    necesitar trabajo. El detalle de QUE archivos rehacer lo decide
    planificar_reproceso dentro del hilo de trabajo."""
    if resultado_requiere_reproceso(resultado, args.ocr, args.modelo, args.backend, args.vision,
                                    modo_extraccion(args), args.vision_solo_escaneadas):
        return True
    # Archivos que aparecieron despues (ej. anexos tecnicos descargados mas tarde).
    registrados = {registro.get("archivo") for registro in resultado.get("resultados_archivos", [])}
    return any(
        str(path.relative_to(carpeta)) not in registrados
        for path in listar_archivos_oferta(carpeta)
    )


def planificar_reproceso(carpeta, previo, args, seven_zip):
    """Decide que hacer con un proveedor ya procesado.

    Devuelve (modo, archivos):
      - ("completo", None): cambio de modelo/backend; se rehace todo.
      - ("parcial", {archivos}): se rehacen solo esos archivos y se reutiliza el
        resto; luego se vuelve a consolidar el proveedor.
      - ("consolidar", set()): solo fallo la consolidacion.
      - ("nada", set()): no hay nada que rehacer.
    """
    cambio_modelo = bool(previo.get("modelo") and str(previo.get("modelo")) != str(args.modelo))
    cambio_backend = bool(previo.get("backend") and str(previo.get("backend")) != str(args.backend))
    cambio_modo = (previo.get("modo") or "por_archivo") != modo_extraccion(args)
    if cambio_modelo or cambio_backend or cambio_modo:
        return "completo", None

    descomprimir(carpeta, seven_zip)
    previos = {registro.get("archivo"): registro for registro in previo.get("resultados_archivos", [])}
    revisar_vision = args.vision and not previo.get("vision") and not args.vision_solo_escaneadas
    rehacer = {}
    for path in listar_archivos_oferta(carpeta):
        verificar_detencion()
        relativo = str(path.relative_to(carpeta))
        registro = previos.get(relativo)
        extension = path.suffix.lower()
        if registro is None:
            rehacer[relativo] = "archivo_nuevo"
        elif archivo_fallo_modelo(registro):
            rehacer[relativo] = "fallo_modelo"
        elif (args.ocr or args.vision) and archivo_pendiente_ocr(registro, previo):
            rehacer[relativo] = "pendiente_ocr"
        elif revisar_vision and not args.solo_texto:
            if extension in EXTENSIONES_IMAGEN:
                rehacer[relativo] = "imagen"
            elif extension == ".pdf" and registro.get("estado_lectura") != "texto_vision":
                # Lectura rapida sin modelo: solo clasifica las paginas.
                detalle = {}
                extraer_archivo(path, args, detalle)
                if detalle.get("paginas_vision"):
                    rehacer[relativo] = "paginas_para_vision"

    if rehacer:
        return "parcial", rehacer
    estado_consolidacion = str(previo.get("estado_consolidacion") or "")
    if estado_consolidacion not in ESTADOS_IA_OK | {"", "omitida", "no_ejecutada"}:
        return "consolidar", {}
    return "nada", {}


def modo_extraccion(args):
    return "por_proveedor" if getattr(args, "por_proveedor", False) else "por_archivo"


# ---------------------------------------------------------------------------
# MODO POR PROVEEDOR (--por-proveedor)
# ---------------------------------------------------------------------------
PATRON_PALABRA_PRODUCTO = re.compile(
    PATRON_EQUIPO_RELEVANTE.pattern + r"|monitor|impresora|multifuncional|all.?in.?one", re.I)
PATRON_MARCA_PAGINA = re.compile(r"\b(?:" + "|".join(re.escape(m) for m in MARCAS) + r")\b", re.I)


def compactar_texto(texto):
    """Quita el relleno de la extraccion con diseno: espacios al final de linea, rachas
    largas de espacios y lineas vacias repetidas. Conserva la separacion entre columnas."""
    lineas = [re.sub(r" {4,}", "   ", linea.rstrip()) for linea in texto.splitlines()]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lineas)).strip()


def pagina_relevante(texto):
    """Una pagina se envia si tiene montos, o si nombra un equipo junto con una marca
    (anexo tecnico). Declaraciones, bases y formularios administrativos quedan fuera."""
    if PATRON_MONTO.search(texto):
        return True
    return bool(PATRON_PALABRA_PRODUCTO.search(texto) and PATRON_MARCA_PAGINA.search(texto))


def lineas_relevantes(texto, limite):
    """Para planillas y documentos largos sin paginas: conserva encabezados y lineas
    con montos o equipos hasta el limite de caracteres."""
    if len(texto) <= limite:
        return texto
    lineas = texto.splitlines()
    elegidas = [l for i, l in enumerate(lineas) if i < 15 or PATRON_MONTO.search(l) or PATRON_PALABRA_PRODUCTO.search(l)]
    return "\n".join(elegidas)[:limite]


def preparar_bloques_proveedor(carpeta, archivos, args):
    """Lee todos los anexos y devuelve (bloques, registros, necesita_ocr, no_soportados, marcas).
    Cada bloque es una pagina (o un archivo sin paginas) con su texto y, si corresponde, su imagen."""
    bloques, registros, necesita_ocr, no_soportados = [], [], [], []
    descartados = []  # paginas con texto que el filtro no envio (respaldo)
    marcas = set()
    limite_archivo = max(4000, args.max_chars_proveedor // 2)
    for path in sorted(archivos, key=prioridad):
        verificar_detencion()
        relativo = str(path.relative_to(carpeta))
        extension = path.suffix.lower()
        registro = {"archivo": relativo, "tipo": extension, "tipo_documental": tipo_documental(path.name),
                    "estado_lectura": "", "caracteres": 0, "estado_ia": "no_ejecutada",
                    "productos_encontrados": 0, "observaciones": None, "paginas_enviadas": []}
        registros.append(registro)

        if extension in EXTENSIONES_IMAGEN:
            registro["estado_lectura"] = "imagen"
            if args.vision:
                bloques.append({"archivo": relativo, "pagina": 1, "texto": "", "imagen": cargar_imagen_archivo(path)})
                registro["paginas_enviadas"] = [1]
            else:
                no_soportados.append(relativo)
            continue

        if extension == ".pdf":
            detalle = {}
            texto, estado = extraer_archivo(path, args, detalle)
            registro["estado_lectura"] = estado
            registro["caracteres"] = len(texto)
            texto_paginas = detalle.get("texto_paginas") or ({0: texto} if texto.strip() else {})
            paginas_vision = detalle.get("paginas_vision") or {}
            if paginas_vision and args.vision:
                registro["paginas_vision"] = {str(i + 1): m for i, m in sorted(paginas_vision.items())}
            escaneadas = sorted(i for i, motivo in paginas_vision.items() if motivo == "sin_texto")
            imagenes = {}
            if args.vision and paginas_vision and not args.vision_solo_escaneadas:
                imagenes = renderizar_paginas_pdf(path, sorted(paginas_vision), args.dpi_vision)
            else:
                sin_leer = list(escaneadas)
                if escaneadas and args.ocr:
                    sin_leer = []
                    for indice in escaneadas:
                        texto_ocr, _ = extraer_pdf_ocr(path, args.max_paginas, args.tesseract_cmd, args.idioma_ocr,
                                                       args.dpi_ocr, paginas={indice})
                        texto_paginas[indice] = texto_ocr
                        if len(texto_ocr.strip()) <= 40:
                            sin_leer.append(indice)  # el OCR no pudo leer la pagina
                if sin_leer and args.vision:
                    imagenes = renderizar_paginas_pdf(path, sin_leer, args.dpi_vision)
                    registro["paginas_vision"] = {str(i + 1): "sin_texto" for i in sin_leer}
                elif sin_leer and relativo not in necesita_ocr:
                    necesita_ocr.append(relativo)
            for indice in sorted(set(texto_paginas) | set(imagenes)):
                texto_pagina = texto_paginas.get(indice, "")
                for marca in MARCAS:
                    if re.search(rf"\b{re.escape(marca)}\b", texto_pagina, re.I):
                        marcas.add(marca)
                bloque = {"archivo": relativo, "pagina": indice + 1,
                          "texto": compactar_texto(texto_pagina)[:args.max_chars_proveedor],
                          "imagen": imagenes.get(indice)}
                if indice in imagenes or pagina_relevante(texto_pagina):
                    bloques.append(bloque)
                    registro["paginas_enviadas"].append(indice + 1)
                elif bloque["texto"].strip():
                    descartados.append(bloque)
            continue

        texto, estado = extraer_archivo(path, args)
        registro["estado_lectura"] = estado
        registro["caracteres"] = len(texto)
        if estado in ("no_soportado", "xls_legacy_no_soportado"):
            no_soportados.append(relativo)
            continue
        for marca in MARCAS:
            if re.search(rf"\b{re.escape(marca)}\b", texto, re.I):
                marcas.add(marca)
        if texto.strip():
            bloque = {"archivo": relativo, "pagina": None,
                      "texto": lineas_relevantes(compactar_texto(texto), limite_archivo), "imagen": None}
            if PATRON_MONTO.search(texto) or PATRON_PALABRA_PRODUCTO.search(texto):
                bloques.append(bloque)
                registro["paginas_enviadas"] = ["todo"]
            else:
                descartados.append(bloque)

    if not bloques and descartados:
        # El filtro descarto todo: en vez de saltar al proveedor, se envian sus anexos
        # (economicos primero, ya vienen ordenados) hasta el limite de una llamada.
        usados = 0
        for bloque in descartados:
            if usados + len(bloque["texto"]) > args.max_chars_proveedor:
                break
            bloques.append(bloque)
            usados += len(bloque["texto"])
            registro = next(r for r in registros if r["archivo"] == bloque["archivo"])
            registro["paginas_enviadas"].append(bloque["pagina"] or "todo")
            registro["respaldo_sin_filtro"] = True
    return bloques, registros, necesita_ocr, no_soportados, marcas


def agrupar_bloques(bloques, max_chars, max_imagenes):
    """Reparte los bloques en llamadas respetando el limite de texto y de imagenes.
    Casi todos los proveedores caben en una sola llamada."""
    grupos, actual, chars, imagenes = [], [], 0, 0
    for bloque in bloques:
        largo = len(bloque["texto"]) + 80
        con_imagen = 1 if bloque.get("imagen") else 0
        if actual and (chars + largo > max_chars or imagenes + con_imagen > max_imagenes):
            grupos.append(actual)
            actual, chars, imagenes = [], 0, 0
        actual.append(bloque)
        chars += largo
        imagenes += con_imagen
    if actual:
        grupos.append(actual)
    return grupos


def archivo_declarado(nombre, archivos_llamada):
    """Asocia el 'archivo' que devuelve el modelo con un archivo real del proveedor."""
    nombre = str(nombre or "").strip().lower()
    if not nombre:
        return None
    for relativo in archivos_llamada:
        if nombre == relativo.lower() or nombre == Path(relativo).name.lower():
            return relativo
    for relativo in archivos_llamada:
        base = Path(relativo).stem.lower()
        if base and (base in nombre or nombre in relativo.lower()):
            return relativo
    return None


def procesar_oferta_por_proveedor(carpeta, args, seven_zip):
    """Una llamada (o pocas) por proveedor con las paginas relevantes de todos sus anexos."""
    inicio = time.perf_counter()
    info = cargar_json(carpeta / "oferta.json")
    proveedor = info.get("proveedor") or carpeta.name
    rut = info.get("rut") or carpeta.name.split("__", 1)[0]
    total_oferta = info.get("total") or info.get("total_oferta") or ""
    archivos_comprimidos = descomprimir(carpeta, seven_zip)
    bloques, registros, necesita_ocr, no_soportados, marcas = preparar_bloques_proveedor(
        carpeta, listar_archivos_oferta(carpeta), args)
    por_archivo = {registro["archivo"]: registro for registro in registros}

    debug_dir = carpeta / "_debug_ia"
    if args.debug:
        debug_dir.mkdir(exist_ok=True)

    parciales, estados, llamadas = [], [], 0
    grupos = [] if args.solo_texto else agrupar_bloques(bloques, args.max_chars_proveedor, args.max_imagenes_proveedor)
    for numero, grupo in enumerate(grupos, 1):
        verificar_detencion()
        partes, imagenes, lista_imagenes = [], [], []
        for bloque in grupo:
            etiqueta = f"=== ARCHIVO: {bloque['archivo']} ({tipo_documental(bloque['archivo'])})" + (
                f" | PAGINA {bloque['pagina']}" if bloque["pagina"] else "") + " ==="
            if bloque.get("imagen"):
                imagenes.append(bloque["imagen"])
                lista_imagenes.append(f"{len(imagenes)}) {bloque['archivo']} pagina {bloque['pagina']}")
                etiqueta += f" [IMAGEN {len(imagenes)}]"
            partes.append(f"{etiqueta}\n{bloque['texto'] or '(sin texto: ver imagen)'}")
        nota = NOTA_IMAGENES_PROVEEDOR.format(lista_imagenes="\n".join(lista_imagenes)) if imagenes else ""
        prompt = PROMPT_PROVEEDOR.format(proveedor=proveedor, total_oferta=total_oferta or "no informado",
                                         nota_imagenes=nota, documentos="\n\n".join(partes))
        datos, estado, cruda, consumo = consultar_modelo(
            prompt, args.modelo, args.timeout, args.num_ctx, args.backend, args.max_tokens,
            args.reintentos_modelo, args.espera_reintento, imagenes=imagenes or None)
        llamadas += 1
        estados.append(estado)
        archivos_llamada = list(dict.fromkeys(bloque["archivo"] for bloque in grupo))
        # el consumo se asigna al primer archivo de la llamada; los demas quedan con 0 llamadas
        for indice, relativo in enumerate(archivos_llamada):
            registro = por_archivo[relativo]
            registro["estado_ia"] = estado if registro["estado_ia"] in ("no_ejecutada", *ESTADOS_IA_OK) else registro["estado_ia"]
            registro["consumo_local"] = {**consumo, "llamadas": 1} if indice == 0 else {"llamadas": 0}
            registro["llamada_proveedor"] = numero
        if args.debug:
            (debug_dir / f"proveedor_{numero:02d}__prompt.txt").write_text(prompt, encoding="utf-8")
            (debug_dir / f"proveedor_{numero:02d}__respuesta.txt").write_text(cruda, encoding="utf-8")
        if isinstance(datos, dict):
            for producto in datos.get("productos", []):
                if not isinstance(producto, dict):
                    continue
                relativo = archivo_declarado(producto.get("archivo"), archivos_llamada) or archivos_llamada[0]
                if producto.get("precio_total") not in (None, "") and not producto.get("precio_total_tipo"):
                    producto["precio_total_tipo"] = "linea"
                limpio = normalizar_producto(producto, proveedor, rut, relativo, tipo_documental(relativo))
                if limpio:
                    parciales.append(limpio)
                    por_archivo[relativo]["productos_encontrados"] += 1
        if args.pausa_archivo > 0:
            time.sleep(args.pausa_archivo)

    relevantes = filtrar_productos_relevantes(parciales)
    if len(grupos) > 1 and not args.sin_consolidar and parciales:
        productos, estado_consolidacion, _, _, consumo_consolidacion = consolidar_parciales(
            parciales, proveedor, rut, total_oferta, args)
        llamadas += 1
    else:
        productos, estado_consolidacion, consumo_consolidacion = relevantes, "omitida", {}

    return {
        "proveedor": proveedor,
        "rut": rut,
        "modelo": args.modelo,
        "backend": args.backend,
        "vision": bool(args.vision),
        "modo": "por_proveedor",
        "estado_oferta": info.get("estado", ""),
        "total_oferta": total_oferta,
        "productos": productos,
        "productos_parciales": relevantes,
        "resultados_archivos": registros,
        "estado_consolidacion": estado_consolidacion,
        "consumo_consolidacion": consumo_consolidacion,
        "archivos_comprimidos": archivos_comprimidos,
        "necesita_ocr": necesita_ocr,
        "no_soportados": no_soportados,
        "marcas_detectadas_texto": sorted(marcas),
        "llamadas_modelo": llamadas,
        "segundos": round(time.perf_counter() - inicio, 1),
    }


def procesar_oferta(carpeta, args, seven_zip, previo=None, archivos_reprocesar=None):
    """Procesa un proveedor.

    Si se entregan `previo` y `archivos_reprocesar`, solo se leen y consultan al
    modelo esos archivos; los demas reutilizan su registro y productos parciales
    del resultado anterior. La consolidacion se vuelve a ejecutar siempre.
    """
    inicio_oferta = time.perf_counter()
    info = cargar_json(carpeta / "oferta.json")
    proveedor = info.get("proveedor") or carpeta.name
    rut = info.get("rut") or carpeta.name.split("__", 1)[0]
    total_oferta = info.get("total") or info.get("total_oferta") or ""
    archivos_comprimidos = descomprimir(carpeta, seven_zip)

    archivos = listar_archivos_oferta(carpeta)

    parciales = []
    resultados_archivos = []
    necesita_ocr = []
    no_soportados = []
    marcas_texto = set()

    reutilizar = previo is not None and archivos_reprocesar is not None
    registros_previos = {}
    parciales_previos = {}
    if reutilizar:
        registros_previos = {registro.get("archivo"): registro for registro in previo.get("resultados_archivos", [])}
        for producto in previo.get("productos_parciales", []):
            parciales_previos.setdefault(producto.get("archivo_fuente"), []).append(producto)
        marcas_texto.update(previo.get("marcas_detectadas_texto") or [])

    debug_dir = carpeta / "_debug_ia"
    if args.debug:
        debug_dir.mkdir(exist_ok=True)

    for indice, path in enumerate(sorted(archivos, key=prioridad), 1):
        verificar_detencion()
        relativo = str(path.relative_to(carpeta))
        if reutilizar and relativo not in archivos_reprocesar and relativo in registros_previos:
            registro = dict(registros_previos[relativo])
            registro["reutilizado"] = True
            parciales.extend(parciales_previos.get(relativo, []))
            if relativo in set(previo.get("necesita_ocr") or []):
                necesita_ocr.append(relativo)
            if relativo in set(previo.get("no_soportados") or []):
                no_soportados.append(relativo)
            resultados_archivos.append(registro)
            continue

        detalle = {} if args.vision and path.suffix.lower() == ".pdf" else None
        texto, estado_lectura = extraer_archivo(path, args, detalle)
        if detalle and args.vision_solo_escaneadas:
            detalle["paginas_vision"] = {i: m for i, m in (detalle.get("paginas_vision") or {}).items()
                                         if m == "sin_texto"}
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
        paginas_vision = (detalle or {}).get("paginas_vision") or {}
        if paginas_vision:
            registro["paginas_vision"] = {str(index + 1): motivo for index, motivo in sorted(paginas_vision.items())}

        if "ocr_pendiente" in estado_lectura:
            necesita_ocr.append(relativo)
        if estado_lectura in ("necesita_ocr", "ocr_dependencia_no_disponible") or estado_lectura.startswith("error_ocr"):
            if relativo not in necesita_ocr:
                necesita_ocr.append(relativo)
        elif estado_lectura in ("no_soportado", "xls_legacy_no_soportado"):
            no_soportados.append(relativo)
        elif estado_lectura in ("texto_vision", "imagen") and not args.solo_texto:
            productos, estado_ia, cruda, prompt, observaciones, consumo = extraer_parciales_mixto(
                path, relativo, detalle or {}, proveedor, rut, total_oferta, args
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
        "vision": bool(args.vision),
        "modo": "por_archivo",
        "segundos": round(time.perf_counter() - inicio_oferta, 1),
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


def diagnosticar_vision(licitaciones, args, ruta_csv):
    """Clasifica las paginas como lo haria --vision, sin llamar al modelo ni tocar
    checkpoints. Sirve para calibrar --vision-min-filas con documentos reales."""
    filas = []
    for licitacion in licitaciones:
        oferentes = sorted(
            carpeta for carpeta in licitacion.iterdir()
            if carpeta.is_dir() and not carpeta.name.startswith("_")
        )
        if args.proveedor:
            oferentes = [carpeta for carpeta in oferentes if carpeta.name == args.proveedor]
        if args.limite_proveedores:
            oferentes = oferentes[:args.limite_proveedores]
        for carpeta in oferentes:
            for path in sorted(listar_archivos_oferta(carpeta), key=prioridad):
                extension = path.suffix.lower()
                if extension not in EXTENSIONES_IMAGEN and extension != ".pdf":
                    continue
                fila = {
                    "licitacion": licitacion.name,
                    "proveedor": carpeta.name,
                    "archivo": str(path.relative_to(carpeta)),
                    "tipo_documental": tipo_documental(path.name),
                    "paginas_leidas": 1,
                    "paginas_vision": "",
                    "motivos": "",
                }
                if extension in EXTENSIONES_IMAGEN:
                    fila.update(paginas_vision="1", motivos="imagen")
                else:
                    detalle = {}
                    extraer_archivo(path, args, detalle)
                    paginas = detalle.get("paginas_vision", {})
                    fila["paginas_leidas"] = len(detalle.get("texto_paginas", {}))
                    fila["paginas_vision"] = ", ".join(str(indice + 1) for indice in sorted(paginas))
                    fila["motivos"] = ", ".join(paginas[indice] for indice in sorted(paginas))
                filas.append(fila)
                print(f"  {licitacion.name} | {carpeta.name[:30]:30} | {fila['archivo'][:40]:40} | "
                      f"vision: {fila['paginas_vision'] or '-'} {('(' + fila['motivos'] + ')') if fila['motivos'] else ''}")

    ruta_csv.parent.mkdir(parents=True, exist_ok=True)
    campos = ["licitacion", "proveedor", "archivo", "tipo_documental", "paginas_leidas", "paginas_vision", "motivos"]
    with ruta_csv.open("w", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=campos)
        escritor.writeheader()
        escritor.writerows(filas)

    con_vision = [fila for fila in filas if fila["paginas_vision"]]
    motivos = {}
    for fila in con_vision:
        for motivo in filter(None, fila["motivos"].split(", ")):
            motivos[motivo] = motivos.get(motivo, 0) + 1
    paginas_imagen = sum(len(fila["paginas_vision"].split(", ")) for fila in con_vision)
    por_llamada = max(1, args.vision_paginas_por_llamada)
    llamadas = sum(
        -(-len(fila["paginas_vision"].split(", ")) // por_llamada) for fila in con_vision
    )
    print("\n" + "=" * 74)
    print(f"DIAGNOSTICO VISION (--vision-min-filas {args.vision_min_filas})")
    print(f"Archivos PDF/imagen revisados: {len(filas)}")
    print(f"Archivos que usarian vision:   {len(con_vision)}")
    print(f"Paginas como imagen:           {paginas_imagen}  (~{llamadas} llamadas con imagen)")
    for motivo, cantidad in sorted(motivos.items(), key=lambda item: -item[1]):
        print(f"  {motivo:18} {cantidad} pagina(s)")
    print(f"Detalle: {ruta_csv}")
    print("=" * 74)


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
    with LOG_LOCK, ruta.open("a", encoding="utf-8") as archivo:
        archivo.write(json.dumps(registro, ensure_ascii=False) + "\n")


def resultado_requiere_reproceso(resultado, usar_ocr=False, modelo=None, backend=None, usar_vision=False,
                                 modo=None, solo_escaneadas=False):
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
        # Al activar --vision se reprocesan los proveedores leidos solo con texto.
        or (usar_vision and (bool(resultado.get("necesita_ocr")) if solo_escaneadas
                             else not resultado.get("vision")))
        or bool(modo and (resultado.get("modo") or "por_archivo") != modo)
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
        llamadas_archivo = int(consumo["llamadas"]) if consumo.get("llamadas") is not None else 1
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
    parser.add_argument("--paralelo", type=int, default=1,
                        help="Proveedores procesados a la vez. Debe coincidir con "
                             "'Max Concurrent Predictions' del modelo en LM Studio (ej. 4).")
    parser.add_argument("--vision", action="store_true",
                        help="Envia como imagen las paginas PDF con tablas sin bordes/colapsadas "
                             "o escaneadas, y los JPG/PNG adjuntos. Requiere modelo con vision.")
    parser.add_argument("--vision-solo-escaneadas", action="store_true",
                        help="Con --vision: solo las paginas sin texto van como imagen (con --ocr, solo las que el "
                             "OCR no pudo leer), y solo se reprocesan proveedores con OCR pendiente.")
    parser.add_argument("--dpi-vision", type=int, default=150,
                        help="Resolucion de las paginas enviadas como imagen (150 suele bastar).")
    parser.add_argument("--vision-paginas-por-llamada", type=int, default=3,
                        help="Paginas-imagen por llamada al modelo en modo --vision.")
    parser.add_argument("--rampa", type=float, default=5,
                        help="Segundos entre el inicio de los primeros proveedores en paralelo, para no "
                             "saturar LM Studio/LM Link al arrancar (0 = todos a la vez).")
    parser.add_argument("--sin-calentamiento", action="store_true",
                        help="No enviar la peticion de prueba inicial al modelo.")
    parser.add_argument("--por-proveedor", action="store_true",
                        help="Una llamada por proveedor con las paginas relevantes de todos sus anexos "
                             "(mas rapido). Sin esto: una llamada por fragmento de cada archivo + consolidacion.")
    parser.add_argument("--max-chars-proveedor", type=int, default=24000,
                        help="Texto maximo por llamada en --por-proveedor; si se excede se usan mas llamadas.")
    parser.add_argument("--max-imagenes-proveedor", type=int, default=4,
                        help="Paginas-imagen maximas por llamada en --por-proveedor (con --vision).")
    parser.add_argument("--vision-min-filas", type=int, default=2,
                        help="Filas con precio necesarias para tratar una pagina sin tabla "
                             "detectada como tabla sin bordes (mas alto = menos vision).")
    parser.add_argument("--diagnostico-vision", action="store_true",
                        help="Solo clasifica paginas (sin modelo ni checkpoints) y genera "
                             "diagnostico_vision.csv para calibrar la regla de vision.")
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

    if args.diagnostico_vision:
        diagnosticar_vision(licitaciones, args, ruta_excel.with_name("diagnostico_vision.csv"))
        return

    print("=" * 74)
    print(f"EXTRACCION IA v6 | backend={args.backend} | modelo={args.modelo} | licitaciones={len(licitaciones)}")
    print(f"7-Zip: {seven_zip or 'NO encontrado'} | solo_texto={args.solo_texto} "
          f"| paralelo={args.paralelo} | vision={args.vision}")
    print("Logica: fragmentos + vision selectiva + consolidacion y validacion auditable")
    print("=" * 74)
    escribir_log(
        ruta_log,
        "inicio_corrida",
        directorio=str(raiz),
        backend=args.backend,
        modelo=args.modelo,
        licitaciones=len(licitaciones),
        paralelo=args.paralelo,
        vision=args.vision,
        modo=modo_extraccion(args),
    )

    detener = False
    interrumpido = False
    if args.paralelo < 1:
        sys.exit("--paralelo debe ser mayor o igual a 1")
    DETENER.clear()

    # --- Fase 1: revisar checkpoints y armar la cola de proveedores pendientes ---
    # IMPORTANTE: los resultados "por revisar" se mantienen en la lista del checkpoint
    # hasta que su reproceso termine bien. Si el reproceso falla o se interrumpe, el
    # resultado anterior sigue intacto en extraccion_ia.json.
    estado_licitaciones = {}
    tareas = []
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
        previos_por_carpeta = {}
        if salida.exists() and not args.rehacer:
            try:
                resultados = json.loads(salida.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                resultados = []
            carpeta_por_rut = {
                str(cargar_json(carpeta / "oferta.json").get("rut") or "").strip(): carpeta
                for carpeta in oferentes
            }
            # Proveedores ya procesados que pueden necesitar trabajo (fallos, OCR
            # pendiente, archivos nuevos o revision para vision). No se rehacen
            # completos: el hilo de trabajo decide que archivos repetir.
            for resultado in resultados:
                carpeta = carpeta_por_rut.get(str(resultado.get("rut") or "").strip())
                if carpeta is not None and revision_necesaria(resultado, carpeta, args):
                    previos_por_carpeta[carpeta] = resultado
                else:
                    registrar_resultado(
                        licitacion.name,
                        metadata_licitaciones.get(licitacion.name, {}),
                        resultado,
                        todos,
                        consumos,
                        resumenes,
                        archivos_ocr,
                    )
            ruts_guardados = {
                str(resultado.get("rut") or "").strip() for resultado in resultados
            }
            oferentes = [
                carpeta for carpeta in oferentes
                if carpeta in previos_por_carpeta
                or str(cargar_json(carpeta / "oferta.json").get("rut") or "").strip()
                not in ruts_guardados
            ]
            if not oferentes:
                print(f"{licitacion.name}: ya procesada ({len(resultados)} proveedores)")
                continue
            nuevos = len(oferentes) - len(previos_por_carpeta)
            print(f"{licitacion.name}: reanudando; {nuevos} proveedores sin procesar y "
                  f"{len(previos_por_carpeta)} por revisar (solo se repiten los archivos necesarios)")

        if args.limite_proveedores is not None:
            if args.limite_proveedores < 1:
                sys.exit("--limite-proveedores debe ser mayor que cero")
            oferentes = oferentes[:args.limite_proveedores]
        # Los "por revisar" que quedaron fuera de la cola (por --limite-proveedores)
        # se registran tal cual para que igual aparezcan en el Excel.
        for carpeta, previo in previos_por_carpeta.items():
            if carpeta not in oferentes:
                registrar_resultado(
                    licitacion.name,
                    metadata_licitaciones.get(licitacion.name, {}),
                    previo,
                    todos,
                    consumos,
                    resumenes,
                    archivos_ocr,
                )

        print(f"{licitacion.name}: {len(oferentes)} proveedores en cola")
        estado_licitaciones[licitacion.name] = {"salida": salida, "resultados": resultados}
        tareas.extend((licitacion, carpeta, previos_por_carpeta.get(carpeta)) for carpeta in oferentes)

    # --- Fase 2: procesar la cola, hasta --paralelo proveedores a la vez ---
    # Cada hilo procesa un proveedor completo (lectura + llamadas al modelo).
    # Los resultados, checkpoints y el Excel se manejan solo desde este hilo principal,
    # asi que el formato de extraccion_ia.json y la reanudacion no cambian.
    if args.pausa_licitacion > 0 and args.paralelo > 1:
        print("Aviso: --pausa-licitacion se ignora cuando --paralelo es mayor que 1.")
    ultima_licitacion = {"nombre": None}

    arranque = {"siguiente": 0}
    lock_arranque = threading.Lock()

    def trabajar(licitacion, carpeta, previo):
        verificar_detencion()
        with lock_arranque:
            orden = arranque["siguiente"]
            arranque["siguiente"] += 1
        if orden < args.paralelo and args.rampa > 0:
            DETENER.wait(orden * args.rampa)  # 0 s, rampa, 2*rampa... (interrumpible con Ctrl+C)
            verificar_detencion()
        if args.paralelo == 1 and args.pausa_licitacion > 0:
            if ultima_licitacion["nombre"] not in (None, licitacion.name):
                print(f"Pausa de {args.pausa_licitacion} segundos para enfriar el equipo...")
                DETENER.wait(args.pausa_licitacion)
                verificar_detencion()
            ultima_licitacion["nombre"] = licitacion.name
        info_carpeta = cargar_json(carpeta / "oferta.json")
        escribir_log(
            ruta_log,
            "inicio_proveedor",
            licitacion=licitacion.name,
            proveedor=info_carpeta.get("proveedor", carpeta.name),
            rut=info_carpeta.get("rut", ""),
        )
        procesar = procesar_oferta_por_proveedor if args.por_proveedor else procesar_oferta
        if previo is None:
            return procesar(carpeta, args, seven_zip)
        modo, archivos = planificar_reproceso(carpeta, previo, args, seven_zip)
        if modo == "completo" or (args.por_proveedor and modo in ("parcial", "consolidar")):
            # en modo por proveedor cualquier cambio rehace el proveedor: son 1 o 2 llamadas
            resultado = procesar(carpeta, args, seven_zip)
        elif modo == "nada":
            resultado = dict(previo)
            resultado["productos"] = list(previo.get("productos", []))
            resultado["vision"] = bool(previo.get("vision") or args.vision)
        else:
            resultado = procesar_oferta(carpeta, args, seven_zip, previo=previo, archivos_reprocesar=set(archivos))
        resultado["reproceso"] = {"modo": modo, "archivos": archivos or {}}
        return resultado

    total_tareas = len(tareas)
    completados = 0
    terminados_ok = set()

    def manejar(futuro, licitacion, carpeta):
        """Procesa el resultado de un proveedor en el hilo principal."""
        nonlocal detener, completados
        if futuro.cancelled():
            return
        info_carpeta = cargar_json(carpeta / "oferta.json")
        try:
            resultado = futuro.result()
        except CorridaDetenida:
            return  # abandonado por Ctrl+C o por caida del backend; el previo se conserva
        except RuntimeError as exc:
            escribir_log(
                ruta_log,
                "error_proveedor",
                licitacion=licitacion.name,
                proveedor=info_carpeta.get("proveedor", carpeta.name),
                rut=info_carpeta.get("rut", ""),
                error=str(exc),
            )
            if not detener:
                print(f"\nERROR FATAL: {exc}")
                print(f"Se detiene la corrida porque {args.backend} dejo de responder.")
                print("Se cancelan los proveedores en cola; los que estaban en curso se abandonan.")
                detener = True
                DETENER.set()
                for pendiente in futuros:
                    pendiente.cancel()
            return
        except Exception as exc:
            print(f"\nERROR EN PROVEEDOR {licitacion.name}/{carpeta.name}: {exc}")
            escribir_log(
                ruta_log,
                "error_proveedor",
                licitacion=licitacion.name,
                proveedor=info_carpeta.get("proveedor", carpeta.name),
                rut=info_carpeta.get("rut", ""),
                error=str(exc),
            )
            return

        completados += 1
        terminados_ok.add(carpeta)
        estado_licitacion = estado_licitaciones[licitacion.name]
        resultados = estado_licitacion["resultados"]
        reemplazar_resultado(resultados, resultado)
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
            segundos=resultado.get("segundos"),
            llamadas=resultado.get("llamadas_modelo"),
            modo=resultado.get("modo"),
        )
        print(
            f"  [{completados}/{total_tareas}] {resultado.get('segundos', '?')}s | {licitacion.name} | "
            f"{resultado['proveedor'][:34]:34} "
            f"parciales={len(resultado['productos_parciales']):2} "
            f"finales={len(resultado['productos']):2} "
            f"ocr={len(resultado['necesita_ocr']):2} "
            f"consol={resultado['estado_consolidacion']}"
        )
        reproceso = resultado.get("reproceso") or {}
        if reproceso.get("modo") == "nada":
            print("      sin cambios: no habia archivos que rehacer")
        elif reproceso.get("modo") == "consolidar":
            print("      solo se repitio la consolidacion")
        elif reproceso.get("modo") == "parcial":
            for archivo_rehecho, motivo in sorted(reproceso["archivos"].items()):
                print(f"      rehecho [{motivo}] {archivo_rehecho[:60]}")
        elif reproceso.get("modo") == "completo":
            print("      rehecho completo (cambio de modelo o backend)")
        # Diagnostico visible por archivo: antes estos estados quedaban solo en JSON.
        for archivo in resultado["resultados_archivos"]:
            estado_ia = archivo.get("estado_ia", "")
            encontrados = archivo.get("productos_encontrados", 0)
            if archivo.get("paginas_vision") and not archivo.get("reutilizado"):
                print(
                    f"      [vision] {archivo.get('archivo','')[:47]} | "
                    f"paginas={', '.join(archivo['paginas_vision'])}"
                )
            archivo_sin_problema = resultado.get("modo") == "por_proveedor" and estado_ia in (
                "no_ejecutada", "ok", "ok_reparado", "ok_extraido")
            if not archivo_sin_problema and (
                    estado_ia not in ("no_ejecutada", "solo_texto", "ok", "ok_reparado") or encontrados == 0):
                print(
                    f"      [{archivo.get('tipo_documental','otro')}] "
                    f"{archivo.get('archivo','')[:47]} | lectura={archivo.get('estado_lectura')} "
                    f"| ia={estado_ia} | productos={encontrados}"
                )
                if archivo.get("observaciones"):
                    print(f"        observacion: {str(archivo['observaciones'])[:180]}")
        for producto in resultado["productos"]:
            print(f"    - {producto['producto'][:55]} | unitario={producto['precio_unitario']}")

        # checkpoint por proveedor para no perder el avance si el backend cae
        escribir_json_atomico(estado_licitacion["salida"], resultados)

    def esperar(futuros_pendientes):
        # wait() con timeout (en vez de as_completed sin timeout) para que Ctrl+C
        # se atienda tambien en Windows, donde una espera sin timeout lo bloquea.
        while futuros_pendientes:
            hechos, futuros_pendientes = wait(futuros_pendientes, timeout=1.0, return_when=FIRST_COMPLETED)
            for futuro in hechos:
                licitacion, carpeta = futuros[futuro]
                manejar(futuro, licitacion, carpeta)

    print(f"\nProveedores por procesar: {total_tareas} | en paralelo: {args.paralelo}")
    if tareas and not args.solo_texto and not args.sin_calentamiento:
        print("Calentando el modelo (activa LM Link y confirma que responde)...")
        segundos_calentamiento = calentar_modelo(args)
        if segundos_calentamiento is None:
            sys.exit(f"El modelo no respondio en {args.backend}. Revisa que este cargado (y LM Link conectado).")
        print(f"Modelo listo ({segundos_calentamiento} s).")
    pool = ThreadPoolExecutor(max_workers=args.paralelo)
    futuros = {pool.submit(trabajar, licitacion, carpeta, previo): (licitacion, carpeta)
               for licitacion, carpeta, previo in tareas}
    try:
        esperar(set(futuros))
    except KeyboardInterrupt:
        interrumpido = True
        detener = True
        DETENER.set()
        pool.shutdown(wait=False, cancel_futures=True)
        print("\nCtrl+C recibido: se cancelan los proveedores en cola y se abandona lo que")
        print("estaba en curso (se espera como maximo la respuesta de las llamadas activas).")
        print("Presiona Ctrl+C otra vez para salir de inmediato; los checkpoints guardados estan completos.")
        try:
            esperar({futuro for futuro in futuros if not futuro.done()})
        except KeyboardInterrupt:
            print("Salida forzada.")
            os._exit(130)
    finally:
        pool.shutdown(wait=True, cancel_futures=True)

    # Resultados "por revisar" cuyo reproceso no termino bien: siguen en el
    # checkpoint (nunca se quitaron) y se registran tal cual para el Excel.
    for licitacion, carpeta, previo in tareas:
        if previo is not None and carpeta not in terminados_ok:
            registrar_resultado(
                licitacion.name,
                metadata_licitaciones.get(licitacion.name, {}),
                previo,
                todos,
                consumos,
                resumenes,
                archivos_ocr,
            )
    if interrumpido:
        print("Corrida interrumpida por el usuario. Repite el comando para retomar.")

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
        interrumpida=interrumpido,
    )
    if interrumpido:
        sys.exit(130)


if __name__ == "__main__":
    main()

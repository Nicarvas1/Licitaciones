#!/usr/bin/env python3
"""Verifica que la maquina este lista para una corrida larga y desatendida del pipeline.

Revisa todo lo que en una corrida de fin de semana puede fallar en silencio:
  - librerias de Python y el navegador de Playwright
  - acceso a mercadopublico.cl
  - Tesseract (OCR) y su idioma espanol; antiword (.doc) y 7-Zip (.rar/.7z)
  - LM Studio: que responda, que el modelo este cargado, que devuelva JSON sin "pensar" y cuantas
    peticiones atiende realmente en paralelo
  - espacio en disco, rutas largas de Windows y que los datos no esten bajo OneDrive

  python verificar_entorno.py --modelo qwen/qwen3.6-35b-a3b --salida D:\\lotes
  python verificar_entorno.py --preparar-tessdata      # deja Tesseract con espanol en .\\tessdata (sin admin)

Sale con codigo 1 si hay algun ERROR. Los AVISO no detienen la corrida pero conviene revisarlos.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import importlib
import json
import os
import shutil
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
TESSDATA_LOCAL = RAIZ / "tessdata"
URL_SPA = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/spa.traineddata"
RUTAS_TESSERACT = (r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe")

resultados = []


def informe(nivel, area, mensaje):
    resultados.append((nivel, area, mensaje))
    marca = {"OK": "[ OK  ]", "AVISO": "[AVISO]", "ERROR": "[ERROR]"}[nivel]
    print(f"{marca} {area:<14} {mensaje}")


# ---------------------------------------------------------------------------
def revisar_librerias():
    faltan = []
    for modulo in ("bs4", "playwright", "pymupdf", "pdfplumber", "pytesseract", "PIL", "docx", "openpyxl", "requests"):
        try:
            importlib.import_module(modulo)
        except Exception:
            faltan.append(modulo)
    if faltan:
        informe("ERROR", "Python", f"faltan librerias: {', '.join(faltan)} -> pip install -r requirements.txt")
    else:
        informe("OK", "Python", f"{sys.version.split()[0]} con todas las librerias")


def revisar_playwright():
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=True)
            navegador.close()
        informe("OK", "Playwright", "Chromium se abre correctamente")
    except Exception as exc:
        informe("ERROR", "Playwright", f"no abre Chromium ({str(exc)[:90]}) -> python -m playwright install chromium")


def revisar_portal():
    import requests
    try:
        r = requests.get("https://www.mercadopublico.cl/Home", timeout=20)
        if r.status_code < 400:
            informe("OK", "Portal", "mercadopublico.cl responde")
        else:
            informe("AVISO", "Portal", f"mercadopublico.cl respondio {r.status_code}")
    except Exception as exc:
        informe("ERROR", "Portal", f"sin acceso a mercadopublico.cl ({str(exc)[:80]})")


def preparar_tessdata():
    """Copia los idiomas instalados a .\\tessdata y agrega spa.traineddata (no requiere permisos de admin)."""
    import requests
    tesseract = shutil.which("tesseract") or next((r for r in RUTAS_TESSERACT if os.path.isfile(r)), None)
    if not tesseract:
        sys.exit("No encontre Tesseract. Instalalo primero (https://github.com/UB-Mannheim/tesseract/wiki).")
    TESSDATA_LOCAL.mkdir(exist_ok=True)
    origen = Path(tesseract).parent / "tessdata"
    for nombre in ("eng.traineddata", "osd.traineddata"):
        if (origen / nombre).is_file() and not (TESSDATA_LOCAL / nombre).is_file():
            shutil.copy2(origen / nombre, TESSDATA_LOCAL / nombre)
            print(f"  copiado {nombre}")
    destino = TESSDATA_LOCAL / "spa.traineddata"
    if not destino.is_file():
        print(f"  descargando {URL_SPA} ...")
        r = requests.get(URL_SPA, timeout=120)
        r.raise_for_status()
        destino.write_bytes(r.content)
        print(f"  guardado {destino} ({len(r.content) // 1024} KB)")
    print(f"Listo: {TESSDATA_LOCAL} (el extractor lo usa automaticamente)")


def revisar_ocr():
    try:
        sys.path.insert(0, str(RAIZ))
        spec = importlib.util.spec_from_file_location("extractor_verificacion", RAIZ / "3_extraer_ia.py")
        extractor = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(extractor)
    except Exception as exc:
        informe("ERROR", "OCR", f"no se pudo cargar 3_extraer_ia.py ({str(exc)[:80]})")
        return
    ok, mensaje = extractor.verificar_ocr(None, "spa+eng")
    if not ok:
        informe("ERROR", "OCR", mensaje)
        return
    nivel = "AVISO" if "AVISO" in mensaje else "OK"
    if nivel == "AVISO":
        mensaje += " -> python verificar_entorno.py --preparar-tessdata"
    informe(nivel, "OCR", mensaje)
    # prueba real: un PDF escaneado sintetico con un monto
    try:
        import pymupdf
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "escaneado.pdf"
            imagen = pymupdf.open()
            pagina = imagen.new_page()
            pagina.insert_text((60, 120), "NOTEBOOK HP PROBOOK 440 TOTAL $ 1.234.567", fontsize=20)
            pix = pagina.get_pixmap(dpi=150)
            escaneado = pymupdf.open()
            hoja = escaneado.new_page(width=pagina.rect.width, height=pagina.rect.height)
            hoja.insert_image(hoja.rect, pixmap=pix)
            escaneado.save(pdf)
            texto, estado = extractor.extraer_pdf_ocr(pdf, 2, None, "spa+eng", 200)
        if "1.234.567" in texto or "1234567" in texto.replace(".", ""):
            informe("OK", "OCR prueba", "lee un PDF escaneado de prueba (monto reconocido)")
        else:
            informe("ERROR", "OCR prueba", f"el OCR no reconocio el texto de prueba (estado {estado})")
    except Exception as exc:
        informe("ERROR", "OCR prueba", f"fallo la prueba de OCR ({str(exc)[:90]})")


def revisar_herramientas():
    antiword = shutil.which("antiword") or next(
        (r for r in (r"C:\Program Files\Git\mingw64\bin\antiword.exe", r"C:\Program Files\Git\usr\bin\antiword.exe")
         if os.path.isfile(r)), None)
    if antiword:
        informe("OK", "antiword", antiword)
    else:
        informe("AVISO", "antiword", "no encontrado: los .doc antiguos quedaran sin leer (instala Git for Windows)")
    siete = shutil.which("7z") or next((r for r in (r"C:\Program Files\7-Zip\7z.exe", r"C:\Program Files (x86)\7-Zip\7z.exe")
                                        if os.path.isfile(r)), None)
    if siete:
        informe("OK", "7-Zip", siete)
    else:
        informe("AVISO", "7-Zip", "no encontrado: los .rar/.7z quedaran sin abrir (los .zip si se abren)")


def _peticion(url, modelo, contenido, max_tokens, timeout):
    import requests
    payload = {
        "model": modelo,
        "messages": [{"role": "system", "content": "Devuelve exclusivamente un objeto JSON valido."},
                     {"role": "user", "content": contenido}],
        "temperature": 0.0, "max_tokens": max_tokens, "reasoning_effort": "none", "stream": False,
        "chat_template_kwargs": {"enable_thinking": False},
    }
    inicio = time.perf_counter()
    r = requests.post(url, json=payload, timeout=timeout)
    r.raise_for_status()
    datos = r.json()
    return time.perf_counter() - inicio, datos


def revisar_lmstudio(url_base, modelo, paralelo):
    import requests
    try:
        modelos = [m.get("id") for m in requests.get(f"{url_base}/v1/models", timeout=15).json().get("data", [])]
    except Exception as exc:
        informe("ERROR", "LM Studio", f"no responde en {url_base} ({str(exc)[:80]})")
        return
    if modelo not in modelos:
        informe("ERROR", "LM Studio", f"'{modelo}' no esta disponible. Modelos: {', '.join(modelos)}")
        return
    url = f"{url_base}/v1/chat/completions"
    try:
        segundos, datos = _peticion(url, modelo, 'Responde {"ok": true, "suma": 2+3} con la suma resuelta.', 60, 300)
    except Exception as exc:
        informe("ERROR", "LM Studio", f"el modelo no contesto ({str(exc)[:90]})")
        return
    mensaje = (datos.get("choices") or [{}])[0].get("message", {})
    contenido = mensaje.get("content") or ""
    razonamiento = mensaje.get("reasoning_content") or mensaje.get("reasoning") or ""
    try:
        json.loads(contenido.strip().strip("`").removeprefix("json").strip())
        informe("OK", "LM Studio", f"{modelo} responde JSON valido ({segundos:.1f} s, incluye carga del modelo)")
    except Exception:
        informe("ERROR", "LM Studio", f"la respuesta no es JSON: {contenido[:80]!r}")
    if "<think>" in contenido or len(razonamiento) > 20:
        informe("AVISO", "LM Studio", "el modelo esta razonando antes de responder: la corrida sera mucho mas lenta")
    # paralelismo real: n peticiones a la vez vs una sola
    tarea = 'Devuelve {"n": [1,2,3,...,60]} con los numeros del 1 al 60.'
    try:
        uno, _ = _peticion(url, modelo, tarea, 300, 300)
        inicio = time.perf_counter()
        with concurrent.futures.ThreadPoolExecutor(paralelo) as pool:
            list(pool.map(lambda _: _peticion(url, modelo, tarea, 300, 600), range(paralelo)))
        varios = time.perf_counter() - inicio
        efectivas = paralelo * uno / varios if varios else 1
        nivel = "OK" if efectivas >= paralelo * 0.6 else "AVISO"
        texto = (f"1 peticion {uno:.1f} s; {paralelo} a la vez {varios:.1f} s -> ~{efectivas:.1f} en paralelo")
        if nivel == "AVISO":
            texto += f" (en LM Studio sube 'Max Concurrent Predictions' a {paralelo})"
        informe(nivel, "Paralelismo", texto)
    except Exception as exc:
        informe("AVISO", "Paralelismo", f"no se pudo medir ({str(exc)[:80]})")


def revisar_disco(salida):
    destino = Path(salida).resolve()
    destino.mkdir(parents=True, exist_ok=True)
    try:
        prueba = destino / ".prueba_escritura"
        prueba.write_text("ok", encoding="utf-8")
        prueba.unlink()
    except Exception as exc:
        informe("ERROR", "Disco", f"no se puede escribir en {destino} ({exc})")
        return
    libre = shutil.disk_usage(destino).free / 1024 ** 3
    informe("OK" if libre >= 30 else "AVISO", "Disco", f"{libre:.0f} GB libres en {destino.anchor} (un ano usa ~10 GB)")
    if "onedrive" in str(destino).lower():
        informe("AVISO", "Disco", "la salida esta bajo OneDrive: la sincronizacion de miles de archivos frena todo")
    if os.name == "nt":
        try:
            import winreg
            clave = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem")
            largas = winreg.QueryValueEx(clave, "LongPathsEnabled")[0]
            informe("OK" if largas else "AVISO", "Rutas largas",
                    "habilitadas" if largas else "deshabilitadas (el pipeline las maneja, pero conviene habilitarlas)")
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(description="Verifica el entorno antes de una corrida larga")
    parser.add_argument("--modelo", default="qwen/qwen3.6-35b-a3b")
    parser.add_argument("--lmstudio", default=os.environ.get("LMSTUDIO_BASE", "http://127.0.0.1:1234"))
    parser.add_argument("--paralelo", type=int, default=4)
    parser.add_argument("--salida", default="lotes", help="carpeta donde quedaran los datos")
    parser.add_argument("--sin-portal", action="store_true", help="no revisa Playwright ni el portal")
    parser.add_argument("--sin-modelo", action="store_true", help="no revisa LM Studio")
    parser.add_argument("--preparar-tessdata", action="store_true",
                        help="copia los idiomas de Tesseract a .\\tessdata y descarga el espanol")
    args = parser.parse_args()
    if args.preparar_tessdata:
        preparar_tessdata()
        return

    print("=" * 74)
    print("Verificacion del entorno")
    print("=" * 74)
    revisar_librerias()
    if not args.sin_portal:
        revisar_playwright()
        revisar_portal()
    revisar_ocr()
    revisar_herramientas()
    if not args.sin_modelo:
        revisar_lmstudio(args.lmstudio.rstrip("/"), args.modelo, args.paralelo)
    revisar_disco(args.salida)
    errores = [r for r in resultados if r[0] == "ERROR"]
    avisos = [r for r in resultados if r[0] == "AVISO"]
    print("=" * 74)
    print(f"Errores: {len(errores)} | Avisos: {len(avisos)}")
    if errores:
        print("Corrige los errores antes de dejar la corrida sola.")
        sys.exit(1)
    print("Listo para correr.")


if __name__ == "__main__":
    main()

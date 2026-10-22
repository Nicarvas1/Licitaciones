#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 scraper_ofertas.py   (v5 - economico primero y fallback tecnico)
=============================================================================
Descarga primero los anexos económicos de TODAS las ofertas de las
licitaciones listadas en para_scrapear.csv (generado por 1_filtrar_licitaciones.py).

FLUJO por licitación:
  1. Ficha (DetailsAcquisition.aspx?idlicitacion=CODIGO)  <- usa el CÓDIGO directo
  2. Extrae link del "Cuadro de Ofertas" (imgCuadroOferta)
  3. Abre el cuadro (grilla vive en un IFRAME)
  4. Por cada oferta (todas las páginas) -> descarga anexos
  5. Guarda resumen_ofertas.csv (proveedor + total + estado)

REANUDACIÓN:
  Actualiza la columna 'scraping_estado' del propio para_scrapear.csv
  (pendiente -> ok / error). Al reejecutar, salta las que ya están 'ok'.

REQUISITOS:  pip install -r requirements.txt
             playwright install chromium

USO:
    # Primera pasada: solo económicos.
  python scraper_ofertas.py --csv para_scrapear.csv --limite 3 --ver

    # Producción económica (todas, reanudable, headless):
  python scraper_ofertas.py --csv para_scrapear.csv

    # Fallback técnico para proveedores que aún no tienen precio en los resultados IA:
    python scraper_ofertas.py --csv para_scrapear.csv --solo-tecnicos-fallback

    # Una licitación suelta por código o qs (modo test):
  python scraper_ofertas.py --codigo 3572-22-LE25 --ver
=============================================================================
"""

import argparse
import csv
import json
import math
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

try:
    from bs4 import BeautifulSoup
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError as e:
    sys.exit(f"Falta librería: {e}\nInstala:\n  pip install playwright beautifulsoup4 pandas\n  playwright install chromium")

BASE = "https://www.mercadopublico.cl"
FICHA_POR_CODIGO = BASE + "/Procurement/Modules/RFB/DetailsAcquisition.aspx?idlicitacion={}"
FICHA_POR_QS = BASE + "/Procurement/Modules/RFB/DetailsAcquisition.aspx?qs={}"
SALIDA = Path("ofertas")
DIAG = False

ANEXOS = {
    "administrativo": "_GvImgbAdministrativeAttachment",
    "tecnico": "_GvImgbTechnicalAttachment",
    "economico": "_GvImgbEconomicAttachment",
}


# ---------------------------------------------------------------------------
# Parseo
# ---------------------------------------------------------------------------
def links_de_ficha(html):
    soup = BeautifulSoup(html, "html.parser")
    out = {"cuadro_ofertas": None, "acta_adjudicacion": None}
    c = soup.find("input", id="imgCuadroOferta")
    if c and c.get("href"):
        out["cuadro_ofertas"] = urljoin(BASE, c.get("href"))
    a = soup.find("input", id="imgAdjudicacion")
    if a and a.get("href"):
        out["acta_adjudicacion"] = urljoin(BASE, a.get("href"))
    return out


def parsear_ofertas(html):
    soup = BeautifulSoup(html, "html.parser")
    filas = {}
    for inp in soup.find_all("input"):
        iid = inp.get("id", "") or ""
        m = re.match(r"(grdSupplies_ctl\d+)_", iid)
        if not m:
            continue
        ctl = m.group(1)
        filas.setdefault(ctl, {"rut": None, "proveedor": None, "total": None,
                               "estado": None, "anexos": {}})
        oc = inp.get("onclick", "") or ""
        mr = re.search(r"(?:ver_declaracion|verFicha)\('([\d.\-kK]+)'", oc)
        if mr and not filas[ctl]["rut"]:
            filas[ctl]["rut"] = mr.group(1)
        for tipo, suf in ANEXOS.items():
            if iid.endswith(suf):
                mu = re.search(r"openPopUp\('([^']+)'", oc)
                if mu:
                    filas[ctl]["anexos"][tipo] = urljoin(BASE, mu.group(1))
    for ctl, d in filas.items():
        prov = soup.find("a", id=f"{ctl}__GvLblProvider")
        if prov: d["proveedor"] = prov.get_text(strip=True)
        tot = soup.find("span", id=f"{ctl}_TotalOferta")
        if tot: d["total"] = tot.get_text(strip=True)
        est = soup.find("span", id=f"{ctl}_EstadoOferta")
        if est: d["estado"] = est.get_text(strip=True)
        if not d["rut"]:
            ar = soup.find("a", id=f"{ctl}__GvLblRutProvider")
            if ar: d["rut"] = ar.get_text(strip=True)
    return [d for d in filas.values() if d["rut"] or d["proveedor"]]


def total_paginas(html):
    soup = BeautifulSoup(html, "html.parser")
    def _v(i):
        el = soup.find("input", id=i)
        try:
            return int(el.get("value")) if el and el.get("value") else None
        except (ValueError, TypeError):
            return None
    count = _v("WucPagerGrid_hidCountRows")
    per = _v("WucPagerGrid_hidMaxRowCount")
    if count and per and per > 0:
        return max(1, math.ceil(count / per))
    pager = soup.find(id="WucPagerGrid__TblPages")
    paginas = {1}
    if pager:
        for div in pager.find_all("div"):
            mm = re.search(r"fnMovePage\((\d+)", div.get("onclick", "") or "")
            if mm: paginas.add(int(mm.group(1)))
            t = div.get_text(strip=True)
            if t.isdigit(): paginas.add(int(t))
    return max(paginas)


def localizar_frame(page, reintentos=3):
    for _ in range(reintentos):
        try:
            if "grdSupplies" in page.content():
                return page, page.content()
        except Exception:
            pass
        for fr in page.frames:
            try:
                h = fr.content()
                if "grdSupplies" in h:
                    return fr, h
            except Exception:
                continue
        time.sleep(1.5)
    return page, (page.content() if page else "")


SELECTOR_DESCARGA = "input[id^='DWNL_grdId_'][id$='_search']"
SELECTOR_DESCARGA_ALTERNO = "input[type='image'][src*='ver.gif']"
# El popup de anexos muestra 6 archivos por pagina (GridView de ASP.NET): el resto esta en "Page$2", "Page$3"...
ARCHIVOS_POR_PAGINA_POPUP = 6
PATRON_PAGINA_POPUP = re.compile(r"__doPostBack\('DWNL\$grdId','Page\$(\d+)'\)")


def botones_de_descarga(popup):
    return popup.query_selector_all(SELECTOR_DESCARGA) or popup.query_selector_all(SELECTOR_DESCARGA_ALTERNO)


def abrir_popup(context, url_popup, intentos=3):
    """Abre el popup de anexos; reintenta si no cargo o llego sin botones (carga intermitente del portal)."""
    for intento in range(1, intentos + 1):
        popup = context.new_page()
        try:
            popup.goto(url_popup, wait_until="networkidle", timeout=45000)
            if botones_de_descarga(popup):
                return popup
        except PWTimeout:
            pass
        popup.close()
        if intento < intentos:
            time.sleep(2.0 * intento)
    return None


def archivo_ya_bajado(carpeta, prefijo, indice):
    return any(Path(carpeta).glob(f"{prefijo}__{indice:02d}__*"))


def ir_a_pagina_popup(popup, pagina):
    """Navega el popup a la pagina N de su grilla (postback de ASP.NET). La 1 es la que carga por defecto."""
    if pagina <= 1:
        return True
    try:
        with popup.expect_navigation(wait_until="networkidle", timeout=45000):
            popup.evaluate(f"__doPostBack('DWNL$grdId','Page${pagina}')")
        return True
    except Exception as e:
        print(f"        ⚠️ pagina {pagina} del popup: {str(e)[:40]}")
        return False


PATRON_ADMIN_UTIL = re.compile(
    r"cotiz|ofert(?!ente)|econ[oó]m|precio|valor|ficha|t[eé]cnic|especific|propuesta|cat[aá]logo|"
    r"resumen|equipo|notebook|computador|monitor|impresora|producto|(?<![a-z])item|(?<![a-z])lote", re.I)
PATRON_ADMIN_SIN_OFERTA = re.compile(
    r"declara|certific|(?<![a-z])rut(?![a-z])|carnet|c[eé]dula|poder|escritura|integridad|pacto|boleta|"
    r"patente|tribut|f[-_ ]?30(?!\d)|antecedente|identificaci|seriedad|constituci|vigencia|experiencia|"
    r"orden(?:es)?[ _-]*de[ _-]*compra|(?<!\d)\d{3,7}-\d{1,4}-(?:se|ag|cm|co|ac|sb)\d{0,2}(?![a-z])|curr[ií]cul|"
    r"inhabilidad|sello|pol[ií]tica|balance|financier|compromiso|post[ _-]*venta|simult|conflicto|reciclaj|"
    r"iso[ _-]?\d|bancari|cuenta[ _-]corriente|mujer|pyme|chileproveedores|garant[ií]a|prevenci|delito|factura|"
    r"interactivo|sociedad|(?<![a-z])c\.?i\.?(?![a-z])|(?<![a-z])oc(?![a-z])", re.I)


def es_administrativo_sin_oferta(nombre):
    """Archivo del anexo administrativo que por su nombre claramente no es la oferta: declaraciones juradas,
    certificados, RUT, escrituras, programas de integridad u ordenes de compra de ventas pasadas (experiencia).
    Lo que nombre una cotizacion, oferta, ficha o equipo siempre se conserva."""
    nombre = str(nombre or "")
    return bool(PATRON_ADMIN_SIN_OFERTA.search(nombre)) and not PATRON_ADMIN_UTIL.search(nombre)


def nombre_en_popup(popup, posicion):
    """Texto de la fila del popup (incluye el nombre del archivo) para el boton de descarga en `posicion`."""
    try:
        botones = botones_de_descarga(popup)
        if posicion < len(botones):
            return botones[posicion].evaluate("b => { const f = b.closest('tr'); return f ? f.innerText : ''; }")
    except Exception:
        pass
    return ""


def descargar_popup(context, url_popup, carpeta, prefijo, max_paginas=40, omitir=None, omitidos=None):
    """Descarga TODOS los archivos de un anexo recorriendo las paginas del popup.

    El portal vuelve a dibujar el popup despues de cada descarga, asi que los botones se piden de nuevo
    antes de cada clic (si no, quedan "desprendidos del DOM" y la descarga falla). Si un archivo falla se
    reabre el popup en la misma pagina y se reintenta. Los archivos que ya existen (mismo prefijo e indice)
    no se vuelven a bajar. Devuelve cuantos archivos nuevos bajo."""
    bajados = 0
    popup = abrir_popup(context, url_popup)
    if popup is None:
        return 0
    indice, pagina, visitadas = 0, 1, set()
    while pagina not in visitadas and len(visitadas) < max_paginas:
        visitadas.add(pagina)
        for posicion in range(len(botones_de_descarga(popup))):
            if omitir is not None:
                fila = nombre_en_popup(popup, posicion)
                if fila and omitir(fila):
                    # Los omitidos no consumen numero: la numeracion queda continua y estable entre corridas.
                    if omitidos is not None:
                        omitidos.append(" ".join(fila.split())[:160])
                    continue
            indice += 1
            if archivo_ya_bajado(carpeta, prefijo, indice):
                continue
            ultimo_error = ""
            for intento in (1, 2, 3):
                try:
                    botones = botones_de_descarga(popup)
                    if posicion >= len(botones):
                        raise RuntimeError("el boton de descarga ya no esta en la pagina")
                    with popup.expect_download(timeout=45000) as dl:
                        botones[posicion].click()
                    d = dl.value
                    nom = re.sub(r"[^\w\s.\-()]", "_", d.suggested_filename or f"{prefijo}_{indice}")[:150]
                    dest = carpeta / f"{prefijo}__{indice:02d}__{nom}"
                    d.save_as(str(dest))
                    bajados += 1
                    print(f"        ✅ {dest.name[:58]}")
                    ultimo_error = ""
                    break
                except Exception as e:
                    ultimo_error = str(e)[:40]
                    time.sleep(1.5 * intento)
                    if intento >= 2:  # el popup quedo en mal estado: se reabre en la misma pagina
                        nuevo_popup = abrir_popup(context, url_popup)
                        if nuevo_popup is not None:
                            try:
                                popup.close()
                            except Exception:
                                pass
                            popup = nuevo_popup
                            ir_a_pagina_popup(popup, pagina)
            if ultimo_error:
                print(f"        ⚠️ archivo {indice}: {ultimo_error}")
            time.sleep(0.4)
        siguientes = sorted(int(n) for n in PATRON_PAGINA_POPUP.findall(popup.content()) if int(n) > pagina)
        if not siguientes:
            break
        if not ir_a_pagina_popup(popup, siguientes[0]):
            break
        pagina = siguientes[0]
    popup.close()
    return bajados


# ---------------------------------------------------------------------------
# Procesar UNA licitación (por código o qs)
# ---------------------------------------------------------------------------
def carpeta_tiene_archivos(carpeta):
    """True si la carpeta del proveedor ya tiene algun documento descargado (no cuenta .json/.csv/.html)."""
    return any(f.is_file() and f.suffix.lower() not in (".json", ".csv", ".html")
               for f in Path(carpeta).iterdir())


def procesar_licitacion(context, identificador, es_qs=False,
                        tipos_anexos=None, ruts_objetivo=None, ruts_excluir=None, admin_si_vacio=True):
    tipos_anexos = set(tipos_anexos or {"economico"})
    ruts_objetivo = {str(rut).strip() for rut in (ruts_objetivo or [])}
    ruts_excluir = {str(rut).strip() for rut in (ruts_excluir or [])}
    page = context.new_page()
    url = FICHA_POR_QS.format(identificador) if es_qs else FICHA_POR_CODIGO.format(identificador)
    print(f"\n▶ {identificador}")
    try:
        page.goto(url, wait_until="networkidle", timeout=60000)
    except PWTimeout:
        print("   ⚠️ timeout ficha"); page.close()
        return {"ofertas": 0, "archivos": 0, "estado": "error_ficha"}

    html_ficha = page.content()
    mcod = re.search(r"\b(\d{3,7}-\d+-[A-Z]{1,2}\d+)\b", html_ficha)
    codigo = mcod.group(1) if mcod else str(identificador)[:25]
    carpeta_lic = SALIDA / re.sub(r"[^\w.\-]", "_", codigo)
    carpeta_lic.mkdir(parents=True, exist_ok=True)

    links = links_de_ficha(html_ficha)
    if not links["cuadro_ofertas"]:
        diagnostico = carpeta_lic / "diagnostico_sin_cuadro.html"
        diagnostico.write_text(html_ficha, encoding="utf-8")
        print("   ⚠️ Sin 'Cuadro de Ofertas' en la respuesta recibida")
        print(
            f"      url_final={page.url} | titulo={page.title()!r} "
            f"| html={len(html_ficha)} caracteres"
        )
        print(f"      diagnóstico: {diagnostico}")
        page.close()
        return {"codigo": codigo, "ofertas": 0, "archivos": 0, "estado": "sin_cuadro"}

    (carpeta_lic / "diagnostico_sin_cuadro.html").unlink(missing_ok=True)

    print("   → Cuadro de Ofertas...")
    try:
        page.goto(links["cuadro_ofertas"], wait_until="networkidle", timeout=60000)
        time.sleep(2.5)
    except PWTimeout:
        print("   ⚠️ timeout cuadro"); page.close()
        return {"codigo": codigo, "ofertas": 0, "archivos": 0, "estado": "error_cuadro"}

    frame, html0 = localizar_frame(page)
    if DIAG:
        Path("debug_cuadro.html").write_text(html0, encoding="utf-8")
        print(f"   [DIAG] frames={len(page.frames)} | HTML->debug_cuadro.html")

    n_pag = total_paginas(html0)
    print(f"   Páginas: {n_pag}")

    total_of = total_arch = 0
    registro = []
    for pag in range(1, n_pag + 1):
        if pag > 1:
            try:
                frame.evaluate(f'fnMovePage({pag}, "WucPagerGrid");')
                time.sleep(3.0)
                frame, _ = localizar_frame(page)
            except Exception as e:
                print(f"   ⚠️ página {pag}: {e}"); continue
        try:
            html_pag = frame.content()
        except Exception:
            html_pag = html0
        ofertas = parsear_ofertas(html_pag)
        # La grilla a veces llega vacia (carga lenta del iframe): se vuelve a leer antes de darla por vacia.
        intentos_grilla = 0
        while not ofertas and intentos_grilla < 3:
            intentos_grilla += 1
            print(f"   ↻ grilla sin ofertas; releyendo ({intentos_grilla}/3)...")
            time.sleep(4.0 * intentos_grilla)
            try:
                if intentos_grilla >= 2:
                    # Releer no basta: se vuelve a abrir el cuadro de ofertas (y la pagina de la grilla).
                    page.goto(links["cuadro_ofertas"], wait_until="networkidle", timeout=60000)
                    time.sleep(3.0)
                    frame, _ = localizar_frame(page)
                    if pag > 1:
                        frame.evaluate(f'fnMovePage({pag}, "WucPagerGrid");')
                        time.sleep(3.0)
                frame, html_pag = localizar_frame(page)
            except Exception as e:
                print(f"   ⚠️ no se pudo releer la grilla: {str(e)[:50]}")
                continue
            ofertas = parsear_ofertas(html_pag)
        print(f"   Página {pag}/{n_pag}: {len(ofertas)} ofertas")
        for of in ofertas:
            rut = of["rut"] or "sin_rut"
            prov = of.get("proveedor") or "?"
            # Nombre de carpeta = RUT + nombre del proveedor (limpio, sin caracteres raros)
            prov_limpio = re.sub(r"[^\w.\-]", "_", prov)[:60].strip("_") or "sin_nombre"
            rut_limpio = re.sub(r"[^\w.\-]", "_", rut)
            carpeta_of = carpeta_lic / f"{rut_limpio}__{prov_limpio}"
            carpeta_of.mkdir(exist_ok=True)
            total_of += 1
            print(f"     • {prov[:28]} | {of.get('total') or '?'} | RUT {rut}")
            registro.append({"codigo": codigo, "rut": rut, "proveedor": prov,
                             "total_oferta": of.get("total") or "",
                             "estado": of.get("estado") or "",
                             "anexos": ",".join(of["anexos"].keys()) or "ninguno"})
            (carpeta_of / "oferta.json").write_text(
                json.dumps({"rut": rut, "proveedor": prov, "total": of.get("total"),
                            "estado": of.get("estado"),
                            "enlaces_anexos": of.get("anexos", {})},
                           ensure_ascii=False, indent=2),
                encoding="utf-8")
            if ruts_objetivo and rut not in ruts_objetivo:
                continue
            if rut in ruts_excluir:
                continue
            omitidos_admin = []
            for tipo, url_popup in of["anexos"].items():
                ya_tiene_admin = tipo == "administrativo" and any(carpeta_of.glob("administrativo__*"))
                if tipo not in tipos_anexos and not ya_tiene_admin:
                    continue
                print(f"       - {tipo}")
                filtrar = tipo == "administrativo" and tipo not in tipos_anexos
                total_arch += descargar_popup(context, url_popup, carpeta_of, tipo,
                                              omitir=es_administrativo_sin_oferta if filtrar else None,
                                              omitidos=omitidos_admin if filtrar else None)
            # Hay proveedores que suben su cotizacion dentro del anexo "administrativo": si no quedo ningun
            # archivo del proveedor, se baja tambien ese anexo.
            if (admin_si_vacio and "administrativo" not in tipos_anexos and "administrativo" in of["anexos"]
                    and not carpeta_tiene_archivos(carpeta_of)):
                print("       - administrativo (el proveedor no tiene economico ni tecnico descargable)")
                total_arch += descargar_popup(context, of["anexos"]["administrativo"], carpeta_of, "administrativo",
                                              omitir=es_administrativo_sin_oferta, omitidos=omitidos_admin)
            if omitidos_admin:
                # Se deja constancia de lo que no se bajo (declaraciones, certificados, OC pasadas...).
                datos_oferta = json.loads((carpeta_of / "oferta.json").read_text(encoding="utf-8"))
                datos_oferta["administrativos_omitidos"] = sorted(set(omitidos_admin))
                (carpeta_of / "oferta.json").write_text(json.dumps(datos_oferta, ensure_ascii=False, indent=2),
                                                        encoding="utf-8")
                print(f"         ({len(set(omitidos_admin))} administrativos omitidos por nombre: no son la oferta)")

    if total_of == 0:
        # Una grilla vacia casi siempre es una respuesta incompleta del portal (en una corrida de 600 licitaciones
        # le paso al 9 %), no una licitacion sin ofertas. Se informa como error reintentable y no se escribe
        # metadata, para que otra pasada (--reparar) la vuelva a intentar en vez de darla por buena.
        page.close()
        print("   ⚠️ la grilla siguio vacia tras reintentos; se marca para reintentar")
        return {"codigo": codigo, "ofertas": 0, "archivos": 0, "estado": "sin_ofertas_leidas"}
    (carpeta_lic / "metadata.json").write_text(
        json.dumps({"codigo": codigo, "ofertas": total_of, "archivos": total_arch,
                    "paginas": n_pag}, ensure_ascii=False, indent=2), encoding="utf-8")
    if registro:
        with open(carpeta_lic / "resumen_ofertas.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=["codigo", "rut", "proveedor",
                                               "total_oferta", "estado", "anexos"])
            w.writeheader(); w.writerows(registro)
    page.close()
    print(f"   ✔ {total_of} ofertas, {total_arch} archivos")
    return {"codigo": codigo, "ofertas": total_of, "archivos": total_arch, "estado": "ok"}


def procesar_con_reintentos(browser, identificador, es_qs, tipos_anexos,
                            ruts_objetivo=None, max_intentos=2,
                            pausa_reintento=15, ruts_excluir=None, admin_si_vacio=True):
    reintentables = {"sin_cuadro", "error_ficha", "error_cuadro", "sin_ofertas_leidas"}
    resultado = {"ofertas": 0, "archivos": 0, "estado": "error_desconocido"}
    for intento in range(1, max(1, max_intentos) + 1):
        context = browser.new_context(accept_downloads=True)
        try:
            resultado = procesar_licitacion(
                context, identificador, es_qs=es_qs,
                tipos_anexos=tipos_anexos, ruts_objetivo=ruts_objetivo,
                ruts_excluir=ruts_excluir, admin_si_vacio=admin_si_vacio
            )
        finally:
            context.close()

        if resultado.get("estado") not in reintentables:
            return resultado
        if intento < max_intentos:
            print(
                f"   ↻ Respuesta anómala ({resultado.get('estado')}); "
                f"reintento {intento + 1}/{max_intentos} con sesión nueva..."
            )
            time.sleep(max(0, pausa_reintento))

    if resultado.get("estado") == "sin_cuadro":
        resultado["estado"] = "error_sin_cuadro"
    elif resultado.get("estado") == "sin_ofertas_leidas":
        resultado["estado"] = "error_sin_ofertas_leidas"
    return resultado


# ---------------------------------------------------------------------------
# Lectura/actualización del para_scrapear.csv (reanudable)
# ---------------------------------------------------------------------------
def leer_csv(ruta):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, encoding=enc, newline="") as fh:
                return list(csv.DictReader(fh)), fh.name
        except Exception:
            continue
    sys.exit(f"No pude leer {ruta}")


def guardar_csv(ruta, filas, campos):
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=campos)
        w.writeheader(); w.writerows(filas)


def ruts_sin_precio(ruta_resultados):
    """Devuelve proveedores cuyo resultado aun no contiene ningun precio."""
    ruta = Path(ruta_resultados)
    if not ruta.is_file():
        return set()
    try:
        resultados = json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return set()
    ruts = set()
    for resultado in resultados:
        productos = resultado.get("productos", [])
        tiene_precio = any(
            producto.get("precio_unitario") not in (None, "")
            or producto.get("precio_total") not in (None, "")
            for producto in productos
            if isinstance(producto, dict)
        )
        if not tiene_precio and resultado.get("rut"):
            ruts.add(str(resultado["rut"]).strip())
    return ruts


def ruts_con_precio(rutas_resultados):
    """RUTs que ya tienen algun precio en cualquiera de los JSON de resultados.
    Se combinan todos los archivos existentes (local y OpenAI historico)."""
    ruts = set()
    for ruta in rutas_resultados:
        ruta = Path(ruta)
        if not ruta.is_file():
            continue
        try:
            resultados = json.loads(ruta.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        for resultado in resultados:
            tiene_precio = any(
                producto.get("precio_unitario") not in (None, "")
                or producto.get("precio_total") not in (None, "")
                for producto in resultado.get("productos", [])
                if isinstance(producto, dict)
            )
            if tiene_precio and resultado.get("rut"):
                ruts.add(str(resultado["rut"]).strip())
    return ruts


def ruts_en_disco(carpeta_licitacion):
    """RUTs de los proveedores ya descargados en la carpeta de la licitacion."""
    ruts = set()
    if not Path(carpeta_licitacion).is_dir():
        return ruts
    for carpeta in Path(carpeta_licitacion).iterdir():
        oferta = carpeta / "oferta.json"
        if carpeta.is_dir() and oferta.is_file():
            try:
                rut = json.loads(oferta.read_text(encoding="utf-8")).get("rut")
            except (OSError, json.JSONDecodeError):
                rut = None
            if rut:
                ruts.add(str(rut).strip())
    return ruts


def licitacion_sin_datos(carpeta_licitacion):
    """True si la licitacion no se descargo (sin metadata) o quedo con 0 ofertas: casi siempre fue una grilla
    que el portal entrego vacia, no una licitacion sin oferentes."""
    meta = Path(carpeta_licitacion) / "metadata.json"
    if not meta.is_file():
        return True
    try:
        return int(json.loads(meta.read_text(encoding="utf-8")).get("ofertas") or 0) == 0
    except (OSError, json.JSONDecodeError, ValueError):
        return True


def faltantes_en_disco(carpeta_licitacion, tipos=("economico", "tecnico")):
    """RUTs con enlace a un anexo (de los tipos pedidos) que no se descargo a disco."""
    faltan = set()
    carpeta_licitacion = Path(carpeta_licitacion)
    if not carpeta_licitacion.is_dir():
        return faltan
    for carpeta in carpeta_licitacion.iterdir():
        oferta = carpeta / "oferta.json"
        if not (carpeta.is_dir() and oferta.is_file()):
            continue
        try:
            datos = json.loads(oferta.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        prefijos = [f.name.split("__")[0] for f in carpeta.iterdir() if f.is_file()]
        bajados = set(prefijos)
        enlaces = datos.get("enlaces_anexos") or {}
        sin_nada = not carpeta_tiene_archivos(carpeta)
        # Huecos en la numeracion (01, 02, 04...): una descarga fallo y no se reintento.
        con_huecos = False
        for t in ("economico", "tecnico", "administrativo"):
            indices = sorted(int(m.group(1)) for f in carpeta.iterdir() if f.is_file()
                             for m in [re.match(rf"{t}__(\d+)__", f.name)] if m)
            if indices and indices != list(range(1, indices[-1] + 1)):
                con_huecos = True
        # Un anexo con exactamente 6 archivos (o multiplo) puede estar truncado: el popup pagina de a 6.
        posible_truncado = con_huecos or any(
            t in enlaces and prefijos.count(t) and prefijos.count(t) % ARCHIVOS_POR_PAGINA_POPUP == 0
            for t in (*tipos, "administrativo"))
        if (any(t in enlaces and t not in bajados for t in tipos) or posible_truncado
                or (sin_nada and "administrativo" in enlaces)):
            if datos.get("rut"):
                faltan.add(str(datos["rut"]).strip())
    return faltan


def recontar_archivos(carpeta_licitacion):
    """Deja metadata.json con el conteo real de archivos en disco (tras una reparacion parcial)."""
    carpeta_licitacion = Path(carpeta_licitacion)
    meta = carpeta_licitacion / "metadata.json"
    if not meta.is_file():
        return 0
    total = sum(1 for f in carpeta_licitacion.rglob("*")
                if f.is_file() and f.suffix.lower() not in (".json", ".csv", ".html"))
    try:
        datos = json.loads(meta.read_text(encoding="utf-8"))
        datos["archivos"] = total
        datos["ofertas"] = sum(1 for c in carpeta_licitacion.iterdir()
                               if c.is_dir() and (c / "oferta.json").is_file())
        meta.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, json.JSONDecodeError):
        pass
    return total


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    global DIAG
    ap = argparse.ArgumentParser(description="Scraper de ofertas integrado con para_scrapear.csv")
    ap.add_argument("--csv", help="para_scrapear.csv (columna 'codigo').")
    ap.add_argument("--codigo", help="Una licitación suelta por código (modo test).")
    ap.add_argument("--qs", help="Una licitación suelta por qs (modo test).")
    ap.add_argument("--limite", type=int, help="Procesar solo las primeras N pendientes.")
    ap.add_argument("--ver", action="store_true", help="Muestra el navegador.")
    ap.add_argument("--diag", action="store_true")
    ap.add_argument("--todos-anexos", action="store_true", help="Incluye el administrativo.")
    ap.add_argument("--incluye-tecnicos", action="store_true",
                    help="Descarga economicos y tecnicos en la primera pasada.")
    ap.add_argument("--solo-tecnicos", action="store_true",
                    help="Descarga solo tecnicos para completar ofertas ya scrapeadas.")
    ap.add_argument("--solo-tecnicos-fallback", action="store_true",
                    help="Descarga tecnicos para proveedores que aun no tienen precio en los resultados IA.")
    ap.add_argument("--reintentar-errores", action="store_true",
                    help="Reintenta también las marcadas 'error' (default: solo 'pendiente').")
    ap.add_argument("--max-intentos", type=int, default=2,
                    help="Intentos por licitación con sesiones independientes (default: 2).")
    ap.add_argument("--pausa-reintento", type=float, default=15,
                    help="Segundos antes de reintentar una respuesta anómala (default: 15).")
    ap.add_argument("--salida", default="ofertas",
                    help="Carpeta donde se guardan las licitaciones (default: ofertas).")
    ap.add_argument("--sin-admin-si-vacio", action="store_true",
                    help="No descargar el anexo administrativo de proveedores sin economico ni tecnico.")
    ap.add_argument("--solo-administrativos", action="store_true",
                    help="Descarga solo los anexos administrativos (para comprobar si traen la oferta).")
    ap.add_argument("--ruts", help="Con --codigo/--qs: RUTs separados por coma a los que limitar la descarga.")
    ap.add_argument("--reparar", action="store_true",
                    help="Repara descargas incompletas: reprocesa las licitaciones con error o sin "
                         "metadata y baja solo los anexos economicos/tecnicos que faltan en disco.")
    ap.add_argument("--pausa-licitaciones", type=float, default=3,
                    help="Segundos entre licitaciones para reducir bloqueos (default: 3).")
    args = ap.parse_args()
    DIAG = args.diag
    if args.solo_tecnicos and (args.incluye_tecnicos or args.todos_anexos or args.solo_tecnicos_fallback):
        sys.exit("No combines --solo-tecnicos con otros modos de anexos.")
    if args.solo_tecnicos_fallback and (args.incluye_tecnicos or args.todos_anexos):
        sys.exit("No combines --solo-tecnicos-fallback con --incluye-tecnicos/--todos-anexos.")
    tipos_anexos = {"economico"}
    if args.solo_administrativos:
        tipos_anexos = {"administrativo"}
    if args.solo_tecnicos:
        tipos_anexos = {"tecnico"}
    if args.incluye_tecnicos:
        tipos_anexos.add("tecnico")
    if args.todos_anexos:
        tipos_anexos.update(ANEXOS)
    if args.solo_tecnicos_fallback:
        tipos_anexos = {"tecnico"}

    if not any([args.csv, args.codigo, args.qs]):
        sys.exit("Indica --csv para_scrapear.csv  o  --codigo <cod>  o  --qs <qs>")

    global SALIDA
    SALIDA = Path(args.salida)
    SALIDA.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.ver)

        # ---- Modo licitación suelta (test) ----
        if args.codigo or args.qs:
            ident = args.qs or args.codigo
            procesar_con_reintentos(
                browser, ident, es_qs=bool(args.qs), tipos_anexos=tipos_anexos,
                ruts_objetivo=[r.strip() for r in args.ruts.split(",")] if args.ruts else None,
                max_intentos=args.max_intentos, pausa_reintento=args.pausa_reintento,
                admin_si_vacio=not args.sin_admin_si_vacio
            )
            browser.close()
            return

        # ---- Modo lote desde CSV (reanudable) ----
        filas, _ = leer_csv(args.csv)
        if not filas:
            sys.exit("El CSV está vacío.")
        campos = list(filas[0].keys())
        # asegurar columnas de estado
        for c in ("scraping_estado", "n_ofertas", "n_archivos"):
            if c not in campos:
                campos.append(c)
                for f in filas:
                    f.setdefault(c, "")

        estados_saltar = {"ok"} if args.reintentar_errores else {"ok"}
        # (por defecto procesamos 'pendiente' y '' ; con --reintentar-errores
        #  también reintentamos las 'error')
        def pendiente(f):
            est = (f.get("scraping_estado") or "").strip().lower()
            if est == "ok":
                return False
            if est.startswith("error") and not args.reintentar_errores:
                return False
            return True

        reparaciones = {}
        if args.reparar:
            tipos_anexos = {"economico", "tecnico"}
            for f in filas:
                cod = (f.get("codigo") or "").strip()
                carpeta = SALIDA / cod
                est = (f.get("scraping_estado") or "").strip().lower()
                if est.startswith("error") or licitacion_sin_datos(carpeta):
                    reparaciones[cod] = None  # licitacion completa
                else:
                    faltan = faltantes_en_disco(carpeta)
                    if faltan:
                        reparaciones[cod] = faltan
            args.reintentar_errores = True

        if args.reparar:
            pendientes = [f for f in filas if (f.get("codigo") or "").strip() in reparaciones]
        elif args.solo_tecnicos or args.solo_tecnicos_fallback:
            pendientes = list(filas)
        else:
            pendientes = [f for f in filas if pendiente(f)]
        if args.limite:
            pendientes = pendientes[:args.limite]

        total = len(pendientes)
        print("=" * 66)
        print(f"  LOTE | CSV: {args.csv}")
        print(f"  Total en CSV: {len(filas)} | Pendientes a procesar: {total}")
        print(f"  Anexos descargados: {', '.join(sorted(tipos_anexos))}")
        print("=" * 66)

        procesadas = 0
        for i, fila in enumerate(pendientes, 1):
            codigo = (fila.get("codigo") or "").strip()
            if not codigo:
                continue
            print(f"\n[{i}/{total}] {codigo}")
            ruts_objetivo = None
            ruts_excluir = None
            if args.reparar:
                ruts_objetivo = reparaciones.get(codigo)
            if args.solo_tecnicos_fallback:
                # Se descargan tecnicos para todos los proveedores SALVO los que ya
                # tienen precio en algun resultado (local o OpenAI historico). Asi
                # tambien entran los que aun no se han procesado con IA.
                carpeta_lic = SALIDA / codigo
                ruts_excluir = ruts_con_precio([
                    carpeta_lic / "extraccion_ia.json",
                    carpeta_lic / "extraccion_openai.json",
                ])
                descargados = ruts_en_disco(carpeta_lic)
                if descargados and descargados <= ruts_excluir:
                    print("   ↷ Todos los proveedores ya tienen precio; no hace falta fallback técnico")
                    continue
            try:
                r = procesar_con_reintentos(
                    browser, codigo, es_qs=False, tipos_anexos=tipos_anexos,
                    ruts_objetivo=ruts_objetivo, ruts_excluir=ruts_excluir,
                    max_intentos=args.max_intentos, pausa_reintento=args.pausa_reintento,
                    admin_si_vacio=not args.sin_admin_si_vacio
                )
                fila["scraping_estado"] = r.get("estado", "ok")
                fila["n_ofertas"] = r.get("ofertas", 0)
                fila["n_archivos"] = r.get("archivos", 0)
                if args.reparar and r.get("estado") == "ok":
                    fila["n_archivos"] = recontar_archivos(SALIDA / (r.get("codigo") or codigo))
            except Exception as e:
                print(f"   ❌ error: {str(e)[:60]}")
                fila["scraping_estado"] = f"error: {str(e)[:40]}"
            procesadas += 1
            # Guardar el CSV tras CADA licitación (reanudable ante cortes)
            guardar_csv(args.csv, filas, campos)
            if i < total and args.pausa_licitaciones > 0:
                time.sleep(args.pausa_licitaciones)

        browser.close()

        # Resumen
        ok = sum(1 for f in filas if (f.get("scraping_estado") or "").lower() == "ok")
        err = sum(1 for f in filas if (f.get("scraping_estado") or "").lower().startswith("error"))
        print("\n" + "=" * 66)
        print(f"  LISTO. Procesadas esta corrida: {procesadas}")
        print(f"  Total OK: {ok} | Con error: {err} | CSV actualizado: {args.csv}")
        print(f"  Carpeta de anexos: {SALIDA.resolve()}")
        print("=" * 66)


if __name__ == "__main__":
    main()

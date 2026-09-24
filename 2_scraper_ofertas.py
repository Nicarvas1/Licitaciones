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


def descargar_popup(context, url_popup, carpeta, prefijo):
    bajados = 0
    popup = context.new_page()
    try:
        popup.goto(url_popup, wait_until="networkidle", timeout=45000)
    except PWTimeout:
        popup.close(); return 0
    botones = popup.query_selector_all("input[id^='DWNL_grdId_'][id$='_search']")
    if not botones:
        botones = popup.query_selector_all("input[type='image'][src*='ver.gif']")
    for idx, b in enumerate(botones, 1):
        try:
            with popup.expect_download(timeout=45000) as dl:
                b.click()
            d = dl.value
            nom = re.sub(r"[^\w\s.\-()]", "_", d.suggested_filename or f"{prefijo}_{idx}")[:150]
            dest = carpeta / f"{prefijo}__{idx:02d}__{nom}"
            d.save_as(str(dest))
            bajados += 1
            print(f"        ✅ {dest.name[:58]}")
        except Exception as e:
            print(f"        ⚠️ archivo {idx}: {str(e)[:40]}")
        time.sleep(0.4)
    popup.close()
    return bajados


# ---------------------------------------------------------------------------
# Procesar UNA licitación (por código o qs)
# ---------------------------------------------------------------------------
def procesar_licitacion(context, identificador, es_qs=False,
                        tipos_anexos=None, ruts_objetivo=None, ruts_excluir=None):
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
            for tipo, url_popup in of["anexos"].items():
                if tipo not in tipos_anexos:
                    continue
                print(f"       - {tipo}")
                total_arch += descargar_popup(context, url_popup, carpeta_of, tipo)

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
                            pausa_reintento=15, ruts_excluir=None):
    reintentables = {"sin_cuadro", "error_ficha", "error_cuadro"}
    resultado = {"ofertas": 0, "archivos": 0, "estado": "error_desconocido"}
    for intento in range(1, max(1, max_intentos) + 1):
        context = browser.new_context(accept_downloads=True)
        try:
            resultado = procesar_licitacion(
                context, identificador, es_qs=es_qs,
                tipos_anexos=tipos_anexos, ruts_objetivo=ruts_objetivo,
                ruts_excluir=ruts_excluir
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
    ap.add_argument("--pausa-licitaciones", type=float, default=3,
                    help="Segundos entre licitaciones para reducir bloqueos (default: 3).")
    args = ap.parse_args()
    DIAG = args.diag
    if args.solo_tecnicos and (args.incluye_tecnicos or args.todos_anexos or args.solo_tecnicos_fallback):
        sys.exit("No combines --solo-tecnicos con otros modos de anexos.")
    if args.solo_tecnicos_fallback and (args.incluye_tecnicos or args.todos_anexos):
        sys.exit("No combines --solo-tecnicos-fallback con --incluye-tecnicos/--todos-anexos.")
    tipos_anexos = {"economico"}
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

    SALIDA.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.ver)

        # ---- Modo licitación suelta (test) ----
        if args.codigo or args.qs:
            ident = args.qs or args.codigo
            procesar_con_reintentos(
                browser, ident, es_qs=bool(args.qs), tipos_anexos=tipos_anexos,
                max_intentos=args.max_intentos, pausa_reintento=args.pausa_reintento
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

        if args.solo_tecnicos or args.solo_tecnicos_fallback:
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
                    max_intentos=args.max_intentos, pausa_reintento=args.pausa_reintento
                )
                fila["scraping_estado"] = r.get("estado", "ok")
                fila["n_ofertas"] = r.get("ofertas", 0)
                fila["n_archivos"] = r.get("archivos", 0)
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

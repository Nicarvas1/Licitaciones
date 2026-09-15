#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 1_filtrar_licitaciones.py   (ETAPA 0 del pipeline)
=============================================================================
Toma el/los CSV descargado(s) del buscador de Mercado Público y produce una
lista LIMPIA de licitaciones de CÓMPUTO, lista para el scraper de ofertas.

ENTRADA (CSV del buscador, separado por ',' o ';'):
  IDLicitacion;NombreLicitacion;Tipo;Estado;FechaPublicacion;Descripcion;
  Moneda;TipoPresupuesto;TipoMonto;MontoLicitacion;Organismo

QUÉ HACE:
  1. Filtra por ESTADO: conserva Publicada / Cerrada / Adjudicada (según config).
  2. Filtra por RUBRO: conserva cómputo real, descarta tablet/arriendo/etc.
  3. Marca cuáles son 'scrapeables' (solo Adjudicada tiene ofertas visibles).
  4. Normaliza el monto (fijo o rango UTM).
  5. Arma la URL de la ficha (idlicitacion=CODIGO).
  6. Genera:
       - licitaciones_computo.csv   (todas las que pasan el filtro)
       - para_scrapear.csv          (solo adjudicadas -> alimenta el scraper)

USO:
  python 1_filtrar_licitaciones.py --in busqueda.csv
  python 1_filtrar_licitaciones.py --in *.csv --out-dir salida/
  python 1_filtrar_licitaciones.py --in busqueda.csv --demo   (usa ejemplo)
=============================================================================
"""

import argparse
import csv
import glob
import re
import sys
from pathlib import Path

BASE_FICHA = "https://www.mercadopublico.cl/Procurement/Modules/RFB/DetailsAcquisition.aspx?idlicitacion={}"

# ---------------------------------------------------------------------------
# FILTRO POR ESTADO
# ---------------------------------------------------------------------------
# Estados que conservamos (el usuario pidió publicada, cerrada, adjudicada).
ESTADOS_CONSERVAR = [
    re.compile(r"adjudicada", re.I),
    re.compile(r"publicada", re.I),
    re.compile(r"cerrada", re.I),
]
# Estados con OFERTAS VISIBLES para el scraper de anexos.
# Cerrada y Adjudicada ya pasaron la apertura -> las ofertas están públicas.
# Publicada NO (aún en secreto, no ha cerrado la recepción de ofertas).
ESTADO_SCRAPEABLE = re.compile(r"adjudicada|cerrada", re.I)
# Estados que SIEMPRE descartamos (por claridad, aunque no matcheen arriba).
ESTADOS_DESCARTAR = re.compile(r"desierta|cancelada|revocada|suspendida|sin ofertas", re.I)

# ---------------------------------------------------------------------------
# FILTRO POR RUBRO (cómputo real vs. lo que no sirve)
# ---------------------------------------------------------------------------
INCLUIR = re.compile(
    r"comput|notebook|laptop|port[aá]til|escritorio|all.?in.?one|\baio\b|"
    r"workstation|servidor|equipamiento\s+comput|equipos?\s+comput|"
    r"monitor|impresora|desktop|pc\b", re.I)

EXCLUIR = re.compile(
    r"\btablet|\bipad|arriendo|renovaci[oó]n\s+de\s+garant|"
    r"servicio\s+de\s+mantenc|solo\s+tinta|\btoner\b\s*$|cartucho\s*$", re.I)


def clasificar_estado(estado):
    """Devuelve (conservar: bool, scrapeable: bool)."""
    e = estado or ""
    if ESTADOS_DESCARTAR.search(e):
        return False, False
    conservar = any(p.search(e) for p in ESTADOS_CONSERVAR)
    scrapeable = bool(ESTADO_SCRAPEABLE.search(e))
    return conservar, scrapeable


def es_computo(nombre, descripcion):
    """True si parece cómputo real y NO está en la lista de exclusión."""
    texto = f"{nombre or ''} {descripcion or ''}"
    if EXCLUIR.search(texto):
        return False
    return bool(INCLUIR.search(texto))


def normalizar_monto(tipo_monto, monto):
    """
    Normaliza el campo de monto. Puede ser:
      - número fijo: '21000000'
      - rango UTM: 'Entre 100 y 1000 UTM', 'Menor a 100 UTM', 'Mayor a 5.000 UTM'
    Devuelve (monto_clp_num_or_None, monto_texto).
    """
    valor = (monto or "").strip()
    # ¿es un número puro?
    solo_num = re.sub(r"[^\d]", "", valor)
    if valor and solo_num == valor.replace(".", "").replace(",", ""):
        try:
            return int(solo_num), valor
        except ValueError:
            pass
    # es un rango en UTM u otra cosa -> lo dejamos como texto
    return None, valor


# ---------------------------------------------------------------------------
# PROCESO
# ---------------------------------------------------------------------------
def procesar(rutas, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    filas_ok = []       # todas las que pasan (computo + estado)
    vistos = set()      # dedup por IDLicitacion
    stats = {"total": 0, "descartada_estado": 0, "descartada_rubro": 0,
             "duplicada": 0, "conservada": 0, "scrapeable": 0}

    for ruta in rutas:
        # Mercado Publico ha entregado archivos separados por ',' y por ';'.
        for enc in ("utf-8-sig", "latin-1"):
            try:
                with open(ruta, encoding=enc, newline="") as fh:
                    muestra = fh.read(8192)
                    fh.seek(0)
                    try:
                        delimitador = csv.Sniffer().sniff(muestra, delimiters=",;").delimiter
                    except csv.Error:
                        primera_linea = muestra.splitlines()[0] if muestra else ""
                        delimitador = ";" if primera_linea.count(";") > primera_linea.count(",") else ","
                    reader = csv.DictReader(fh, delimiter=delimitador)
                    encabezados = {(campo or "").strip() for campo in (reader.fieldnames or [])}
                    if "IDLicitacion" not in encabezados:
                        raise ValueError("El CSV no contiene la columna IDLicitacion")
                    filas = list(reader)
                break
            except Exception:
                continue
        else:
            print(f"⚠️ No pude leer {ruta}"); continue

        print(f"Leyendo {ruta}: {len(filas)} filas | separador={delimitador!r}")
        for row in filas:
            stats["total"] += 1
            # normalizar claves (por si vienen con espacios)
            row = {(k or "").strip(): (v or "").strip() for k, v in row.items()}
            codigo = row.get("IDLicitacion", "")
            if not codigo:
                continue
            if codigo in vistos:
                stats["duplicada"] += 1
                continue

            estado = row.get("Estado", "")
            conservar, scrapeable = clasificar_estado(estado)
            if not conservar:
                stats["descartada_estado"] += 1
                continue

            nombre = row.get("NombreLicitacion", "")
            desc = row.get("Descripcion", "")
            if not es_computo(nombre, desc):
                stats["descartada_rubro"] += 1
                continue

            vistos.add(codigo)
            monto_num, monto_txt = normalizar_monto(row.get("TipoMonto", ""),
                                                    row.get("MontoLicitacion", ""))
            filas_ok.append({
                "codigo": codigo,
                "nombre": nombre,
                "tipo": row.get("Tipo", ""),
                "estado": estado,
                "scrapeable": "si" if scrapeable else "no",
                "fecha_publicacion": row.get("FechaPublicacion", ""),
                "monto_num": monto_num if monto_num is not None else "",
                "monto_texto": monto_txt,
                "organismo": row.get("Organismo", ""),
                "url_ficha": BASE_FICHA.format(codigo),
                # estados del pipeline (para reanudar)
                "scraping_estado": "pendiente" if scrapeable else "no_aplica",
                "ia_estado": "pendiente" if scrapeable else "no_aplica",
            })
            stats["conservada"] += 1
            if scrapeable:
                stats["scrapeable"] += 1

    # --- Escribir salidas ---
    campos = ["codigo", "nombre", "tipo", "estado", "scrapeable",
              "fecha_publicacion", "monto_num", "monto_texto", "organismo",
              "url_ficha", "scraping_estado", "ia_estado"]

    p_all = out / "licitaciones_computo.csv"
    with open(p_all, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=campos)
        w.writeheader(); w.writerows(filas_ok)

    # para_scrapear = cerradas + adjudicadas (tienen ofertas visibles)
    scrapeables = [f for f in filas_ok if f["scrapeable"] == "si"]
    p_scr = out / "para_scrapear.csv"
    with open(p_scr, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=campos)
        w.writeheader(); w.writerows(scrapeables)

    # --- Reporte ---
    print("\n" + "=" * 60)
    print("  RESULTADO DEL FILTRADO")
    print("=" * 60)
    print(f"  Total leídas          : {stats['total']}")
    print(f"  Descartadas x estado  : {stats['descartada_estado']}")
    print(f"  Descartadas x rubro   : {stats['descartada_rubro']}")
    print(f"  Duplicadas            : {stats['duplicada']}")
    print(f"  ── Conservadas        : {stats['conservada']}")
    print(f"     scrapeables (cerradas + adjudicadas): {stats['scrapeable']}")
    print(f"     no scrapeables (publicadas)         : {stats['conservada'] - stats['scrapeable']}")
    print("=" * 60)
    print(f"\n  Archivos generados:")
    print(f"    {p_all}     (todas las de cómputo: publicadas/cerradas/adjudicadas)")
    print(f"    {p_scr}     (cerradas + adjudicadas -> para el scraper de ofertas)")

    if filas_ok:
        print(f"\n  Muestra (primeras 8):")
        print(f"    {'CÓDIGO':<16}{'SCRAP':<7}{'ESTADO':<14}NOMBRE")
        for f in filas_ok[:8]:
            est_corto = f["estado"][:12]
            print(f"    {f['codigo']:<16}{f['scrapeable']:<7}{est_corto:<14}{f['nombre'][:34]}")


def main():
    ap = argparse.ArgumentParser(description="Filtra el CSV del buscador -> lista de cómputo.")
    ap.add_argument("--in", dest="entrada", nargs="+", help="CSV(s) del buscador. Acepta comodines.")
    ap.add_argument("--out-dir", default=".", help="Carpeta de salida.")
    ap.add_argument("--demo", action="store_true", help="Usa un CSV de ejemplo.")
    args = ap.parse_args()

    if args.demo:
        rutas = ["/tmp/licitaciones_ejemplo.csv"]
    elif args.entrada:
        rutas = []
        for patron in args.entrada:
            rutas.extend(glob.glob(patron))
        if not rutas:
            sys.exit(f"No encontré archivos: {args.entrada}")
    else:
        sys.exit("Indica --in <archivo.csv>  o  --demo")

    procesar(rutas, args.out_dir)


if __name__ == "__main__":
    main()

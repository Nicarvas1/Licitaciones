#!/usr/bin/env python3
"""Auditoria de una extraccion contra los documentos originales (sin modelo de IA).

Para cada proveedor con `extraccion_ia.json` relee sus anexos y comprueba:

  FALTA     filas del anexo economico (cantidad x unitario = total) que son equipos del
            alcance y ningun producto recoge; se incorporan (marcadas) en la hoja Productos.
  MEZCLA    precio unitario/total que no aparece en ningun documento del proveedor, o
            precio repetido entre categorias (senal de datos copiados de otra fila).
  MARCA     marca que no aparece escrita en los documentos del proveedor.
  SIN LEER  archivos no leidos (OCR pendiente, formato no soportado, fallo del modelo).

Un producto es "verificado" solo si pasa la validacion y cada valor esta en los documentos.
No modifica los checkpoints: escribe un Excel aparte.

Ejemplo:
  python 5_auditar_extraccion.py --dir lotes_z8/2025-11/ofertas --excel auditoria.xlsx
  python 5_auditar_extraccion.py --dir lotes_z8/*/ofertas lotes_z4/*/ofertas --workers 4
"""

from __future__ import annotations

import argparse
import glob
import importlib.util
import json
import re
import sys
import types
import unicodedata
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
_SPEC = importlib.util.spec_from_file_location("extractor_ia", ROOT / "3_extraer_ia.py")
E = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(E)

ARGS_LECTURA = types.SimpleNamespace(
    max_paginas=200, ocr=False, vision=False, tesseract_cmd=None, idioma_ocr="spa+eng",
    dpi_ocr=200, vision_min_filas=2, max_filas_excel=5000, max_columnas_excel=40,
)
PATRON_NUMERO = re.compile(r"\d[\d.,]*")
ALIAS_MARCA = {"hp": ("hp", "hewlett"), "asus": ("asus",), "lenovo": ("lenovo", "thinkpad", "ideapad")}


def sin_tildes(texto):
    return unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii").lower()


def numeros_del_texto(texto):
    encontrados = set()
    for coincidencia in PATRON_NUMERO.finditer(texto):
        valor = E.normalizar_numero(coincidencia.group(0).rstrip(".,"))
        if isinstance(valor, (int, float)):
            encontrados.add(round(valor))
    return encontrados


def marca_en_texto(marca, texto_plano):
    clave = sin_tildes(marca).strip()
    if not clave:
        return True
    return any(alias in texto_plano for alias in ALIAS_MARCA.get(clave, (clave,)))


def leer_proveedor(carpeta):
    """(texto_total, bloques_economicos, archivos_sin_texto)"""
    texto_total, bloques, sin_texto = [], [], []
    for ruta in E.listar_archivos_oferta(carpeta):
        try:
            texto, estado = E.extraer_archivo(ruta, ARGS_LECTURA)
        except Exception as exc:  # documento corrupto: se informa, no se detiene la auditoria
            sin_texto.append(f"{ruta.name} (error_lectura: {exc})")
            continue
        if not texto.strip():
            if ruta.suffix.lower() in (".pdf", ".png", ".jpg", ".jpeg", ".webp"):
                sin_texto.append(f"{ruta.name} ({estado or 'sin_texto'})")
            continue
        texto_total.append(texto)
        if E.tipo_documental(ruta.name) == "economico":
            bloques.append({"archivo": str(ruta.relative_to(carpeta)), "pagina": None, "texto": texto})
    return "\n".join(texto_total), bloques, sin_texto


def auditar_proveedor(tarea):
    carpeta_licitacion, proveedor_json = tarea
    carpeta_licitacion = Path(carpeta_licitacion)
    rut, nombre = proveedor_json.get("rut", ""), proveedor_json.get("proveedor", "")
    candidatas = [c for c in carpeta_licitacion.iterdir() if c.is_dir() and rut and rut in c.name]
    productos = [
        {k: v for k, v in p.items() if k not in ("alertas", "estado_validacion")}
        for p in proveedor_json.get("productos", [])
    ]
    parciales = proveedor_json.get("productos_parciales", [])
    salida = {"licitacion": carpeta_licitacion.name, "proveedor": nombre, "rut": rut,
              "productos": [], "sin_leer": [], "notas": []}
    if not candidatas:
        salida["notas"].append("carpeta_del_proveedor_no_encontrada")
        texto, bloques, sin_texto = "", [], []
    else:
        texto, bloques, sin_texto = leer_proveedor(candidatas[0])
    numeros = numeros_del_texto(texto)
    texto_plano = sin_tildes(texto)

    # Archivos no leidos: sin texto, OCR pendiente, sin soporte o con fallo del modelo.
    salida["sin_leer"].extend(sin_texto)
    for registro in proveedor_json.get("resultados_archivos", []):
        if E.archivo_fallo_modelo(registro):
            salida["sin_leer"].append(f"{registro.get('archivo')} (fallo_modelo)")
    salida["sin_leer"].extend(f"{a} (necesita_ocr)" for a in proveedor_json.get("necesita_ocr", []))
    carpeta_proveedor = candidatas[0] if candidatas else None
    for archivo in proveedor_json.get("no_soportados", []):
        if Path(archivo).suffix.lower() in (".zip", ".rar", ".7z") and carpeta_proveedor:
            extraido = carpeta_proveedor / "_extraidos" / Path(archivo).stem
            if extraido.is_dir() and any(extraido.rglob("*.*")):
                continue  # comprimido ya expandido: sus archivos se leyeron aparte
        salida["sin_leer"].append(f"{archivo} (no_soportado)")
    salida["sin_leer"] = sorted(set(salida["sin_leer"]))
    legible = bool(texto.strip())

    productos = E.rescatar_filas_economicas(productos, parciales, bloques, nombre, rut)
    validados = E.validar_productos(productos, parciales, proveedor_json.get("total_oferta"))
    for producto in validados:
        motivos = []
        if producto.get("estado_validacion") != "ok":
            motivos.append("validacion_" + str(producto.get("estado_validacion")))
        if legible:
            unitario, total, cantidad = producto.get("precio_unitario"), producto.get("precio_total"), producto.get("cantidad")
            unitario_visible = isinstance(unitario, (int, float)) and round(unitario) in numeros
            total_visible = isinstance(total, (int, float)) and round(total) in numeros
            coherente = all(isinstance(v, (int, float)) and v > 0 for v in (unitario, total, cantidad)) and                 abs(cantidad * unitario - total) <= max(2, total * 0.001)
            # Un valor que el documento no escribe pero que sale de cantidad x unitario (o total / cantidad) esta
            # respaldado por el otro valor; solo se marca si ninguno de los dos aparece.
            if isinstance(unitario, (int, float)) and not unitario_visible and not (total_visible and coherente):
                motivos.append("unitario_no_aparece_en_documentos")
            if isinstance(total, (int, float)) and not total_visible and not (unitario_visible and coherente):
                motivos.append("total_no_aparece_en_documentos")
            marca = producto.get("marca")
            if marca and not marca_en_texto(marca, texto_plano) and                     E.marca_desde_descripcion(f"{producto.get('producto') or ''} {producto.get('modelo') or ''}") != marca:
                motivos.append("marca_no_aparece_en_documentos")
        else:
            motivos.append("sin_texto_para_verificar")
        if producto.get("rescatado_fila_economica"):
            motivos.append("rescatado_de_fila_economica")
        nivel = "verificado" if not motivos else (
            "no_verificable" if motivos == ["sin_texto_para_verificar"] else "revisar")
        producto["nivel_auditoria"] = nivel
        producto["motivos_auditoria"] = motivos
        producto["licitacion_codigo"] = carpeta_licitacion.name
        salida["productos"].append(producto)
    if not validados:
        salida["notas"].append("proveedor_sin_productos")
    return salida


def cargar_tareas(directorios, limite):
    tareas, vistos = [], set()
    for patron in directorios:
        for directorio in sorted(glob.glob(patron)) or [patron]:
            base = Path(directorio)
            for archivo in sorted(base.glob("*/extraccion_ia.json")):
                if archivo in vistos:
                    continue
                vistos.add(archivo)
                try:
                    datos = json.loads(archivo.read_text(encoding="utf-8"))
                except Exception:
                    continue
                for proveedor in datos if isinstance(datos, list) else []:
                    tareas.append((str(archivo.parent), proveedor))
            if limite and len({t[0] for t in tareas}) >= limite:
                break
    if limite:
        permitidas = sorted({t[0] for t in tareas})[:limite]
        tareas = [t for t in tareas if t[0] in permitidas]
    return tareas


def escribir_excel(resultados, ruta, metadata):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    libro = Workbook()
    hojas = {}

    def hoja(nombre, columnas):
        ws = libro.active if not hojas else libro.create_sheet()
        ws.title = nombre
        for c, titulo in enumerate(columnas, 1):
            celda = ws.cell(1, c, titulo)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="1F4E78")
        ws.freeze_panes = "A2"
        hojas[nombre] = ws
        return ws

    cols_prod = ["Licitación", "Fecha licitación", "Proveedor", "Producto", "Marca producto", "Cantidad",
                 "Precio unitario", "Precio total", "Nivel auditoría", "Motivos", "Estado validación",
                 "Alertas", "Archivo precio", "Archivo producto"]
    ws_p = hoja("Productos", cols_prod)
    ws_r = hoja("Resumen", ["Licitación", "Proveedores", "Productos", "Verificados", "% verificados",
                            "A revisar", "No verificables", "Rescatados", "Proveedores sin productos",
                            "Proveedores con archivos sin leer"])
    ws_s = hoja("Sin_leer", ["Licitación", "Proveedor", "RUT", "Archivo / motivo"])
    ws_n = hoja("Sin_productos", ["Licitación", "Proveedor", "RUT", "Archivos sin leer"])

    por_licitacion = {}
    for r in resultados:
        agregado = por_licitacion.setdefault(r["licitacion"], {
            "prov": 0, "prod": 0, "ver": 0, "rev": 0, "nov": 0, "res": 0, "sinprod": 0, "sinleer": 0})
        agregado["prov"] += 1
        if "proveedor_sin_productos" in r["notas"]:
            agregado["sinprod"] += 1
            ws_n.append([r["licitacion"], r["proveedor"], r["rut"], "; ".join(r["sin_leer"])])
        if r["sin_leer"]:
            agregado["sinleer"] += 1
            for archivo in r["sin_leer"]:
                ws_s.append([r["licitacion"], r["proveedor"], r["rut"], archivo])
        for p in r["productos"]:
            agregado["prod"] += 1
            agregado[{"verificado": "ver", "revisar": "rev", "no_verificable": "nov"}[p["nivel_auditoria"]]] += 1
            agregado["res"] += bool(p.get("rescatado_fila_economica"))
            meta = metadata.get(r["licitacion"], {})
            fuentes = p.get("fuente_precio") or p.get("archivo_fuente")
            ws_p.append([
                r["licitacion"], meta.get("fecha_publicacion", ""), r["proveedor"], p.get("producto"),
                p.get("marca"), p.get("cantidad"), p.get("precio_unitario"), p.get("precio_total"),
                p["nivel_auditoria"], "; ".join(p["motivos_auditoria"]), p.get("estado_validacion"),
                "; ".join(p.get("alertas") or []), fuentes, p.get("fuente_producto") or p.get("archivo_fuente"),
            ])
    total = {k: 0 for k in ("prov", "prod", "ver", "rev", "nov", "res", "sinprod", "sinleer")}
    for codigo, a in sorted(por_licitacion.items()):
        ws_r.append([codigo, a["prov"], a["prod"], a["ver"], round(a["ver"] / a["prod"], 3) if a["prod"] else None,
                     a["rev"], a["nov"], a["res"], a["sinprod"], a["sinleer"]])
        for k in total:
            total[k] += a[k]
    ws_r.append(["TOTAL", total["prov"], total["prod"], total["ver"],
                 round(total["ver"] / total["prod"], 3) if total["prod"] else None,
                 total["rev"], total["nov"], total["res"], total["sinprod"], total["sinleer"]])
    for ws in hojas.values():
        ws.auto_filter.ref = ws.dimensions
    libro.save(ruta)
    return total


def cargar_metadata(directorios):
    metadata = {}
    for patron in directorios:
        for directorio in glob.glob(patron) or [patron]:
            csv_path = Path(directorio).parent / "para_scrapear.csv"
            if csv_path.is_file():
                try:
                    metadata.update(E.cargar_metadata_licitaciones(csv_path))
                except Exception:
                    pass
    return metadata


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dir", nargs="+", required=True, help="carpeta(s) 'ofertas' (acepta comodines)")
    parser.add_argument("--excel", default="auditoria.xlsx")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--limite-licitaciones", type=int, default=0)
    args = parser.parse_args()

    tareas = cargar_tareas(args.dir, args.limite_licitaciones)
    if not tareas:
        raise SystemExit("No se encontraron extraccion_ia.json en --dir.")
    print(f"Auditando {len(tareas)} proveedores...")
    if args.workers > 1:
        with ProcessPoolExecutor(max_workers=args.workers) as pool:
            resultados = []
            for i, r in enumerate(pool.map(auditar_proveedor, tareas, chunksize=4), 1):
                resultados.append(r)
                if i % 100 == 0:
                    print(f"  {i}/{len(tareas)}", flush=True)
    else:
        resultados = [auditar_proveedor(t) for t in tareas]
    total = escribir_excel(resultados, args.excel, cargar_metadata(args.dir))
    productos = total["prod"] or 1
    print(f"\nProductos: {total['prod']} | verificados: {total['ver']} ({total['ver'] / productos:.0%}) | "
          f"a revisar: {total['rev']} | no verificables: {total['nov']} | rescatados de anexos: {total['res']}")
    print(f"Proveedores sin productos: {total['sinprod']} | con archivos sin leer: {total['sinleer']}")
    print(f"Excel: {args.excel}")


if __name__ == "__main__":
    main()

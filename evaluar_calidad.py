#!/usr/bin/env python3
"""Mide la calidad de una corrida de 3_extraer_ia.py contra valores verificados a mano ("gold").

El gold es un JSON {codigo_licitacion: {rut: [producto, ...]}} donde cada producto trae
tipo, marca, cantidad, precio_unitario y precio_total leidos del documento original. Un
proveedor con lista vacia significa "no ofrece equipos del alcance": cualquier producto
extraido para el es un falso positivo.

Un producto gold se da por encontrado si la corrida tiene uno con el mismo precio unitario
(+-1) y la misma cantidad. Sobre los encontrados se mide marca, tipo, cantidad y precios.

  python evaluar_calidad.py --dir pruebas_calidad/A/ofertas --gold pruebas_calidad/gold.json
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
import types
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
_SPEC = importlib.util.spec_from_file_location("extractor_ia", ROOT / "3_extraer_ia.py")
E = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(E)


def norm(texto):
    plano = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii").lower()
    return "".join(c for c in plano if c.isalnum())


def misma_marca(a, b):
    a, b = norm(E.normalizar_marca(a) or a), norm(E.normalizar_marca(b) or b)
    return bool(a) and bool(b) and (a == b or a in b or b in a)


def cercano(x, y, tol=1):
    return isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) <= tol


def tipo_de(producto):
    return E.subcategoria_producto(producto)


def emparejar(gold, predichos):
    usados = set()
    pares = []
    for g in gold:
        mejor = None
        for i, p in enumerate(predichos):
            if i in usados:
                continue
            if cercano(p.get("precio_unitario"), g["precio_unitario"]) and (
                    p.get("cantidad") in (None, g["cantidad"]) or cercano(p.get("cantidad"), g["cantidad"], 0)):
                mejor = i
                break
        if mejor is not None:
            usados.add(mejor)
        pares.append((g, predichos[mejor] if mejor is not None else None))
    sobrantes = [p for i, p in enumerate(predichos) if i not in usados]
    return pares, sobrantes


def cargar_corrida(directorio):
    corrida = {}
    for archivo in sorted(Path(directorio).glob("*/extraccion_ia.json")):
        datos = json.loads(archivo.read_text(encoding="utf-8"))
        corrida[archivo.parent.name] = {p["rut"]: p for p in datos}
    return corrida


ARGS_LECTURA = types.SimpleNamespace(
    max_paginas=200, ocr=False, vision=False, tesseract_cmd=None, idioma_ocr="spa+eng",
    dpi_ocr=200, vision_min_filas=2, max_filas_excel=5000, max_columnas_excel=40)


def aplicar_rescate(directorio, licitacion, registro):
    """Simula sobre un checkpoint el rescate de filas economicas (sin llamar al modelo)."""
    carpetas = [c for c in (Path(directorio) / licitacion).iterdir()
                if c.is_dir() and registro["rut"] in c.name]
    if not carpetas:
        return registro["productos"]
    carpeta = carpetas[0]
    bloques = E.bloques_economicos_completos(carpeta, E.listar_archivos_oferta(carpeta), [], ARGS_LECTURA)
    def de_historial(producto):
        for campo in ("archivo_fuente", "fuente_precio", "fuente_producto"):
            relativo = producto.get(campo)
            if not relativo or relativo == "consolidacion":
                continue
            ruta = carpeta / relativo
            texto = ""
            if ruta.is_file() and ruta.suffix.lower() in (".pdf", ".docx", ".xlsx", ".xlsm"):
                try:
                    texto, _ = E.extraer_archivo(ruta, ARGS_LECTURA)
                except Exception:
                    texto = ""
            if E.documento_de_historial(relativo, texto):
                return True
        return False

    # Simula que los documentos de historial (OC pasadas) ya no se envian al modelo
    registro = dict(registro)
    registro["productos"] = [p for p in registro["productos"] if not de_historial(p)]
    registro["productos_parciales"] = [p for p in registro.get("productos_parciales", []) if not de_historial(p)]
    parciales = registro["productos_parciales"]
    limpios = [{k: v for k, v in p.items() if k not in ("alertas", "estado_validacion")} for p in registro["productos"]]
    if registro.get("modo") != "por_proveedor":  # el modo por proveedor ya reconcilia al procesar
        limpios = E.reconciliar_precios_con_anexos(limpios, bloques, registro.get("total_oferta"), anular=False)
    resultado = E.rescatar_filas_economicas(limpios, parciales, bloques, registro["proveedor"], registro["rut"])
    resultado = E.completar_marcas_por_nombre_de_archivo(
        resultado, [str(a.relative_to(carpeta)) for a in E.listar_archivos_oferta(carpeta)])
    if registro.get("modo") != "por_proveedor":
        numeros = set()
        for ruta in E.listar_archivos_oferta(carpeta):
            try:
                texto, _ = E.extraer_archivo(ruta, ARGS_LECTURA)
            except Exception:
                continue
            numeros |= E.numeros_del_texto(texto)
        resultado = E.marcar_precios_sin_respaldo(resultado, numeros, bloques)
    return resultado


def evaluar(directorio, gold, verbose=True, rescate=False):
    corrida = cargar_corrida(directorio)
    total = {"gold": 0, "encontrados": 0, "completos": 0, "pred": 0, "tp": 0, "fp": 0,
             "marca_ok": 0, "tipo_ok": 0, "cant_ok": 0, "pt_ok": 0, "sin_resultado": 0,
             "ok_total": 0, "ok_buenos": 0}
    fallos = []
    segundos = 0.0
    for licitacion, proveedores in gold.items():
        for rut, items in proveedores.items():
            registro = (corrida.get(licitacion) or {}).get(rut)
            if registro is None:
                total["sin_resultado"] += 1
                total["gold"] += len(items)
                fallos.append((licitacion, rut, "PROVEEDOR SIN RESULTADO", None))
                continue
            segundos += registro.get("segundos") or 0
            predichos = aplicar_rescate(directorio, licitacion, registro) if rescate else list(registro["productos"])
            validados = E.validar_productos([{k: v for k, v in p.items() if k not in ("alertas", "estado_validacion")}
                                             for p in predichos], registro.get("productos_parciales", []),
                                            registro.get("total_oferta"))
            pares, sobrantes = emparejar(items, validados)
            ok_pred = [p for p in validados if p.get("estado_validacion") == "ok"]
            # Items ambiguos (p.ej. pantallas interactivas): ni suman ni restan.
            ambiguos_hallados = [p for g, p in pares if g.get("ambiguo") and p is not None]
            sobrantes = [p for p in sobrantes if p not in ambiguos_hallados]
            pares = [(g, p) for g, p in pares if not g.get("ambiguo")]
            total["gold"] += len(pares)
            total["pred"] += len(validados)
            total["fp"] += len(sobrantes)
            total["ok_total"] += len(ok_pred)
            total["ok_buenos"] += sum(1 for g, p in pares if p is not None and p.get("estado_validacion") == "ok"
                                      and cercano(p.get("precio_total"), g["precio_total"], max(2, g["precio_total"] * 0.001)))
            for g, p in pares:
                if p is None:
                    fallos.append((licitacion, rut, "FALTA", g))
                    continue
                total["encontrados"] += 1
                total["tp"] += 1
                m = True if g.get("marca") == "*" else (
                    misma_marca(p.get("marca"), g.get("marca")) if g.get("marca") else not p.get("marca"))
                t = g.get("tipo") == "*" or norm(tipo_de(p)) == norm(g.get("tipo"))
                c = cercano(p.get("cantidad"), g["cantidad"], 0)
                pt = cercano(p.get("precio_total"), g["precio_total"], max(2, g["precio_total"] * 0.001))
                total["marca_ok"] += m
                total["tipo_ok"] += t
                total["cant_ok"] += c
                total["pt_ok"] += pt
                if m and t and c and pt:
                    total["completos"] += 1
                else:
                    malos = [n for n, ok in (("marca", m), ("tipo", t), ("cantidad", c), ("total", pt)) if not ok]
                    fallos.append((licitacion, rut, "CAMPO " + ",".join(malos),
                                   {"gold": g, "pred": {k: p.get(k) for k in ("producto", "marca", "cantidad",
                                                                              "precio_unitario", "precio_total")}}))
            for p in sobrantes:
                fallos.append((licitacion, rut, "SOBRA", {k: p.get(k) for k in (
                    "producto", "marca", "cantidad", "precio_unitario", "precio_total", "estado_validacion")}))
    g, e = total["gold"] or 1, total["encontrados"] or 1
    ok_total = total["ok_total"] or 1
    recall = total["encontrados"] / g
    precision = total["tp"] / ((total["tp"] + total["fp"]) or 1)
    resumen = {
        "productos_gold": total["gold"], "recall": round(recall, 3), "precision": round(precision, 3),
        "marca_ok": round(total["marca_ok"] / e, 3), "tipo_ok": round(total["tipo_ok"] / e, 3),
        "cantidad_ok": round(total["cant_ok"] / e, 3), "total_ok": round(total["pt_ok"] / e, 3),
        "fiabilidad_completa": round(total["completos"] / g, 3),
        "falsos_positivos": total["fp"], "proveedores_sin_resultado": total["sin_resultado"],
        "ok_productos": total["ok_total"], "ok_precision": round(total["ok_buenos"] / ok_total, 3),
        "ok_recall": round(total["ok_buenos"] / g, 3),
        "segundos_proveedores": round(segundos),
    }
    if verbose:
        for lic, rut, tipo, detalle in fallos:
            print(f"  [{tipo}] {lic} {rut} {json.dumps(detalle, ensure_ascii=False, default=str)[:420] if detalle else ''}")
    return resumen, fallos


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dir", required=True)
    parser.add_argument("--gold", required=True)
    parser.add_argument("--silencioso", action="store_true")
    parser.add_argument("--rescate", action="store_true",
                        help="aplica el rescate de filas economicas sobre los checkpoints (sin modelo)")
    args = parser.parse_args()
    gold = {}
    for ruta in sorted(Path(args.gold).glob("*.json")) if Path(args.gold).is_dir() else [Path(args.gold)]:
        for licitacion, proveedores in json.loads(ruta.read_text(encoding="utf-8")).items():
            gold.setdefault(licitacion, {}).update(proveedores)
    resumen, _ = evaluar(args.dir, gold, verbose=not args.silencioso, rescate=args.rescate)
    print("\nRESUMEN:", json.dumps(resumen, ensure_ascii=False))


if __name__ == "__main__":
    main()

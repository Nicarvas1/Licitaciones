"""Mide cuantas llamadas por minuto sostiene LM Studio con distintos niveles de paralelismo.

Usa llamadas reales de extraccion (el mismo prompt del pipeline) sobre una oferta de ejemplo.
No modifica ningun archivo del proyecto.
"""
import concurrent.futures as cf
import importlib.util
import sys
import time
import types
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
spec = importlib.util.spec_from_file_location("e", RAIZ / "3_extraer_ia.py")
e = importlib.util.module_from_spec(spec)
spec.loader.exec_module(e)

MODELO = sys.argv[1] if len(sys.argv) > 1 else "qwen/qwen3.6-35b-a3b"
NIVELES = [int(x) for x in (sys.argv[2].split(",") if len(sys.argv) > 2 else ["1", "4", "6", "8"])]
ARGS = types.SimpleNamespace(max_paginas=60, ocr=False, vision=False, tesseract_cmd=None, idioma_ocr="spa+eng",
                             dpi_ocr=200, vision_min_filas=2, max_filas_excel=400, max_columnas_excel=25)

base = RAIZ / "pruebas_calidad" / "B" / "ofertas" / "3604-33-LE25"
prompts = []
for carpeta in sorted(c for c in base.iterdir() if c.is_dir()):
    for archivo in e.listar_archivos_oferta(carpeta):
        if e.tipo_documental(archivo.name) != "economico" or archivo.suffix.lower() != ".pdf":
            continue
        texto, estado = e.extraer_archivo(archivo, ARGS)
        if not estado.startswith("texto") or len(texto) < 400:
            continue
        prompts.append(e.PROMPT_ARCHIVO.format(
            proveedor=carpeta.name, rut="", archivo=archivo.name, tipo_documental="economico",
            total_oferta="", pistas_producto="(ninguna)", pistas_precio="(ninguna)",
            texto=e.compactar_texto(texto)[:9000]))
prompts = prompts[:16]
print(f"{len(prompts)} prompts reales de extraccion; modelo {MODELO}")


def una(prompt):
    inicio = time.time()
    datos, estado, cruda, consumo = e.consultar_modelo(prompt, MODELO, 600, 8192, "lmstudio", 2048, 1, 2)
    return time.time() - inicio, estado, (consumo or {}).get("completion_tokens") or 0


for nivel in NIVELES:
    carga = (prompts * 3)[: max(nivel * 2, 8)]
    inicio = time.time()
    with cf.ThreadPoolExecutor(nivel) as pool:
        resultados = list(pool.map(una, carga))
    total = time.time() - inicio
    tokens = sum(r[2] for r in resultados)
    errores = sum(1 for r in resultados if r[1] not in e.ESTADOS_IA_OK)
    print(f"paralelo {nivel}: {len(carga)} llamadas en {total:.0f}s -> {len(carga) / total * 60:.1f} llamadas/min, "
          f"{tokens / total:.0f} tokens/s, p50 {sorted(r[0] for r in resultados)[len(resultados) // 2]:.0f}s, errores {errores}")

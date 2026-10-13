#!/usr/bin/env python3
"""Copia licitaciones descargadas a una carpeta de prueba limpia (sin resultados previos).

Una prueba de calidad no debe tocar los lotes reales ni arrastrar `extraccion_ia.json` anteriores.
Este script copia solo los anexos descargados de las licitaciones indicadas y deja listo el CSV de
metadatos (nombre y fecha de cada licitacion) para el Excel final.

Ejemplo (las 10 licitaciones del gold de desarrollo, de noviembre):
  python calidad/preparar_copia.py --origen lotes_z8/2025-11 --destino pruebas_calidad/A \
      --licitaciones 5124-14-LE25 3604-33-LE25 4291-125-LE25 3289-13-LE25 3937-18-L125 \
                     1260-50-LE25 3650-20-LE25 564162-131-L125 2591-160-LP25 1057062-13-LE25

Luego:
  bash calidad/correr_prueba.sh A --rapido
  python evaluar_calidad.py --dir pruebas_calidad/A/ofertas --gold calidad/gold_desarrollo
"""

import argparse
import csv
import shutil
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--origen", required=True, help="carpeta del mes, p.ej. lotes_z8/2025-11 (contiene ofertas/)")
    parser.add_argument("--destino", required=True, help="carpeta de prueba, p.ej. pruebas_calidad/A")
    parser.add_argument("--licitaciones", nargs="+", required=True, help="codigos de licitacion a copiar")
    args = parser.parse_args()

    origen, destino = Path(args.origen), Path(args.destino)
    ofertas = origen / "ofertas"
    if not ofertas.is_dir():
        sys.exit(f"No existe {ofertas}")
    ignorar = shutil.ignore_patterns("extraccion_*.json", "_extraidos", "_debug_ia")
    copiadas = []
    for codigo in args.licitaciones:
        fuente = ofertas / codigo
        if not fuente.is_dir():
            print(f"  (no existe {codigo})")
            continue
        final = destino / "ofertas" / codigo
        if not final.exists():
            shutil.copytree(fuente, final, ignore=ignorar)
        copiadas.append(codigo)

    csv_origen = origen / "para_scrapear.csv"
    if csv_origen.is_file():
        for encoding in ("utf-8-sig", "latin-1"):
            try:
                with open(csv_origen, encoding=encoding, newline="") as archivo:
                    filas = [f for f in csv.DictReader(archivo) if f.get("codigo") in copiadas]
                    campos = list(next(csv.DictReader(open(csv_origen, encoding=encoding, newline=""))).keys())
                break
            except UnicodeDecodeError:
                continue
        with open(destino / "para_scrapear.csv", "w", encoding="utf-8-sig", newline="") as archivo:
            escritor = csv.DictWriter(archivo, fieldnames=campos)
            escritor.writeheader()
            escritor.writerows(filas)
    print(f"{len(copiadas)} licitaciones copiadas a {destino}")


if __name__ == "__main__":
    main()

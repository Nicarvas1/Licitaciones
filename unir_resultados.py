#!/usr/bin/env python3
"""Une los Excel mensuales (resultado_YYYY-MM.xlsx) en uno solo, con una columna "Mes".

  python unir_resultados.py --dir D:\\licitaciones_datos --excel D:\\licitaciones_datos\\resultado_anio.xlsx
"""

import argparse
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

PATRON = re.compile(r"^resultado_(\d{4}-\d{2})\.xlsx$")


def main():
    parser = argparse.ArgumentParser(description="Une los Excel mensuales en uno solo")
    parser.add_argument("--dir", required=True, help="carpeta con los meses (se busca en 3 niveles)")
    parser.add_argument("--excel", required=True)
    args = parser.parse_args()

    archivos = sorted(
        (ruta for ruta in Path(args.dir).rglob("resultado_*.xlsx") if PATRON.match(ruta.name)),
        key=lambda ruta: ruta.name,
    )
    if not archivos:
        sys.exit(f"No encontre resultado_YYYY-MM.xlsx en {args.dir}")

    columnas, filas = None, []
    for ruta in archivos:
        mes = PATRON.match(ruta.name).group(1)
        libro = load_workbook(ruta, read_only=True, data_only=True)
        if "Productos" not in libro.sheetnames:
            libro.close()
            continue
        encabezado, *registros = list(libro["Productos"].iter_rows(values_only=True))
        libro.close()
        if columnas is None:
            columnas = list(encabezado)
        indice = {nombre: i for i, nombre in enumerate(encabezado)}
        for registro in registros:
            if not any(valor not in (None, "") for valor in registro):
                continue
            filas.append([mes] + [registro[indice[c]] if c in indice else None for c in columnas])
        print(f"  {mes}: {len(registros)} productos")

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Productos"
    hoja.append(["Mes"] + columnas)
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    for fila in filas:
        hoja.append(fila)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    libro.save(args.excel)
    print(f"Total: {len(filas)} productos de {len(archivos)} meses -> {args.excel}")


if __name__ == "__main__":
    main()

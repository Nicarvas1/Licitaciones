#!/usr/bin/env python3
"""Cuenta lo descargado y compara contra lo que el portal ofrecia (oferta.json de cada proveedor).

Para cada proveedor, `oferta.json` guarda los anexos que el portal mostraba (enlaces_anexos). Aqui se verifica
que cada tipo de anexo con enlace tenga archivos en disco. Una carpeta vacia solo es legitima si el portal no
ofrecia anexos o solo ofrecia administrativos (que se bajan unicamente cuando no hay economico ni tecnico).

  python contar_descargas.py --dir D:\\lotes_prueba\\2025-11\\ofertas
  python contar_descargas.py --dir D:\\lotes_prueba\\2025-11\\ofertas --csv faltantes.csv
"""

import argparse
import csv
import json
import re
import sys
from collections import Counter
from pathlib import Path

PREFIJO = re.compile(r"^(economico|tecnico|administrativo)__\d{2}__")
IGNORAR = {".json", ".csv", ".html"}


def archivos_de(carpeta):
    return [f for f in carpeta.rglob("*") if f.is_file() and f.suffix.lower() not in IGNORAR]


def main():
    parser = argparse.ArgumentParser(description="Cuenta descargas y detecta anexos faltantes")
    parser.add_argument("--dir", required=True, help="carpeta ofertas (una licitacion por subcarpeta)")
    parser.add_argument("--csv", help="escribe aqui los proveedores con anexos faltantes")
    args = parser.parse_args()

    raiz = Path(args.dir)
    if not raiz.is_dir():
        sys.exit(f"No existe {raiz}")
    c, faltantes, vacias_legitimas = Counter(), [], Counter()
    licitaciones = [d for d in sorted(raiz.iterdir()) if d.is_dir()]
    for lic in licitaciones:
        if not (lic / "metadata.json").is_file():
            c["licitaciones_sin_metadata"] += 1
        for prov in sorted(p for p in lic.iterdir() if p.is_dir() and not p.name.startswith("_")):
            c["proveedores"] += 1
            try:
                enlaces = (json.loads((prov / "oferta.json").read_text(encoding="utf-8")).get("enlaces_anexos") or {})
            except (OSError, json.JSONDecodeError):
                enlaces = {}
                c["sin_oferta_json"] += 1
            archivos = archivos_de(prov)
            por_tipo = Counter((PREFIJO.match(f.name).group(1) if PREFIJO.match(f.name) else "otro") for f in archivos)
            for tipo, n in por_tipo.items():
                c[f"archivos_{tipo}"] += n
            c["archivos"] += len(archivos)
            if archivos:
                c["proveedores_con_archivos"] += 1
            for tipo in ("economico", "tecnico"):
                if tipo in enlaces and not por_tipo[tipo]:
                    faltantes.append((lic.name, prov.name, tipo, "enlace sin archivos"))
            if not archivos and "administrativo" in enlaces and not (set(enlaces) - {"administrativo"}):
                faltantes.append((lic.name, prov.name, "administrativo", "solo administrativo sin bajar"))
            if not archivos and not enlaces:
                vacias_legitimas["sin_anexos_en_portal"] += 1
            # un anexo con exactamente 6 archivos (o multiplo) puede estar truncado por la paginacion del popup
            for tipo in ("economico", "tecnico", "administrativo"):
                if por_tipo[tipo] and por_tipo[tipo] % 6 == 0:
                    c[f"anexos_{tipo}_con_6_o_multiplo"] += 1

    print(f"Carpeta: {raiz}")
    print(f"Licitaciones:                {len(licitaciones)}  (sin metadata.json: {c['licitaciones_sin_metadata']})")
    print(f"Proveedores (carpetas):      {c['proveedores']}")
    print(f"  con al menos un archivo:   {c['proveedores_con_archivos']}")
    print(f"  vacios legitimos (portal sin anexos): {vacias_legitimas['sin_anexos_en_portal']}")
    print(f"Archivos descargados:        {c['archivos']}")
    for tipo in ("economico", "tecnico", "administrativo", "otro"):
        print(f"  {tipo:<16}{c['archivos_' + tipo]}")
    print(f"Anexos con 6 archivos (o multiplo; posible truncado): "
          f"eco={c['anexos_economico_con_6_o_multiplo']} tec={c['anexos_tecnico_con_6_o_multiplo']} "
          f"adm={c['anexos_administrativo_con_6_o_multiplo']}")
    print(f"\nProveedores con anexos que el portal ofrecia y NO hay en disco: {len(faltantes)}")
    for fila in faltantes[:15]:
        print("  ", *fila)
    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["licitacion", "proveedor", "tipo", "motivo"])
            w.writerows(faltantes)
        print(f"CSV: {args.csv}")


if __name__ == "__main__":
    main()

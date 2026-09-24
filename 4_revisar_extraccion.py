#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera una pagina HTML de revision con la evidencia de cada producto extraido.

Por cada producto muestra:
  - PDF: un recorte de la pagina con la fila marcada en rojo y las celdas de
    cantidad (azul), precio unitario (verde) y total (morado) resaltadas.
  - Excel / Word / CSV: el fragmento de la planilla o tabla con la fila destacada.
  - Un enlace para abrir el documento original (en la pagina correcta si es PDF).
  - Botones Correcto / Incorrecto / Dudoso y comentario. Las decisiones se guardan
    en el navegador y se exportan a CSV.

Verificacion automatica: antes de dar una ubicacion por "exacta" se comprueba que
el precio unitario aparece en el texto bajo el rectangulo. Si no, se busca el
precio en el documento; si tampoco aparece, el producto se marca como
"precio no encontrado en el documento".

Funciona con los resultados por reglas (extraccion_reglas.json, con coordenadas
exactas) y con los de la IA (extraccion_ia.json): en ambos casos la fila se ubica
buscando los montos en el documento y se verifica palabra por palabra.

Uso:
  python 4_revisar_extraccion.py --dir ".\\lotes\\2025-11\\ofertas"
  python 4_revisar_extraccion.py --dir ".\\lotes\\2025-11\\ofertas" --fuente ia
Luego abrir revision_reglas\\index.html (o revision_ia\\index.html) en Chrome o Edge.
"""

import argparse
import csv
import html
import importlib.util
import json
import os
import re
import sys
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ_PROYECTO))
_spec = importlib.util.spec_from_file_location("extractor_ia", RAIZ_PROYECTO / "3_extraer_ia.py")
ia = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ia)
fitz = ia.fitz

COLOR_FILA = (0.86, 0.10, 0.10)
COLOR_CELDA = {"cantidad": (0.10, 0.35, 0.90), "unitario": (0.05, 0.65, 0.20), "total": (0.55, 0.20, 0.75)}
ARCHIVOS_JSON = {"reglas": "extraccion_reglas.json", "ia": "extraccion_ia.json"}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def variantes_monto(valor):
    """Formas en que un monto puede aparecer escrito: 650.000 / 650,000 / 650000."""
    if not isinstance(valor, (int, float)) or valor <= 0:
        return []
    variantes = []
    if float(valor).is_integer():
        entero = int(valor)
        variantes += [f"{entero:,}".replace(",", "."), f"{entero:,}", str(entero)]
    else:
        variantes += [f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), f"{valor:,.2f}", f"{valor:.2f}"]
    return [v for i, v in enumerate(variantes) if v not in variantes[:i]]


def digitos(valor):
    if not isinstance(valor, (int, float)):
        return ""
    return str(int(round(valor))) if float(valor).is_integer() else re.sub(r"\D", "", f"{valor:.2f}")


def texto_contiene_monto(texto, valor):
    objetivo = digitos(valor)
    return bool(objetivo) and objetivo in re.sub(r"\D", "", texto or "")


def letra_columna(indice):
    letras = ""
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def celda_texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    if isinstance(valor, int):
        return f"{valor:,}".replace(",", ".")
    return " ".join(str(valor).split())


def enlace_relativo(destino, desde):
    return Path(os.path.relpath(destino, desde)).as_posix()


# ---------------------------------------------------------------------------
# Cache de documentos
# ---------------------------------------------------------------------------
class Documentos:
    def __init__(self):
        self.pdfs = {}
        self.hojas = {}
        self.docx = {}

    def pdf(self, ruta):
        if ruta not in self.pdfs:
            self.pdfs[ruta] = fitz.open(ruta)
        return self.pdfs[ruta]

    def filas_excel(self, ruta, hoja=None):
        clave = (ruta, hoja)
        if clave not in self.hojas:
            libro = ia.load_workbook(ruta, read_only=True, data_only=True)
            try:
                hoja_obj = libro[hoja] if hoja and hoja in libro.sheetnames else libro.worksheets[0]
                self.hojas[clave] = [list(fila) for fila in hoja_obj.iter_rows(max_row=500, max_col=40, values_only=True)]
            finally:
                libro.close()
        return self.hojas[clave]

    def tablas_docx(self, ruta):
        if ruta not in self.docx:
            documento = ia.Document(ruta)
            tablas = [[[celda.text for celda in fila.cells] for fila in tabla.rows] for tabla in documento.tables]
            parrafos = [[p.text] for p in documento.paragraphs if p.text.strip()]
            self.docx[ruta] = (tablas, parrafos)
        return self.docx[ruta]

    def filas_csv(self, ruta):
        clave = (ruta, "csv")
        if clave not in self.hojas:
            contenido = ""
            for encoding in ("utf-8-sig", "latin-1"):
                try:
                    contenido = Path(ruta).read_text(encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            try:
                dialecto = csv.Sniffer().sniff(contenido[:4096], delimiters=",;\t")
                self.hojas[clave] = list(csv.reader(contenido.splitlines(), dialecto))
            except csv.Error:
                self.hojas[clave] = [[linea] for linea in contenido.splitlines()]
        return self.hojas[clave]

    def cerrar(self):
        for documento in self.pdfs.values():
            documento.close()


# ---------------------------------------------------------------------------
# Evidencia en PDF
# ---------------------------------------------------------------------------
def palabras_visuales(pagina):
    """Palabras de la pagina: (rect visual con la rotacion aplicada, rect sin rotar, texto).
    Las filas se razonan en coordenadas visuales; el dibujo usa las sin rotar."""
    resultado = []
    for x0, y0, x1, y1, texto, *_ in pagina.get_text("words"):
        sin_rotar = fitz.Rect(x0, y0, x1, y1)
        resultado.append((sin_rotar * pagina.rotation_matrix, sin_rotar, texto))
    return resultado


def palabra_es_monto(texto, valor):
    objetivo = digitos(valor)
    if not objetivo:
        return False
    d = re.sub(r"\D", "", texto)
    # admite decimales en cero: 650.000,00 -> 65000000
    return d == objetivo or (d.startswith(objetivo) and len(d) - len(objetivo) <= 2 and set(d[len(objetivo):]) <= {"0"})


def palabra_es_entero(texto, valor):
    if not isinstance(valor, (int, float)) or not float(valor).is_integer():
        return False
    return texto.strip(" .,;:()$-") == str(int(valor))


PALABRAS_VACIAS = {"de", "del", "la", "el", "los", "las", "con", "para", "por", "en", "una", "gb", "ram", "ssd"}


def tokens_descripcion(producto):
    texto = ia.texto_normalizado(" ".join(str(producto.get(c) or "") for c in ("producto", "modelo")))
    return {t for t in texto.split() if len(t) >= 3 and t not in PALABRAS_VACIAS}


def evaluar_linea(palabras, banda, producto):
    """Que datos del producto aparecen en la linea (banda horizontal en coordenadas visuales)."""
    en_linea = [(v, r, t) for v, r, t in palabras if banda.y0 <= (v.y0 + v.y1) / 2 <= banda.y1]
    encontrado = {"unitario": None, "total": None, "cantidad": None}
    for _, r, t in en_linea:
        if encontrado["unitario"] is None and palabra_es_monto(t, producto.get("precio_unitario")):
            encontrado["unitario"] = r
        elif encontrado["total"] is None and palabra_es_monto(t, producto.get("precio_total")):
            encontrado["total"] = r
        elif encontrado["cantidad"] is None and palabra_es_entero(t, producto.get("cantidad")):
            encontrado["cantidad"] = r
    texto_linea = set(ia.texto_normalizado(" ".join(t for _, _, t in en_linea)).split())
    coincidencias = len(tokens_descripcion(producto) & texto_linea)
    return encontrado, coincidencias, en_linea


def dibujar_y_recortar(documento, indice_pagina, marcas, recorte_visual, ruta_imagen, dpi):
    """marcas: [(rect_sin_rotar, color, relleno)]. recorte_visual: zona a renderizar en
    coordenadas visuales (get_pixmap trabaja sobre la pagina ya rotada)."""
    temporal = fitz.open()
    temporal.insert_pdf(documento, from_page=indice_pagina, to_page=indice_pagina)
    pagina = temporal[0]
    for rect, color, relleno in marcas:
        pagina.draw_rect(rect, color=color, fill=color if relleno else None,
                         fill_opacity=0.25 if relleno else 0, width=0.8 if relleno else 1.6)
    if recorte_visual.height < 120:
        recorte_visual = fitz.Rect(recorte_visual.x0, recorte_visual.y0, recorte_visual.x1,
                                   min(pagina.rect.y1, recorte_visual.y0 + 120))
    pixmap = pagina.get_pixmap(dpi=dpi, clip=recorte_visual)
    try:
        ruta_final = ruta_imagen.parent / f"{ruta_imagen.name}.jpg"
        pixmap.save(str(ruta_final), jpg_quality=82)
    except Exception:
        ruta_final = ruta_imagen.parent / f"{ruta_imagen.name}.png"
        pixmap.save(str(ruta_final))
    temporal.close()
    return ruta_final


def evidencia_pdf(documentos, ruta, producto, ruta_imagen, dpi):
    """Ubica la fila del producto buscando sus montos con PyMuPDF (el mismo sistema de
    coordenadas que se usa para dibujar) y la verifica palabra por palabra.
    Si un monto aparece varias veces, elige la linea que ademas contiene el total,
    la cantidad y palabras de la descripcion."""
    documento = documentos.pdf(ruta)
    ubicacion = producto.get("ubicacion") or {}
    pagina_num = ubicacion.get("pagina") or producto.get("pagina")
    try:
        pagina_num = int(pagina_num) if pagina_num else None
    except (TypeError, ValueError):
        pagina_num = None
    paginas = list(range(min(len(documento), 30)))
    if pagina_num and 1 <= pagina_num <= len(documento):
        paginas.remove(pagina_num - 1)
        paginas.insert(0, pagina_num - 1)

    campos_ancla = [c for c in ("precio_unitario", "precio_total") if isinstance(producto.get(c), (int, float))]
    mejor = None
    for indice in paginas:
        pagina = documento[indice]
        palabras = palabras_visuales(pagina)
        for campo in campos_ancla:
            for visual, _, texto in palabras:
                if not palabra_es_monto(texto, producto.get(campo)):
                    continue
                banda = fitz.Rect(pagina.rect.x0, visual.y0 - 1.5, pagina.rect.x1, visual.y1 + 1.5)
                encontrado, coincidencias, en_linea = evaluar_linea(palabras, banda, producto)
                puntaje = (3 * bool(encontrado["unitario"]) + 2 * bool(encontrado["total"])
                           + 2 * bool(encontrado["cantidad"]) + min(coincidencias, 3)
                           + (1 if indice == (pagina_num or 0) - 1 else 0))
                if mejor is None or puntaje > mejor[0]:
                    mejor = (puntaje, indice, banda, encontrado, coincidencias, en_linea)
            if mejor and mejor[3]["unitario"]:
                break  # anclado por el unitario; no hace falta probar con el total
        if mejor and mejor[0] >= 8:
            break  # linea con unitario, total, cantidad y descripcion en la pagina esperada

    if mejor is None:
        return {"pagina": pagina_num, "ubicado": "no_encontrado",
                "nota": "el precio no aparece como texto en el PDF (escaneado o monto inexistente)"}

    _, indice, banda, encontrado, coincidencias, en_linea = mejor
    pagina = documento[indice]
    faltan = [nombre for campo, clave, nombre in (("precio_unitario", "unitario", "precio unitario"),
                                                 ("precio_total", "total", "total"),
                                                 ("cantidad", "cantidad", "cantidad"))
              if isinstance(producto.get(campo), (int, float)) and not encontrado[clave]]

    visuales = [v for v, _, _ in en_linea]
    fila_visual = fitz.Rect(min(v.x0 for v in visuales) - 4, banda.y0 - 1, max(v.x1 for v in visuales) + 4, banda.y1 + 1)
    marcas = [(r, COLOR_CELDA[tipo], True) for tipo, r in encontrado.items() if r is not None]
    marcas.append((fila_visual * pagina.derotation_matrix, COLOR_FILA, False))

    # Recorte visual: se incluye el encabezado de la tabla si esta cerca, sobre la fila
    arriba = banda.y0 - 45
    for v, _, t in palabras_visuales(pagina):
        if v.y1 < banda.y0 and banda.y0 - v.y1 < 300 and re.search(r"(?i)unitari|precio|cantidad|descrip", t):
            arriba = min(arriba, v.y0 - 8)
    recorte_visual = fitz.Rect(pagina.rect.x0, max(pagina.rect.y0, arriba),
                               pagina.rect.x1, min(pagina.rect.y1, banda.y1 + 30))
    imagen = dibujar_y_recortar(documento, indice, marcas, recorte_visual, ruta_imagen, dpi)

    notas = []
    if faltan:
        estado = "parcial"
        notas.append("la linea marcada NO contiene: " + ", ".join(faltan) + " (revisar si esos datos vienen de otra fila)")
    else:
        estado = "exacto"
        if ubicacion.get("continuacion"):
            notas.append("la tabla continua de la pagina anterior (el encabezado esta alli)")
    if coincidencias == 0 and tokens_descripcion(producto):
        notas.append("la descripcion no esta en la misma linea (puede ocupar varias lineas)")
    return {"imagen": imagen, "pagina": indice + 1, "ubicado": estado, "nota": "; ".join(notas)}


# ---------------------------------------------------------------------------
# Evidencia en planillas y tablas
# ---------------------------------------------------------------------------
def celda_con_valor(fila, valor):
    for celda in fila:
        if isinstance(celda, (int, float)) and abs(celda - valor) <= 0.5:
            return True
        if isinstance(celda, str) and any(palabra_es_monto(t, valor) or palabra_es_entero(t, valor)
                                          for t in celda.split()):
            return True
    return False


def fila_con_monto(filas, producto):
    for campo in ("precio_unitario", "precio_total"):
        objetivo = producto.get(campo)
        if not isinstance(objetivo, (int, float)) or objetivo <= 0:
            continue
        for indice, fila in enumerate(filas):
            for celda in fila:
                if isinstance(celda, (int, float)) and abs(celda - objetivo) <= 0.5:
                    return indice
                if isinstance(celda, str) and texto_contiene_monto(celda, objetivo) and len(re.sub(r"\D", "", celda)) <= len(digitos(objetivo)) + 2:
                    return indice
    return None


def tabla_html(filas, fila_objetivo, fila_encabezado, columnas, titulo):
    if fila_objetivo is None or fila_objetivo >= len(filas):
        return None
    mostrar = []
    if fila_encabezado is not None and 0 <= fila_encabezado < fila_objetivo:
        if fila_encabezado > 0:
            mostrar.append(fila_encabezado - 1)
        mostrar.append(fila_encabezado)
    inicio = max(fila_objetivo - 2, (fila_encabezado + 1) if fila_encabezado is not None else 0)
    mostrar += list(range(inicio, min(len(filas), fila_objetivo + 3)))
    mostrar = sorted(set(mostrar))
    ancho = max((len(filas[i]) for i in mostrar), default=0)
    usadas = [c for c in range(ancho) if any(c < len(filas[i]) and celda_texto(filas[i][c]) for i in mostrar)][:14]
    clase_celda = {indice: tipo for tipo, indice in (columnas or {}).items() if tipo in ("cantidad", "unitario", "total")}
    partes = [f'<div class="tabla-titulo">{html.escape(titulo)}</div><table class="grilla"><tr><th></th>']
    partes += [f"<th>{letra_columna(c)}</th>" for c in usadas]
    partes.append("</tr>")
    anterior = None
    for i in mostrar:
        if anterior is not None and i > anterior + 1:
            partes.append(f'<tr class="salto"><td colspan="{len(usadas) + 1}">…</td></tr>')
        clase = "objetivo" if i == fila_objetivo else ("encabezado" if i == fila_encabezado else "")
        partes.append(f'<tr class="{clase}"><td class="num">{i + 1}</td>')
        for c in usadas:
            valor = celda_texto(filas[i][c]) if c < len(filas[i]) else ""
            clase_c = f"c-{clase_celda[c]}" if i == fila_objetivo and c in clase_celda else ""
            partes.append(f'<td class="{clase_c}">{html.escape(valor[:120])}</td>')
        partes.append("</tr>")
        anterior = i
    partes.append("</table>")
    return "".join(partes)


def evidencia_tabular(documentos, ruta, producto):
    ubicacion = producto.get("ubicacion") or {}
    extension = Path(ruta).suffix.lower()
    if extension in (".xlsx", ".xlsm"):
        filas = documentos.filas_excel(ruta, ubicacion.get("hoja"))
        titulo = f"Hoja {ubicacion.get('hoja') or '(primera)'}"
    elif extension == ".docx":
        tablas, parrafos = documentos.tablas_docx(ruta)
        if ubicacion.get("tipo_fuente") == "docx_texto":
            filas, titulo = parrafos, "Texto del documento"
        elif ubicacion.get("tabla") and ubicacion["tabla"] <= len(tablas):
            filas, titulo = tablas[ubicacion["tabla"] - 1], f"Tabla {ubicacion['tabla']}"
        else:
            filas, titulo = None, ""
            for numero, tabla in enumerate(tablas, 1):
                if fila_con_monto(tabla, producto) is not None:
                    filas, titulo = tabla, f"Tabla {numero}"
                    break
            if filas is None:
                filas, titulo = parrafos, "Texto del documento"
    else:
        filas, titulo = documentos.filas_csv(ruta), "CSV"

    fila = ubicacion.get("fila")
    if fila is None or fila >= len(filas) or fila_con_monto([filas[fila]], producto) is None:
        fila = fila_con_monto(filas, producto)
        encabezado, columnas = None, {}
    else:
        encabezado, columnas = ubicacion.get("fila_encabezado"), ubicacion.get("columnas")
    tabla = tabla_html(filas, fila, encabezado, columnas, titulo)
    if tabla is None:
        return {"ubicado": "no_encontrado", "nota": "el precio no aparece en la planilla/tabla"}
    faltan = [nombre for campo, nombre in (("precio_unitario", "precio unitario"), ("precio_total", "total"),
                                           ("cantidad", "cantidad"))
              if isinstance(producto.get(campo), (int, float)) and not celda_con_valor(filas[fila], producto[campo])]
    if faltan:
        return {"html": tabla, "ubicado": "parcial",
                "nota": "la fila marcada NO contiene: " + ", ".join(faltan) + " (revisar si esos datos vienen de otra fila)"}
    return {"html": tabla, "ubicado": "exacto", "nota": ""}


# ---------------------------------------------------------------------------
# Recoleccion
# ---------------------------------------------------------------------------
def carpetas_por_rut(licitacion):
    mapa = {}
    for carpeta in licitacion.iterdir():
        if carpeta.is_dir() and (carpeta / "oferta.json").is_file():
            rut = str(ia.cargar_json(carpeta / "oferta.json").get("rut") or "").strip()
            mapa[rut] = carpeta
    return mapa


def archivos_candidatos(producto, carpeta):
    nombres = []
    for campo in ("archivo_fuente", "fuente_precio", "fuente_producto"):
        valor = producto.get(campo)
        if valor and valor != "consolidacion":
            nombres.append(valor)
    nombres += [f for f in producto.get("fuentes_respaldo") or [] if f != "consolidacion"]
    rutas = []
    for nombre in nombres:
        ruta = carpeta / nombre
        if ruta.is_file() and ruta not in rutas:
            rutas.append(ruta)
    return rutas


def precios_otro_metodo(licitacion, fuente):
    otro = "ia" if fuente == "reglas" else "reglas"
    ruta = licitacion / ARCHIVOS_JSON[otro]
    if not ruta.is_file():
        return None
    try:
        resultados = json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return {
        str(r.get("rut") or "").strip(): [p.get("precio_unitario") for p in r.get("productos", [])
                                         if isinstance(p.get("precio_unitario"), (int, float))]
        for r in resultados
    }


def comparar_con_otro(producto, precios_otro):
    if precios_otro is None:
        return "sin_datos"
    if not precios_otro:
        return "otro_sin_precios"
    unitario = producto.get("precio_unitario")
    if not isinstance(unitario, (int, float)):
        return "sin_unitario"
    return "mismo_precio" if any(ia.numeros_iguales(unitario, p) for p in precios_otro) else "precio_distinto"


def main():
    parser = argparse.ArgumentParser(description="Pagina de revision con evidencia por producto.")
    parser.add_argument("--dir", required=True, help="Carpeta de ofertas del lote (ej. lotes/2025-11/ofertas)")
    parser.add_argument("--fuente", choices=("reglas", "ia"), default="reglas")
    parser.add_argument("--salida", help="Carpeta de salida (por defecto <lote>/revision_<fuente>)")
    parser.add_argument("--dpi", type=int, default=110)
    parser.add_argument("--limite-licitaciones", type=int)
    parser.add_argument("--muestra", type=int, default=50, help="Tamano de la muestra aleatoria sugerida")
    args = parser.parse_args()

    raiz = Path(args.dir).resolve()
    if not raiz.is_dir():
        sys.exit(f"No existe: {raiz}")
    salida = Path(args.salida).resolve() if args.salida else raiz.parent / f"revision_{args.fuente}"
    carpeta_img = salida / "img"
    carpeta_img.mkdir(parents=True, exist_ok=True)
    licitaciones = ia.detectar_licitaciones(raiz)
    if args.limite_licitaciones:
        licitaciones = licitaciones[:args.limite_licitaciones]

    documentos = Documentos()
    registros = []
    conteo = {"exacto": 0, "parcial": 0, "no_encontrado": 0, "sin_archivo": 0}
    for licitacion in licitaciones:
        ruta_json = licitacion / ARCHIVOS_JSON[args.fuente]
        if not ruta_json.is_file():
            continue
        try:
            resultados = json.loads(ruta_json.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            print(f"  {licitacion.name}: no se pudo leer {ruta_json.name}")
            continue
        carpetas = carpetas_por_rut(licitacion)
        otros = precios_otro_metodo(licitacion, args.fuente)
        for resultado in resultados:
            rut = str(resultado.get("rut") or "").strip()
            carpeta = carpetas.get(rut)
            for numero, producto in enumerate(resultado.get("productos", []), 1):
                identificador = ia.id_revision(licitacion.name, rut, producto)
                registro = {
                    "id": identificador,
                    "licitacion": licitacion.name,
                    "proveedor": resultado.get("proveedor", ""),
                    "rut": rut,
                    "producto": producto.get("producto"),
                    "marca": producto.get("marca"),
                    "modelo": producto.get("modelo"),
                    "categoria": producto.get("categoria"),
                    "tipo": producto.get("subcategoria") or ia.subcategoria_producto(dict(producto)),
                    "cantidad": producto.get("cantidad"),
                    "precio_unitario": producto.get("precio_unitario"),
                    "precio_total": producto.get("precio_total"),
                    "moneda": producto.get("moneda"),
                    "estado_validacion": producto.get("estado_validacion"),
                    "alertas": producto.get("alertas") or [],
                    "metodo": producto.get("metodo") or ("ia" if args.fuente == "ia" else ""),
                    "evidencia_texto": producto.get("evidencia"),
                    "otro_metodo": comparar_con_otro(
                        producto, None if otros is None else otros.get(rut, [])),
                }
                rutas = archivos_candidatos(producto, carpeta) if carpeta else []
                evidencia = {"ubicado": "sin_archivo", "nota": "no se encontro el archivo fuente"}
                for ruta in rutas:
                    try:
                        if ruta.suffix.lower() == ".pdf":
                            nombre_imagen = re.sub(r"[^\w-]", "_", identificador)  # sin puntos (RUT 76.596.570-5)
                            evidencia = evidencia_pdf(documentos, ruta, producto, carpeta_img / nombre_imagen, args.dpi)
                        elif ruta.suffix.lower() in (".xlsx", ".xlsm", ".docx", ".csv", ".tsv", ".txt"):
                            evidencia = evidencia_tabular(documentos, ruta, producto)
                        elif ruta.suffix.lower() in ia.EXTENSIONES_IMAGEN:
                            evidencia = {"imagen_original": ruta, "ubicado": "no_encontrado",
                                         "nota": "imagen: revisar a mano"}
                        else:
                            continue
                    except Exception as exc:
                        evidencia = {"ubicado": "no_encontrado", "nota": f"error al generar evidencia: {exc}"}
                    evidencia["archivo"] = ruta
                    if evidencia["ubicado"] in ("exacto", "parcial"):
                        break
                conteo[evidencia["ubicado"]] = conteo.get(evidencia["ubicado"], 0) + 1
                if evidencia.get("archivo"):
                    enlace = enlace_relativo(evidencia["archivo"], salida)
                    if Path(evidencia["archivo"]).suffix.lower() == ".pdf" and evidencia.get("pagina"):
                        enlace += f"#page={evidencia['pagina']}"
                    registro["abrir"] = enlace
                    registro["archivo"] = str(Path(evidencia["archivo"]).relative_to(licitacion))
                registro["pagina"] = evidencia.get("pagina")
                registro["ubicado"] = evidencia["ubicado"]
                registro["nota"] = evidencia.get("nota", "")
                if evidencia.get("imagen"):
                    registro["imagen"] = enlace_relativo(evidencia["imagen"], salida)
                if evidencia.get("imagen_original"):
                    registro["imagen"] = enlace_relativo(evidencia["imagen_original"], salida)
                if evidencia.get("html"):
                    registro["tabla_html"] = evidencia["html"]
                registros.append(registro)
        print(f"  {licitacion.name}: {len(resultados)} proveedores", end="\r")
    documentos.cerrar()

    plantilla = PLANTILLA_HTML
    datos = json.dumps(registros, ensure_ascii=False, default=str).replace("</", "<\\/")
    clave = f"revision_v3_{args.fuente}_{raiz.parent.name}"  # v3: la v2 mostraba imagenes cruzadas (RUT con puntos)
    pagina = (plantilla.replace("__DATOS__", datos)
              .replace("__CLAVE__", json.dumps(clave))
              .replace("__TITULO__", html.escape(f"Revision {args.fuente} - {raiz.parent.name}"))
              .replace("__MUESTRA__", str(args.muestra)))
    (salida / "index.html").write_text(pagina, encoding="utf-8")

    print(" " * 80)
    print("=" * 74)
    print(f"Productos: {len(registros)}")
    print(f"  fila verificada:               {conteo.get('exacto', 0)}")
    print(f"  datos que NO calzan con la fila: {conteo.get('parcial', 0)}   <- revisar primero")
    print(f"  precio NO encontrado:          {conteo.get('no_encontrado', 0)}   <- revisar primero")
    print(f"  sin archivo fuente:            {conteo.get('sin_archivo', 0)}")
    print(f"Abrir en Chrome o Edge: {salida / 'index.html'}")
    print("=" * 74)


PLANTILLA_HTML = r"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITULO__</title>
<style>
:root{--fondo:#f4f5f7;--tarjeta:#fff;--texto:#1d2330;--suave:#667085;--borde:#d9dde5;--ok:#1e8e3e;--mal:#c5221f;--duda:#b06000;--azul:#1f4e78}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 "Segoe UI",system-ui,sans-serif;background:var(--fondo);color:var(--texto)}
header{position:sticky;top:0;z-index:5;background:var(--azul);color:#fff;padding:10px 18px;box-shadow:0 2px 6px #0003}
header h1{margin:0 0 6px;font-size:17px}.filtros{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.filtros input,.filtros select,.filtros button{font:inherit;padding:4px 8px;border-radius:5px;border:1px solid #fff5}
.filtros button{background:#fff;color:var(--azul);cursor:pointer;border:0;font-weight:600}
.stats{margin-top:6px;font-size:13px;opacity:.95}
main{max-width:1250px;margin:14px auto;padding:0 12px}
.tarjeta{background:var(--tarjeta);border:1px solid var(--borde);border-radius:8px;margin-bottom:14px;padding:12px 14px;border-left:6px solid var(--borde)}
.tarjeta.d-correcto{border-left-color:var(--ok)}.tarjeta.d-incorrecto{border-left-color:var(--mal)}.tarjeta.d-dudoso{border-left-color:var(--duda)}
.cabecera{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.cabecera h2{font-size:15px;margin:0}.sub{color:var(--suave);font-size:12.5px}
.datos{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:4px 14px;margin:8px 0}
.datos div b{display:block;font-size:11px;color:var(--suave);font-weight:600;text-transform:uppercase}
.badge{display:inline-block;padding:1px 7px;border-radius:10px;font-size:12px;margin:1px 3px 1px 0;background:#eef1f5}
.b-ok{background:#e6f4ea;color:var(--ok)}.b-mal{background:#fce8e6;color:var(--mal)}.b-duda{background:#fef7e0;color:var(--duda)}
.evidencia{margin-top:8px;overflow-x:auto;border:1px solid var(--borde);border-radius:6px;background:#fafbfc;padding:6px}
.evidencia img{max-width:100%;display:block}
.leyenda{font-size:12px;color:var(--suave);margin:4px 0}
.leyenda span{display:inline-block;width:11px;height:11px;border-radius:2px;vertical-align:-1px;margin:0 3px 0 8px}
table.grilla{border-collapse:collapse;font-size:12.5px}table.grilla td,table.grilla th{border:1px solid #cfd5de;padding:3px 6px;max-width:320px}
table.grilla th{background:#eef1f5;color:var(--suave);font-weight:600}.num{color:var(--suave);background:#f3f5f8;text-align:right}
tr.encabezado td{background:#eceff3;font-weight:600}tr.objetivo td{outline:2px solid #db1a1a;outline-offset:-1px;background:#fff6f6}
td.c-unitario{background:#d6f2dc!important}td.c-cantidad{background:#dbe6fb!important}td.c-total{background:#ecdcf6!important}
tr.salto td{text-align:center;color:var(--suave);border:0}.tabla-titulo{font-size:12px;color:var(--suave);margin-bottom:3px}
.acciones{display:flex;gap:8px;align-items:center;margin-top:8px;flex-wrap:wrap}
.acciones button{font:inherit;padding:5px 12px;border-radius:5px;border:1px solid var(--borde);background:#fff;cursor:pointer}
.acciones button.sel-correcto{background:var(--ok);color:#fff;border-color:var(--ok)}
.acciones button.sel-incorrecto{background:var(--mal);color:#fff;border-color:var(--mal)}
.acciones button.sel-dudoso{background:var(--duda);color:#fff;border-color:var(--duda)}
.acciones input{flex:1;min-width:200px;font:inherit;padding:5px 8px;border:1px solid var(--borde);border-radius:5px}
.aviso{color:var(--mal);font-weight:600}.paginacion{text-align:center;margin:10px 0 30px}
.paginacion button{font:inherit;padding:5px 12px;margin:0 4px}
a{color:var(--azul)}
</style></head><body>
<header>
  <h1>__TITULO__</h1>
  <div class="filtros">
    <input id="f-texto" placeholder="Buscar proveedor, producto, modelo..." size="30">
    <select id="f-ubicado"><option value="">Ubicación: todas</option><option value="no_encontrado">Precio NO encontrado</option><option value="parcial">Datos no calzan con la fila</option><option value="exacto">Fila verificada</option><option value="sin_archivo">Sin archivo</option></select>
    <select id="f-validacion"><option value="">Validación: todas</option><option>ok</option><option>revisar</option><option>incompleto</option><option>inconsistente</option></select>
    <select id="f-otro"><option value="">Otro método: todos</option><option value="precio_distinto">Precio distinto</option><option value="mismo_precio">Mismo precio</option><option value="otro_sin_precios">Otro sin precios</option></select>
    <select id="f-decision"><option value="">Revisión: todas</option><option value="pendiente">Sin revisar</option><option value="correcto">Correcto</option><option value="incorrecto">Incorrecto</option><option value="dudoso">Dudoso</option></select>
    <label><input type="checkbox" id="f-muestra"> Solo muestra aleatoria (__MUESTRA__)</label>
    <button id="exportar">Exportar CSV</button>
  </div>
  <div class="stats" id="stats"></div>
</header>
<main><div id="lista"></div><div class="paginacion" id="paginacion"></div></main>
<script>
const DATOS = __DATOS__;
const CLAVE = __CLAVE__;
const TAM_MUESTRA = Math.min(__MUESTRA__, DATOS.length);
const POR_PAGINA = 40;
let decisiones = {};
try { decisiones = JSON.parse(localStorage.getItem(CLAVE) || "{}"); } catch (e) { decisiones = {}; }
function guardar() { try { localStorage.setItem(CLAVE, JSON.stringify(decisiones)); } catch (e) {} }

// Muestra aleatoria reproducible (misma semilla = misma muestra cada vez)
function mulberry32(a){return function(){a|=0;a=a+0x6D2B79F5|0;let t=Math.imul(a^a>>>15,1|a);t=t+Math.imul(t^t>>>7,61|t)^t;return((t^t>>>14)>>>0)/4294967296}}
const azar = mulberry32(20251101);
const indices = DATOS.map((_, i) => i);
for (let i = indices.length - 1; i > 0; i--) { const j = Math.floor(azar() * (i + 1)); [indices[i], indices[j]] = [indices[j], indices[i]]; }
const EN_MUESTRA = new Set(indices.slice(0, TAM_MUESTRA).map(i => DATOS[i].id));

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const monto = v => (typeof v === "number") ? v.toLocaleString("es-CL") : (v ?? "—");
let pagina = 0;

function filtrados() {
  const texto = $("f-texto").value.toLowerCase(), ub = $("f-ubicado").value, val = $("f-validacion").value,
        otro = $("f-otro").value, dec = $("f-decision").value, muestra = $("f-muestra").checked;
  return DATOS.filter(d => {
    const decision = (decisiones[d.id] || {}).decision || "pendiente";
    if (muestra && !EN_MUESTRA.has(d.id)) return false;
    if (ub && d.ubicado !== ub) return false;
    if (val && d.estado_validacion !== val) return false;
    if (otro && d.otro_metodo !== otro) return false;
    if (dec && decision !== dec) return false;
    if (texto && ![d.proveedor, d.producto, d.marca, d.modelo, d.licitacion, d.rut].join(" ").toLowerCase().includes(texto)) return false;
    return true;
  });
}

function wilson(ok, n) {
  if (!n) return null; const z = 1.96, p = ok / n, den = 1 + z*z/n;
  const centro = (p + z*z/(2*n)) / den, margen = z * Math.sqrt(p*(1-p)/n + z*z/(4*n*n)) / den;
  return [Math.max(0, centro - margen), Math.min(1, centro + margen)];
}

function estadisticas(lista) {
  const cuenta = {correcto:0, incorrecto:0, dudoso:0};
  DATOS.forEach(d => { const x = (decisiones[d.id]||{}).decision; if (x) cuenta[x]++; });
  const m = {correcto:0, incorrecto:0};
  DATOS.forEach(d => { if (EN_MUESTRA.has(d.id)) { const x = (decisiones[d.id]||{}).decision; if (x in m) m[x]++; } });
  const n = m.correcto + m.incorrecto, ic = wilson(m.correcto, n);
  let txt = `Mostrando ${lista.length} de ${DATOS.length} · Revisados: ${cuenta.correcto + cuenta.incorrecto + cuenta.dudoso} ` +
            `(✔ ${cuenta.correcto} · ✘ ${cuenta.incorrecto} · ? ${cuenta.dudoso})`;
  if (n) txt += ` · Muestra: ${n}/${TAM_MUESTRA} revisados, precisión estimada ${(100*m.correcto/n).toFixed(0)}% (IC 95%: ${(100*ic[0]).toFixed(0)}–${(100*ic[1]).toFixed(0)}%)`;
  $("stats").textContent = txt;
}

function badgeUbicado(d) {
  if (d.ubicado === "exacto") return '<span class="badge b-ok">fila verificada</span>';
  if (d.ubicado === "parcial") return '<span class="badge b-mal">datos no calzan con la fila</span>';
  if (d.ubicado === "sin_archivo") return '<span class="badge b-mal">sin archivo fuente</span>';
  return '<span class="badge b-mal">precio NO encontrado en el documento</span>';
}
function badgeOtro(d) {
  const t = {mismo_precio:["b-ok","el otro método encontró el mismo precio"], precio_distinto:["b-mal","el otro método encontró otro precio"],
             otro_sin_precios:["b-duda","el otro método no encontró precios"], sin_unitario:["",""], sin_datos:["",""]}[d.otro_metodo] || ["",""];
  return t[1] ? `<span class="badge ${t[0]}">${t[1]}</span>` : "";
}

function tarjeta(d) {
  const dec = decisiones[d.id] || {};
  const val = {ok:"b-ok", revisar:"b-duda", incompleto:"b-mal", inconsistente:"b-mal"}[d.estado_validacion] || "";
  let evidencia = "";
  if (d.imagen) evidencia = `<div class="leyenda">Fila<span style="border:2px solid #db1a1a"></span> Cantidad<span style="background:#1a59e6"></span> P. unitario<span style="background:#0da633"></span> Total<span style="background:#8c33bf"></span></div><img loading="lazy" src="${esc(d.imagen)}">`;
  else if (d.tabla_html) evidencia = d.tabla_html;
  else evidencia = `<div class="aviso">${esc(d.nota || "sin evidencia")}</div><div class="sub">Evidencia guardada: ${esc(d.evidencia_texto || "—")}</div>`;
  return `<div class="tarjeta ${dec.decision ? "d-" + dec.decision : ""}" id="t-${esc(d.id)}">
    <div class="cabecera"><div><h2>${esc(d.producto)}</h2>
      <div class="sub">${esc(d.licitacion)} · ${esc(d.proveedor)} (${esc(d.rut)}) · ${esc(d.archivo || "")}${d.pagina ? " · pág. " + d.pagina : ""}
      ${d.abrir ? ` · <a href="${esc(d.abrir)}" target="_blank">abrir documento</a>` : ""}</div></div>
      <div>${badgeUbicado(d)}<span class="badge ${val}">validación: ${esc(d.estado_validacion || "—")}</span>${badgeOtro(d)}${EN_MUESTRA.has(d.id) ? '<span class="badge">en muestra</span>' : ""}</div></div>
    <div class="datos"><div><b>Marca</b>${esc(d.marca || "—")}</div><div><b>Modelo</b>${esc(d.modelo || "—")}</div>
      <div><b>Tipo</b>${esc(d.tipo || "—")}</div><div><b>Cantidad</b>${monto(d.cantidad)}</div>
      <div><b>Precio unitario</b>${monto(d.precio_unitario)} ${esc(d.moneda || "")}</div><div><b>Total</b>${monto(d.precio_total)}</div>
      <div><b>Método</b>${esc(d.metodo || "—")}</div></div>
    ${d.alertas.length ? `<div>${d.alertas.map(a => `<span class="badge b-duda">${esc(a)}</span>`).join("")}</div>` : ""}
    ${d.nota && (d.imagen || d.tabla_html) ? `<div class="${d.ubicado === "parcial" ? "aviso" : "sub"}">${esc(d.nota)}</div>` : ""}
    <div class="evidencia">${evidencia}</div>
    <div class="acciones">
      <button data-id="${esc(d.id)}" data-d="correcto" class="${dec.decision==="correcto"?"sel-correcto":""}">✔ Correcto</button>
      <button data-id="${esc(d.id)}" data-d="incorrecto" class="${dec.decision==="incorrecto"?"sel-incorrecto":""}">✘ Incorrecto</button>
      <button data-id="${esc(d.id)}" data-d="dudoso" class="${dec.decision==="dudoso"?"sel-dudoso":""}">? Dudoso</button>
      <input data-id="${esc(d.id)}" class="comentario" placeholder="Comentario (qué está mal)" value="${esc(dec.comentario || "")}">
    </div></div>`;
}

function render() {
  const lista = filtrados();
  const paginas = Math.max(1, Math.ceil(lista.length / POR_PAGINA));
  pagina = Math.min(pagina, paginas - 1);
  $("lista").innerHTML = lista.slice(pagina * POR_PAGINA, (pagina + 1) * POR_PAGINA).map(tarjeta).join("") || "<p>Sin resultados con estos filtros.</p>";
  $("paginacion").innerHTML = paginas > 1 ? `<button ${pagina===0?"disabled":""} onclick="pagina--;render();scrollTo(0,0)">← Anterior</button> Página ${pagina+1} de ${paginas} <button ${pagina>=paginas-1?"disabled":""} onclick="pagina++;render();scrollTo(0,0)">Siguiente →</button>` : "";
  estadisticas(lista);
}

document.addEventListener("click", e => {
  const b = e.target.closest("button[data-d]"); if (!b) return;
  const id = b.dataset.id, d = b.dataset.d, actual = decisiones[id] || {};
  decisiones[id] = {...actual, decision: actual.decision === d ? undefined : d};
  guardar();
  const t = document.getElementById("t-" + id);
  t.outerHTML = tarjeta(DATOS.find(x => x.id === id));
  estadisticas(filtrados());
});
document.addEventListener("change", e => {
  if (!e.target.classList.contains("comentario")) return;
  const id = e.target.dataset.id; decisiones[id] = {...(decisiones[id] || {}), comentario: e.target.value}; guardar();
});
["f-texto","f-ubicado","f-validacion","f-otro","f-decision","f-muestra"].forEach(id =>
  $(id).addEventListener(id === "f-texto" ? "input" : "change", () => { pagina = 0; render(); }));
$("exportar").addEventListener("click", () => {
  const cols = ["id","licitacion","proveedor","rut","producto","tipo","marca","modelo","cantidad","precio_unitario","precio_total","estado_validacion","metodo","ubicado","otro_metodo","archivo","pagina","en_muestra","decision","comentario"];
  const filas = DATOS.map(d => { const x = decisiones[d.id] || {};
    return cols.map(c => c === "decision" ? (x.decision || "") : c === "comentario" ? (x.comentario || "") : c === "en_muestra" ? (EN_MUESTRA.has(d.id) ? "si" : "no") : (d[c] ?? "")); });
  const csv = "\ufeff" + [cols, ...filas].map(f => f.map(v => `"${String(v).replace(/"/g, '""')}"`).join(";")).join("\r\n");
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv], {type: "text/csv"}));
  a.download = CLAVE + ".csv"; a.click();
});
render();
</script></body></html>
"""

if __name__ == "__main__":
    main()

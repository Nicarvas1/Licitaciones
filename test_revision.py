"""
Pruebas del generador de revision (4_revisar_extraccion.py).

Ejecutar desde la raiz del proyecto:
    python -m unittest tests.test_revision -v
"""

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

spec = importlib.util.spec_from_file_location("revisar_extraccion", RAIZ / "4_revisar_extraccion.py")
rev = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rev)
fitz = rev.fitz
from openpyxl import Workbook  # noqa: E402


def pdf_tabla(ruta):
    doc = fitz.open()
    pagina = doc.new_page(width=842, height=595)
    filas = [["Item", "Descripcion", "Cant.", "Precio unitario", "Total"],
             ["1", "Notebook Dell Latitude 3540", "10", "$ 690.000", "$ 6.900.000"],
             ["2", "Monitor Dell P2422H", "10", "$ 80.000", "$ 800.000"]]
    xs = [40, 80, 380, 450, 560, 660]
    for i, fila in enumerate(filas):
        for j, valor in enumerate(fila):
            pagina.insert_text((xs[j] + 3, 70 + i * 22 + 15), valor, fontsize=8)
    doc.save(str(ruta))
    doc.close()


class TestUtilidades(unittest.TestCase):
    def test_variantes_monto(self):
        self.assertEqual(rev.variantes_monto(650000), ["650.000", "650,000", "650000"])
        self.assertEqual(rev.variantes_monto(None), [])

    def test_texto_contiene_monto(self):
        self.assertTrue(rev.texto_contiene_monto("1 Notebook 10 $ 690.000 $ 6.900.000", 690000))
        self.assertFalse(rev.texto_contiene_monto("2 Monitor 10 $ 80.000 $ 800.000", 690000))


class TestEvidenciaPdf(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.pdf = self.dir / "oferta.pdf"
        pdf_tabla(self.pdf)
        self.docs = rev.Documentos()

    def tearDown(self):
        self.docs.cerrar()
        self.tmp.cleanup()

    def test_ubicacion_exacta_verificada(self):
        producto = {"precio_unitario": 690000, "ubicacion": {"pagina": 1, "bbox_fila": [40, 92, 660, 114]}}
        evidencia = rev.evidencia_pdf(self.docs, self.pdf, producto, self.dir / "a", 72)
        self.assertEqual(evidencia["ubicado"], "exacto")
        self.assertTrue(Path(evidencia["imagen"]).is_file())

    def test_coordenadas_erroneas_no_afectan(self):
        # el rectangulo guardado apunta a la fila del monitor: se ignora y se ubica por los montos
        producto = {"precio_unitario": 690000, "precio_total": 6900000, "cantidad": 10,
                    "producto": "Notebook Dell Latitude 3540",
                    "ubicacion": {"pagina": 1, "bbox_fila": [40, 114, 660, 136]}}
        evidencia = rev.evidencia_pdf(self.docs, self.pdf, producto, self.dir / "b", 72)
        self.assertEqual(evidencia["ubicado"], "exacto")

    def test_datos_mezclados_de_otra_fila(self):
        # precio del notebook pero cantidad y total del monitor
        producto = {"precio_unitario": 690000, "precio_total": 800000, "cantidad": 10,
                    "producto": "Notebook Dell Latitude 3540", "pagina": 1}
        evidencia = rev.evidencia_pdf(self.docs, self.pdf, producto, self.dir / "m", 72)
        self.assertEqual(evidencia["ubicado"], "parcial")
        self.assertIn("total", evidencia["nota"])

    def test_elige_la_linea_correcta_si_el_precio_se_repite(self):
        doc = fitz.open()
        pagina = doc.new_page()
        pagina.insert_text((40, 80), "Total neto 800.000", fontsize=9)
        pagina.insert_text((40, 120), "2 Monitor Dell P2422H 10 80.000 800.000", fontsize=9)
        ruta = self.dir / "repetido.pdf"
        doc.save(str(ruta))
        doc.close()
        producto = {"precio_unitario": 80000, "precio_total": 800000, "cantidad": 10,
                    "producto": "Monitor Dell P2422H", "pagina": 1}
        evidencia = rev.evidencia_pdf(self.docs, ruta, producto, self.dir / "r", 72)
        self.assertEqual(evidencia["ubicado"], "exacto")

    def test_pagina_rotada(self):
        doc = fitz.open()
        pagina = doc.new_page(width=595, height=842)
        # contenido escrito girado para que, con /Rotate 90, se lea horizontal
        pagina.set_rotation(90)
        for i, linea in enumerate(["Item Descripcion Cant. Unitario Total",
                                   "1 Notebook Dell Latitude 3540 10 690.000 6.900.000",
                                   "2 Monitor Dell P2422H 10 80.000 800.000"]):
            punto = fitz.Point(60, 80 + i * 20) * pagina.derotation_matrix
            pagina.insert_text(punto, linea, fontsize=9, rotate=90)
        ruta = self.dir / "rotada.pdf"
        doc.save(str(ruta))
        doc.close()
        producto = {"precio_unitario": 80000, "precio_total": 800000, "cantidad": 10,
                    "producto": "Monitor Dell P2422H", "pagina": 1}
        evidencia = rev.evidencia_pdf(self.docs, ruta, producto, self.dir / "rot", 72)
        self.assertEqual(evidencia["ubicado"], "exacto")
        self.__class__.imagen_rotada = evidencia["imagen"]

    def test_precio_inventado_no_se_encuentra(self):
        evidencia = rev.evidencia_pdf(self.docs, self.pdf, {"precio_unitario": 615000, "pagina": 1}, self.dir / "c", 72)
        self.assertEqual(evidencia["ubicado"], "no_encontrado")
        self.assertNotIn("imagen", evidencia)


class TestEvidenciaExcel(unittest.TestCase):
    def test_resalta_fila_y_celdas(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "oferta.xlsx"
            libro = Workbook()
            hoja = libro.active
            hoja.title = "Oferta"
            hoja.append(["Descripcion", "Cantidad", "Precio unitario", "Total"])
            hoja.append(["Notebook HP ProBook 440 G11", 10, 650000, 6500000])
            libro.save(ruta)
            producto = {"precio_unitario": 650000, "ubicacion": {
                "hoja": "Oferta", "fila": 1, "fila_encabezado": 0,
                "columnas": {"descripcion": 0, "cantidad": 1, "unitario": 2, "total": 3}}}
            docs = rev.Documentos()
            evidencia = rev.evidencia_tabular(docs, ruta, producto)
        self.assertEqual(evidencia["ubicado"], "exacto")
        self.assertIn('class="objetivo"', evidencia["html"])
        self.assertIn('class="c-unitario"', evidencia["html"])

    def test_fila_excel_con_datos_de_otra_fila(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "oferta.xlsx"
            libro = Workbook()
            hoja = libro.active
            hoja.append(["Descripcion", "Cantidad", "Precio unitario", "Total"])
            hoja.append(["Notebook HP ProBook 440 G11", 10, 650000, 6500000])
            hoja.append(["Monitor HP P24 G5", 5, 120000, 600000])
            libro.save(ruta)
            producto = {"precio_unitario": 650000, "cantidad": 5, "precio_total": 600000}
            evidencia = rev.evidencia_tabular(rev.Documentos(), ruta, producto)
        self.assertEqual(evidencia["ubicado"], "parcial")


if __name__ == "__main__":
    unittest.main()

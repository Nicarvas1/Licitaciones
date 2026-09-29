import importlib.util
import tempfile
import unittest
from pathlib import Path

import pymupdf


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("revisor", ROOT / "4_revisar_extraccion.py")
revisor = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(revisor)


class RevisorExtraccionTest(unittest.TestCase):
    def test_modelos_con_misma_familia_no_se_verifican_por_marca_sola(self):
        producto = {"producto": "Notebook HP ProBook 440 G11", "modelo": "ProBook 440 G11", "marca": "HP"}
        estado, _ = revisor.verificar_descripcion(producto, "Notebook HP ProBook 450 G11 2 100000 200000")
        self.assertNotEqual(estado, "exacto")

    def test_montos_sin_descripcion_quedan_para_revision(self):
        estado, _ = revisor.verificar_descripcion({"precio_unitario": 200000}, "Monitor Samsung 200000")
        self.assertEqual(estado, "parcial")

    def test_busqueda_pdf_no_confirma_producto_distinto_sin_bbox(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "tabla.pdf"
            with pymupdf.open() as doc:
                doc.new_page().insert_text((50, 100), "Monitor Samsung 27 1 200000 200000")
                doc.save(ruta)
            docs = revisor.Documentos()
            try:
                producto = {"producto": "Impresora Brother HL1202", "cantidad": 1,
                            "precio_unitario": 200000, "precio_total": 200000}
                evidencia = revisor.evidencia_pdf(docs, ruta, producto, Path(directorio) / "evidencia", 72)
                self.assertEqual(evidencia["ubicado"], "descripcion_no_coincide")
            finally:
                docs.cerrar()

    def test_tabla_no_confirma_producto_distinto_por_precio(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "oferta.xlsx"
            libro = Workbook()
            libro.active.append(["Monitor Samsung 27", 1, 200000, 200000])
            libro.save(ruta)
            producto = {"producto": "Impresora Brother HL1202", "cantidad": 1,
                        "precio_unitario": 200000, "precio_total": 200000}
            evidencia = revisor.evidencia_tabular(revisor.Documentos(), ruta, producto)
            self.assertEqual(evidencia["ubicado"], "descripcion_no_coincide")

    def test_tabla_elige_producto_correcto_si_se_repite_precio(self):
        producto = {"producto": "Impresora Brother HL1202", "cantidad": 1,
                    "precio_unitario": 200000, "precio_total": 200000}
        filas = [["Monitor Samsung 27", 1, 200000, 200000],
                 ["Impresora Brother HL1202", 1, 200000, 200000]]
        self.assertEqual(revisor.fila_con_monto(filas, producto), 1)

    def test_html_muestra_descripcion_distinta_y_conserva_mejor_evidencia(self):
        import contextlib
        import io
        import json
        import sys
        from unittest import mock
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as directorio:
            raiz = Path(directorio)
            ofertas = raiz / "ofertas"
            licitacion = ofertas / "LIC-1"
            carpeta = licitacion / "1__Proveedor"
            carpeta.mkdir(parents=True)
            (carpeta / "oferta.json").write_text(json.dumps({"rut": "1", "proveedor": "Proveedor"}), encoding="utf-8")
            libro = Workbook()
            libro.active.append(["Monitor Samsung 27", 1, 200000, 200000])
            libro.save(carpeta / "economico.xlsx")
            (carpeta / "tecnico.txt").write_text("Declaracion sin precios", encoding="utf-8")
            producto = {"producto": "Impresora Brother HL1202", "cantidad": 1,
                        "precio_unitario": 200000, "precio_total": 200000,
                        "archivo_fuente": "economico.xlsx", "fuentes_respaldo": ["tecnico.txt"]}
            (licitacion / "extraccion_ia.json").write_text(
                json.dumps([{"rut": "1", "proveedor": "Proveedor", "productos": [producto]}]), encoding="utf-8")
            salida = raiz / "revision"
            with mock.patch.object(sys, "argv", ["4_revisar_extraccion.py", "--dir", str(ofertas),
                                                  "--fuente", "ia", "--salida", str(salida)]), \
                    contextlib.redirect_stdout(io.StringIO()):
                revisor.main()
            pagina = (salida / "index.html").read_text(encoding="utf-8")
            self.assertIn('"ubicado": "descripcion_no_coincide"', pagina)
            self.assertIn('"sugerencia_automatica": "incorrecto"', pagina)
            self.assertIn('value="descripcion_no_coincide"', pagina)

    def test_nombre_imagen_no_colisiona_por_puntos_del_rut(self):
        primero = revisor.nombre_imagen_revision("2744-72-LE25__76.596.570-5__abc")
        segundo = revisor.nombre_imagen_revision("2744-72-LE25__76.596.570-5__xyz")
        self.assertNotEqual(primero, segundo)
        self.assertNotIn(".", primero)

    def test_evidencia_pdf_prefiere_bbox_original(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta_pdf = Path(directorio) / "tabla.pdf"
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_text((50, 100), "Monitor A 1 100000 100000")
            pagina.insert_text((50, 200), "Impresora B 2 200000 400000")
            documento.save(ruta_pdf)
            documento.close()

            producto = {
                "producto": "Impresora B",
                "cantidad": 2,
                "precio_unitario": 200000,
                "precio_total": 400000,
                "ubicacion": {
                    "pagina": 1,
                    "bbox_fila": [40, 180, 400, 215],
                    "bbox_celdas": {
                        "cantidad": [125, 180, 145, 215],
                        "unitario": [145, 180, 220, 215],
                        "total": [220, 180, 300, 215],
                    },
                },
            }
            documento = pymupdf.open(ruta_pdf)
            evidencia = revisor.evidencia_pdf_coordenadas(
                documento, producto, Path(directorio) / "evidencia", 110
            )
            documento.close()

            self.assertIsNotNone(evidencia)
            self.assertIn("coordenadas originales", evidencia["nota"])
            self.assertTrue(Path(evidencia["imagen"]).is_file())

    def test_evidencia_pdf_detecta_descripcion_de_otro_producto(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta_pdf = Path(directorio) / "tabla.pdf"
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_text((50, 100), "Monitor Samsung 27 1 200000 200000")
            pagina.insert_text((50, 200), "Impresora Brother HL1202 8 52900 423200")
            documento.save(ruta_pdf)
            documento.close()

            # El bbox apunta a la fila del monitor, pero el producto describe la
            # impresora: mismos precio/cantidad no deben bastar para marcarlo exacto.
            producto = {
                "producto": "Impresora Brother HL1202",
                "cantidad": 1,
                "precio_unitario": 200000,
                "precio_total": 200000,
                "ubicacion": {
                    "pagina": 1,
                    "bbox_fila": [40, 80, 400, 115],
                },
            }
            documento = pymupdf.open(ruta_pdf)
            evidencia = revisor.evidencia_pdf_coordenadas(
                documento, producto, Path(directorio) / "evidencia", 110
            )
            documento.close()

            self.assertIsNotNone(evidencia)
            self.assertEqual(evidencia["ubicado"], "descripcion_no_coincide")
            self.assertIn("otro producto", evidencia["nota"])

    def test_decision_automatica_clasifica_correcto_incorrecto_y_dudoso(self):
        self.assertEqual(
            revisor.decidir_automatica_producto(
                {"producto": "Monitor 27"}, {"ubicado": "exacto"}, "mismo_precio"
            ),
            "correcto",
        )
        self.assertEqual(
            revisor.decidir_automatica_producto(
                {"producto": "Monitor 27"}, {"ubicado": "descripcion_no_coincide"}, "precio_distinto"
            ),
            "incorrecto",
        )
        self.assertEqual(
            revisor.decidir_automatica_producto(
                {"producto": "Monitor 27"}, {"ubicado": "parcial"}, "mismo_precio"
            ),
            "dudoso",
        )


if __name__ == "__main__":
    unittest.main()

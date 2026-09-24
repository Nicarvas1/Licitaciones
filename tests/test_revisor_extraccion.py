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
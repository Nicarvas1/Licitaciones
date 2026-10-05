import json
import tempfile
import unittest
from pathlib import Path

import pymupdf

import cobertura_documental as cob


def proveedor(raiz, archivos, rut="76.000.000-1", licitacion="100-1-LE26"):
    carpeta = Path(raiz) / licitacion / f"{rut}__PROVEEDOR"
    carpeta.mkdir(parents=True)
    (carpeta / "oferta.json").write_text(json.dumps({"rut": rut}), encoding="utf-8")
    for nombre, texto in archivos.items():
        if nombre.endswith(".pdf"):
            with pymupdf.open() as doc:
                pagina = doc.new_page()
                if texto:
                    pagina.insert_text((50, 100), texto)
                doc.save(carpeta / nombre)
        else:
            (carpeta / nombre).write_text(texto, encoding="utf-8")
    return carpeta


def producto(**extra):
    base = {"producto": "Notebook HP ProBook 440 G11", "modelo": "440 G11", "marca": "HP", "cantidad": 5,
            "precio_unitario": 500000, "precio_total": 2500000, "estado_validacion": "ok", "alertas": []}
    base.update(extra)
    return base


class CoberturaProductoTest(unittest.TestCase):
    def cobertura(self, archivos, prod, resultado=None):
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = proveedor(tmp, archivos)
            resumen = cob.analizar_proveedor(carpeta, resultado or {}, [prod])
            return resumen["cobertura"][0], resumen

    def test_producto_completo_y_ok(self):
        (categoria, motivo), _ = self.cobertura({"economico__01__a.txt": "Notebook HP ProBook 440 G11 5 500.000 2.500.000"},
                                                producto())
        self.assertEqual(categoria, cob.COMPLETO)

    def test_ficha_tecnica_sin_precios_es_dato_que_no_existe(self):
        (categoria, motivo), resumen = self.cobertura(
            {"tecnico__01__ficha.txt": "Notebook HP ProBook 440 G11 procesador Intel Core i5 memoria RAM 16 GB SSD 512 GB"},
            producto(precio_unitario=None, precio_total=None, cantidad=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.DATO_NO_EXISTE)
        self.assertEqual(resumen["clase"], "solo_tecnico")

    def test_precio_junto_al_producto_es_no_encontrado(self):
        (categoria, motivo), _ = self.cobertura(
            {"economico__01__oferta.txt": "Notebook HP ProBook 440 G11 | 5 | $ 500.000 | $ 2.500.000"},
            producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.NO_ENCONTRADO)
        self.assertIn("economico__01__oferta.txt", motivo)

    def test_texto_ambiental_con_total_no_cuenta_como_precio(self):
        (categoria, _), resumen = self.cobertura(
            {"tecnico__01__ficha.txt": "Notebook HP ProBook 440 G11\nTotal de CO2e (kg) 1.200 emisiones"},
            producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.DATO_NO_EXISTE)
        self.assertEqual(resumen["con_precios"], [])

    def test_linea_economica_generica_con_precio_es_candidata_al_cruce(self):
        (categoria, motivo), _ = self.cobertura(
            {"economico__01__oferta.txt": "Computador | 5 | $ 500.000 | $ 2.500.000",
             "tecnico__01__ficha.txt": "Notebook HP ProBook 440 G11 procesador Intel Core i5"},
            producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.NO_ENCONTRADO)
        self.assertIn("generica", motivo)

    def test_formulario_vertical_con_valor_unitario_es_no_encontrado(self):
        (categoria, motivo), _ = self.cobertura(
            {"economico__01__anexo.txt": "FILA 5: Valor unitario || $ 3.500.059" + chr(10) + "FILA 6: Valor Neto || $ 7.000.118",
             "tecnico__01__ficha.txt": "Notebook HP ProBook 440 G11 procesador Intel Core i5"},
            producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.NO_ENCONTRADO)
        self.assertIn("valor unitario", motivo)

    def test_pdf_sin_texto_deja_el_dato_pendiente_de_ocr(self):
        (categoria, motivo), resumen = self.cobertura(
            {"economico__01__escaneado.pdf": ""},
            producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.PENDIENTE_OCR)
        self.assertIn("con_documentos_sin_texto", resumen["clase"])

    def test_orden_de_compra_historica_no_cuenta_como_evidencia(self):
        (categoria, _), resumen = self.cobertura(
            {"tecnico__01__EXP_OCOMPRA.txt": "Nombre Orden de Compra\nNotebook HP ProBook 440 G11 5 $ 500.000 $ 2.500.000"},
            producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.DATO_NO_EXISTE)
        self.assertIn("con_historial", resumen["clase"])

    def test_marca_en_la_linea_del_producto_es_no_encontrada(self):
        (categoria, _), _ = self.cobertura(
            {"economico__01__oferta.txt": "Impresora Epson EcoTank L4260 | 1 | $ 179.600 | $ 179.600"},
            producto(producto="Impresora EcoTank L4260", modelo="L4260", marca=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.NO_ENCONTRADO)

    def test_inconsistente_con_datos_completos_es_contradictorio(self):
        (categoria, _), _ = self.cobertura({"economico__01__a.txt": "Notebook HP ProBook 440 G11 5 500.000 2.500.000"},
                                           producto(estado_validacion="inconsistente", alertas=["suma supera el total"]))
        self.assertEqual(categoria, cob.CONTRADICTORIO)

    def test_proveedor_sin_documentos(self):
        (categoria, motivo), resumen = self.cobertura(
            {}, producto(precio_unitario=None, precio_total=None, estado_validacion="incompleto"))
        self.assertEqual(categoria, cob.DATO_NO_EXISTE)
        self.assertEqual(resumen["clase"], "sin_documentos")


class ClasificacionDeDocumentosTest(unittest.TestCase):
    def test_roles_por_prefijo_del_scraper(self):
        self.assertEqual(cob.rol_del_documento("economico__01__x.pdf"), "economico")
        self.assertEqual(cob.rol_del_documento("tecnico__02__x.pdf"), "tecnico")
        self.assertEqual(cob.rol_del_documento("administrativo__01__x.pdf"), "administrativo")

    def test_un_pdf_por_producto(self):
        with tempfile.TemporaryDirectory() as tmp:
            archivos = {f"tecnico__0{i}__ficha{i}.txt": f"Notebook modelo X{i}00 procesador Intel" for i in range(1, 5)}
            resumen = cob.analizar_proveedor(proveedor(tmp, archivos), {}, [])
        self.assertIn("uno_por_producto", resumen["clase"])

    def test_montos_plausibles_ignora_cifras_chicas_y_enormes(self):
        self.assertEqual(cob.montos_plausibles("RAM 16 GB 1.200 y 2.500.000 y 9.999.999.999"), [1200, 2500000])


class AnotarProductosTest(unittest.TestCase):
    def test_anota_cobertura_y_motivo_en_cada_producto(self):
        with tempfile.TemporaryDirectory() as tmp:
            proveedor(tmp, {"economico__01__a.txt": "Notebook HP ProBook 440 G11 5 500.000 2.500.000"})
            productos = [dict(producto(), codigo="100-1-LE26", rut="76.000.000-1"),
                         dict(producto(), codigo="100-1-LE26", rut="99.999.999-9")]
            cob.anotar_productos(productos, tmp, workers=1)
        self.assertEqual(productos[0]["cobertura_documental"], cob.COMPLETO)
        self.assertEqual(productos[1]["cobertura_documental"], "")
        self.assertIn("no encontrada", productos[1]["cobertura_motivo"])


if __name__ == "__main__":
    unittest.main()

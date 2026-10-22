import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("extractor_finde", ROOT / "3_extraer_ia.py")
E = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(E)


def prod(**extra):
    base = {"producto": "Impresora multifuncional EcoTank L4260", "modelo": "L4260", "marca": None, "cantidad": 2,
            "precio_unitario": 200000, "precio_total": 400000, "categoria": "impresora", "alertas": []}
    base.update(extra)
    return base


class MarcaPorLineaTest(unittest.TestCase):
    def test_marca_escrita_junto_al_codigo_del_modelo(self):
        bloques = [{"archivo": "economico__01__a.pdf", "pagina": 1,
                    "texto": "FILA 2: Impresora Epson EcoTank L4260 || 2 || 200.000 || 400.000"}]
        resultado = E.completar_marcas_por_linea([prod()], bloques)
        self.assertEqual(resultado[0]["marca"], "Epson")
        self.assertEqual(resultado[0]["metodo_marca"], "linea_del_producto")

    def test_dos_marcas_distintas_en_las_lineas_del_codigo_no_se_asigna(self):
        bloques = [{"archivo": "a.pdf", "pagina": 1, "texto": "Epson L4260"},
                   {"archivo": "b.pdf", "pagina": 1, "texto": "Brother L4260 garantia 12 meses"}]
        self.assertIsNone(E.completar_marcas_por_linea([prod()], bloques)[0]["marca"])

    def test_requisito_similar_no_cuenta_como_marca(self):
        bloques = [{"archivo": "a.pdf", "pagina": 1, "texto": "IMPRESORA SIMILAR EPSON ECOTANK L4260"}]
        self.assertIsNone(E.completar_marcas_por_linea([prod()], bloques)[0]["marca"])

    def test_codigos_genericos_no_sirven_para_ubicar_el_producto(self):
        producto = prod(producto="Notebook i5-1335U 16GB DDR4 512GB", modelo=None)
        bloques = [{"archivo": "a.pdf", "pagina": 1, "texto": "Notebook HP i5-1335U 16GB DDR4"}]
        self.assertEqual(E.codigos_distintivos(producto), set())
        self.assertIsNone(E.completar_marcas_por_linea([producto], bloques)[0]["marca"])


class IvaSegunTotalDelPortalTest(unittest.TestCase):
    def test_lineas_con_iva_se_pasan_a_neto_si_la_suma_es_el_total_por_1_19(self):
        productos = [
            {"producto": "Notebook HP ProBook 440 G11", "marca": "HP", "categoria": "equipo", "cantidad": 10,
             "precio_unitario": 1_190_000, "precio_total": 11_900_000, "alertas": [], "archivo_fuente": "economico__01__a.pdf"},
            {"producto": "Monitor HP P24 G5", "marca": "HP", "categoria": "monitor", "cantidad": 10,
             "precio_unitario": 238_000, "precio_total": 2_380_000, "alertas": [], "archivo_fuente": "economico__01__a.pdf"},
        ]
        validados = E.validar_productos(productos, [], "$ 12.000.000")
        unitarios = sorted(p["precio_unitario"] for p in validados)
        self.assertEqual(unitarios, [200_000, 1_000_000])
        self.assertTrue(all("suma_productos_supera_total_oferta" not in p.get("alertas", []) for p in validados))

    def test_suma_que_no_calza_con_iva_sigue_marcada(self):
        productos = [
            {"producto": "Notebook HP ProBook 440 G11", "marca": "HP", "categoria": "equipo", "cantidad": 10,
             "precio_unitario": 1_500_000, "precio_total": 15_000_000, "alertas": [], "archivo_fuente": "economico__01__a.pdf"},
        ]
        validados = E.validar_productos(productos, [], "$ 12.000.000")
        self.assertEqual(validados[0]["precio_unitario"], 1_500_000)



class FilaAgregadaTest(unittest.TestCase):
    BLOQUE = {"archivo": "economico__01__ANEXO_N13.pdf", "pagina": 1, "texto": "\n".join([
        "VALOR TOTAL OFERTA (IMPUESTO INCLUIDO) $46.838.400",
        "FILA 1: LA OFERTA EN MONEDA NACIONAL ES DE: PRODUCTOS || VALOR CON IVA",
        "FILA 2: 52 NOTEBOOKS || $ 45.172.400",
        "FILA 3: 1 CARRO TECNOLOGICO || $ 1.666.000"])}

    def ficha(self, **extra):
        base = {"producto": "Notebook Gamer HP Victus 15 I5-12450H RAM 16GB", "marca": "HP", "categoria": "equipo",
                "cantidad": 52, "precio_unitario": None, "precio_total": None, "alertas": []}
        base.update(extra)
        return base

    def test_linea_agregada_con_iva_da_cantidad_unitario_y_total_netos(self):
        producto = E.cruzar_filas_agregadas([self.ficha()], [self.BLOQUE])[0]
        self.assertEqual((producto["cantidad"], producto["precio_unitario"], producto["precio_total"]),
                         (52, 730_000, 37_960_000))
        self.assertTrue(producto["precio_desde_linea_agregada"])

    def test_formato_en_columnas_cantidad_descripcion_total(self):
        bloque = {"archivo": "economico__01__oferta.pdf", "pagina": 1,
                  "texto": "FILA 2: 10 || Notebook || $ 8.000.000"}
        producto = E.cruzar_filas_agregadas([self.ficha(cantidad=None)], [bloque])[0]
        self.assertEqual((producto["cantidad"], producto["precio_unitario"]), (10, 800_000))

    def test_no_asigna_si_hay_dos_productos_sin_precio_del_mismo_tipo(self):
        productos = [self.ficha(), self.ficha(producto="Notebook Lenovo LOQ 15")]
        resultado = E.cruzar_filas_agregadas(productos, [self.BLOQUE])
        self.assertTrue(all(p["precio_unitario"] is None for p in resultado))

    def test_no_asigna_si_la_cantidad_no_calza(self):
        producto = E.cruzar_filas_agregadas([self.ficha(cantidad=30)], [self.BLOQUE])[0]
        self.assertIsNone(producto["precio_unitario"])

    def test_no_usa_documentos_que_no_son_economicos(self):
        bloque = dict(self.BLOQUE, archivo="tecnico__01__ficha.pdf")
        producto = E.cruzar_filas_agregadas([self.ficha()], [bloque])[0]
        self.assertIsNone(producto["precio_unitario"])



class OcrNoSeRepiteTest(unittest.TestCase):
    def test_archivo_sin_texto_tras_ocr_no_se_rehace_en_cada_corrida(self):
        previo = {"necesita_ocr": ["economico__01__escaneado.pdf"], "ocr_intentado": True}
        self.assertFalse(E.archivo_pendiente_ocr({"archivo": "economico__01__escaneado.pdf"}, previo))
        self.assertFalse(E.resultado_requiere_reproceso(dict(previo, resultados_archivos=[]), usar_ocr=True))

    def test_checkpoint_sin_ocr_intentado_si_se_rehace(self):
        previo = {"necesita_ocr": ["economico__01__escaneado.pdf"]}
        self.assertTrue(E.archivo_pendiente_ocr({"archivo": "economico__01__escaneado.pdf"}, previo))
        self.assertTrue(E.resultado_requiere_reproceso(dict(previo, resultados_archivos=[]), usar_ocr=True))



class TipoYCifrasTest(unittest.TestCase):
    def test_impresora_aio_no_es_un_computador_all_in_one(self):
        self.assertEqual(E.subcategoria_producto({"producto": "All-in-One Epson L14150", "marca": "Epson"}), "impresora")
        self.assertEqual(E.subcategoria_producto({"producto": "HP Smart Tank 580 AIO", "marca": "HP"}), "impresora")
        self.assertEqual(E.subcategoria_producto({"producto": "AIO HP ProOne 440 G9", "marca": "HP"}), "all-in-one")

    def test_fila_con_comas_de_miles_cifras_partidas_y_unitario_redondeado(self):
        texto = ("90PT03X6-M06CL0 ASUS AIO P440VAK-BPCD12X core 5 210H, 23.8 FHD, 512GB SSD WIN 11 PRO "
                 "24 $ 7 35,069 17,641,660")
        filas = E.filas_precio_texto_libre([{"archivo": "economico__01__c.pdf", "pagina": 1, "texto": texto}])
        self.assertEqual([(f["cantidad"], f["precio_unitario"], f["precio_total"]) for f in filas],
                         [(24, 735_069, 17_641_660)])



class ArchivosDuplicadosTest(unittest.TestCase):
    def test_el_mismo_pdf_como_economico_y_tecnico_se_lee_una_vez(self):
        import tempfile, types
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = Path(tmp)
            texto = "Notebook HP ProBook 440 G11 cantidad 5 valor unitario $ 500.000 total $ 2.500.000"
            (carpeta / "economico__01__oferta.txt").write_text(texto, encoding="utf-8")
            (carpeta / "tecnico__01__oferta.txt").write_text(texto, encoding="utf-8")
            args = types.SimpleNamespace(max_paginas=50, ocr=False, vision=False, tesseract_cmd=None, idioma_ocr="eng",
                                         dpi_ocr=150, vision_min_filas=2, max_filas_excel=500, max_columnas_excel=20,
                                         vision_solo_escaneadas=False, dpi_vision=150, max_fragmentos_tecnico=3,
                                         sin_filtro_tecnicos=False, max_chars_archivo=7000)
            _, registros, _, _, _ = E.preparar_bloques_proveedor(carpeta, sorted(carpeta.iterdir()), args)
        estados = {r["archivo"]: r["estado_ia"] for r in registros}
        self.assertEqual(estados["tecnico__01__oferta.txt"], "omitido_duplicado")
        self.assertNotEqual(estados["economico__01__oferta.txt"], "omitido_duplicado")


if __name__ == "__main__":
    unittest.main()

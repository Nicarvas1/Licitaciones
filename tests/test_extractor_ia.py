import importlib.util
import json
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("extractor_ia", ROOT / "3_extraer_ia.py")
extractor = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(extractor)


class ExtractorIATest(unittest.TestCase):
    def test_alcance_comercial(self):
        casos = {
            "Notebook HP con mouse incluido": "equipo",
            "Asus B1503CVA-S75898X": "equipo",
            "Desktop Dell OptiPlex": "equipo",
            "All in One Lenovo": "equipo",
            "Monitor HP 24 pulgadas": "monitor",
            "Impresora Epson L5590": "impresora",
            "Notebook HP + Smart TV Samsung": None,
            "Smart TV con modo PC": None,
            "Toner para impresora HP": None,
            "Servicio de instalacion": None,
        }
        for descripcion, esperado in casos.items():
            with self.subTest(descripcion=descripcion):
                producto = {"producto": descripcion, "categoria": "equipo"}
                self.assertEqual(extractor.clasificar_producto(producto), esperado)

    def test_normaliza_separadores_chilenos(self):
        self.assertEqual(extractor.normalizar_numero("$ 18.527.280"), 18527280)
        self.assertEqual(extractor.normalizar_numero("$617,576"), 617576)
        self.assertEqual(extractor.normalizar_numero("123,50"), 123.5)

    def test_deriva_unitario_y_cantidad_solo_con_evidencia(self):
        desde_total = extractor.normalizar_producto(
            {
                "producto": "Notebook HP",
                "cantidad": 4,
                "precio_total": 400000,
                "precio_total_tipo": "linea",
            },
            "Proveedor", "1", "a.pdf", "economico",
        )
        desde_ambos = extractor.normalizar_producto(
            {
                "producto": "Monitor HP",
                "precio_unitario": 100000,
                "precio_total": 400000,
                "precio_total_tipo": "linea",
            },
            "Proveedor", "1", "b.pdf", "economico",
        )
        solo_total = extractor.normalizar_producto(
            {"producto": "Impresora Epson", "precio_total": 400000},
            "Proveedor", "1", "c.pdf", "economico",
        )
        self.assertEqual(desde_total["precio_unitario"], 100000)
        self.assertEqual(desde_ambos["cantidad"], 4)
        self.assertEqual(desde_ambos["cantidad_fuente"], "inferida_total_dividido_unitario")
        self.assertIsNone(solo_total["precio_unitario"])

        total_oferta = extractor.normalizar_producto(
            {
                "producto": "Notebook Dell",
                "cantidad": 10,
                "precio_total": 10000000,
                "precio_total_tipo": "oferta",
            },
            "Proveedor", "1", "d.pdf", "economico",
        )
        self.assertIsNone(total_oferta["precio_unitario"])

    def test_parsea_json_embebido_sin_regex_codiciosa(self):
        respuesta = 'diagnostico {"error": true}\nresultado {"productos": [{"producto": "Notebook"}]}'
        datos, estado = extractor.parsear_json(respuesta)
        self.assertEqual(estado, "ok_extraido")
        self.assertEqual(datos["productos"][0]["producto"], "Notebook")

    def test_fragmentacion_no_pierde_filas(self):
        texto = "\n".join(f"FILA {indice}: producto {indice}" for indice in range(100))
        fragmentos = extractor.dividir_texto(texto, 300, 2)
        combinado = "\n".join(fragmentos)
        self.assertGreater(len(fragmentos), 1)
        for indice in range(100):
            self.assertIn(f"FILA {indice}:", combinado)

    def test_detecta_suma_superior_al_total_ofertado(self):
        productos = [
            extractor.normalizar_producto(
                {
                    "producto": "Notebook ASUS VivoBook S16",
                    "categoria": "equipo",
                    "marca": "ASUS",
                    "cantidad": 1,
                    "precio_unitario": 4213990,
                    "precio_total": 4213990,
                    "precio_total_tipo": "linea",
                },
                "Austin", "1", "a.pdf", "economico",
            ),
            extractor.normalizar_producto(
                {
                    "producto": "NOTEBOOK 15,6",
                    "categoria": "equipo",
                    "cantidad": 30,
                    "precio_unitario": 807133,
                    "precio_total": 24213990,
                    "precio_total_tipo": "linea",
                },
                "Austin", "1", "b.pdf", "economico",
            ),
        ]
        validados = extractor.validar_productos(productos, productos, "25.125.550")
        self.assertTrue(all(
            "suma_productos_supera_total_oferta" in producto["alertas"]
            for producto in validados
        ))
        self.assertTrue(all(
            producto["estado_validacion"] == "inconsistente"
            for producto in validados
        ))

    def test_detecta_conflictos_entre_fuentes_del_mismo_modelo(self):
        productos = [
            {
                "producto": "Notebook HP EliteBook",
                "categoria": "equipo",
                "marca": "HP",
                "modelo": "840 G10",
                "cantidad": 10,
                "precio_unitario": 900000,
                "precio_total": 9000000,
                "precio_total_tipo": "linea",
                "confianza": "alta",
                "archivo_fuente": "a.pdf",
            },
            {
                "producto": "HP EliteBook 840 G10",
                "categoria": "equipo",
                "marca": "HP",
                "modelo": "840 G10",
                "cantidad": 12,
                "precio_unitario": 950000,
                "precio_total": 11400000,
                "precio_total_tipo": "linea",
                "confianza": "alta",
                "archivo_fuente": "b.pdf",
            },
        ]
        validados = extractor.validar_productos(productos, productos, None)
        self.assertTrue(all(
            producto["estado_validacion"] == "inconsistente"
            for producto in validados
        ))
        self.assertTrue(any(
            "conflicto_entre_fuentes_precio_unitario" in producto["alertas"]
            for producto in validados
        ))

    def test_consolidacion_por_lotes_no_corta_registros(self):
        registros = [
            {"producto": f"Notebook {indice}", "categoria": "equipo", "evidencia": "x" * 80}
            for indice in range(20)
        ]
        lotes = extractor.dividir_registros_json(registros, 500)
        self.assertEqual(sum(len(lote) for lote in lotes), len(registros))
        self.assertTrue(all(lote for lote in lotes))

    def test_deduplicacion_combina_datos_y_fuentes(self):
        tecnico = {
            "producto": "Notebook HP X1",
            "categoria": "equipo",
            "modelo": "X1",
            "cantidad": 2,
            "archivo_fuente": "tecnico.pdf",
        }
        economico = {
            "producto": "Notebook HP X1",
            "categoria": "equipo",
            "modelo": "X1",
            "precio_unitario": 100,
            "precio_total": 200,
            "archivo_fuente": "economico.pdf",
        }
        resultado = extractor.deduplicar_productos([tecnico, economico])
        self.assertEqual(len(resultado), 1)
        self.assertEqual(resultado[0]["cantidad"], 2)
        self.assertEqual(resultado[0]["precio_unitario"], 100)
        self.assertEqual(
            resultado[0]["fuentes_respaldo"],
            ["economico.pdf", "tecnico.pdf"],
        )

    def test_checkpoint_atomico(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "extraccion_ia.json"
            extractor.escribir_json_atomico(ruta, [{"rut": "1"}])
            self.assertEqual(json.loads(ruta.read_text(encoding="utf-8")), [{"rut": "1"}])
            self.assertFalse(ruta.with_suffix(".json.tmp").exists())

    def test_resultado_fallido_se_reprocesa(self):
        fallido = {
            "resultados_archivos": [{"estado_ia": "http_500"}],
            "estado_consolidacion": "omitida",
            "necesita_ocr": [],
        }
        completo = {
            "resultados_archivos": [{"estado_ia": "ok"}],
            "estado_consolidacion": "ok",
            "necesita_ocr": [],
        }
        self.assertTrue(extractor.resultado_requiere_reproceso(fallido))
        self.assertFalse(extractor.resultado_requiere_reproceso(completo))
        completo["modelo"] = "modelo-a"
        completo["backend"] = "lmstudio"
        self.assertTrue(
            extractor.resultado_requiere_reproceso(
                completo, modelo="modelo-b", backend="lmstudio"
            )
        )


if __name__ == "__main__":
    unittest.main()

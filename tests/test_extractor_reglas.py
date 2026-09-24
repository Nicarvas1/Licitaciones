import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("extractor_reglas", ROOT / "3_extraer_reglas.py")
reglas = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(reglas)


class ExtractorReglasTest(unittest.TestCase):
    def test_excluye_productos_fuera_de_alcance_sin_perder_paquetes(self):
        excluidos = (
            "Bolso para Notebook HP 15.6",
            "Microsoft Licencias de Software FQC-10553",
            "SERV17 Masterizacion de PCs y Notebooks",
            "Proyector Epson portatil 3000 lumenes",
            "Scanner Brother ADS-4300N Desktop Professional",
            "FortiSIEM All-In-One Subscription License",
        )
        for descripcion in excluidos:
            with self.subTest(descripcion=descripcion):
                self.assertIsNone(reglas.categoria_reglas({"producto": descripcion}))

        self.assertEqual(
            reglas.categoria_reglas({"producto": "Notebook HP con licencia Office incluida"}),
            "equipo",
        )
        self.assertEqual(
            reglas.categoria_reglas({"producto": "Impresora portatil HP OfficeJet 200"}),
            "impresora",
        )
        self.assertEqual(
            reglas.categoria_reglas({"producto": "Desktop Ryzen 5 con monitor 24 pulgadas"}),
            "equipo",
        )

    def test_detecta_documentos_historicos(self):
        historicos = (
            "RESPALDO_EXPERIENCIA/FACTURA N487.pdf",
            "orden_de_compra_2024.pdf",
            "OC_1057536-2218-SE24.pdf",
            "certificado_distribuidor.pdf",
        )
        for nombre in historicos:
            with self.subTest(nombre=nombre):
                self.assertTrue(reglas.es_documento_historico(nombre))
        self.assertFalse(reglas.es_documento_historico("Oferta_Economica_1057062.pdf"))

    def test_fusiona_anexo_generico_con_cotizacion_detallada(self):
        generico = {
            "producto": "MONITOR MSI 27 FHD",
            "categoria": "monitor",
            "cantidad": 1,
            "precio_unitario": 79950,
            "precio_total": 79950,
            "archivo_fuente": "economico__02__ANEXO_CARTA_OFERTA.pdf",
        }
        detallado = {
            "producto": "27E3H2 Monitor AOC de 27 IPS Full HD 100Hz",
            "categoria": "monitor",
            "marca": "AOC",
            "modelo": "27E3H2",
            "cantidad": 1,
            "precio_unitario": 79950,
            "precio_total": 79950,
            "archivo_fuente": "economico__01__COTIZACION_INTERNA_ALCA.pdf",
        }
        resultado = reglas.fusionar_productos_reglas([generico, detallado])
        self.assertEqual(len(resultado), 1)
        self.assertEqual(resultado[0]["marca"], "AOC")
        self.assertEqual(resultado[0]["modelo"], "27E3H2")
        self.assertEqual(len(resultado[0]["fuentes_respaldo"]), 2)

    def test_no_completa_modelo_si_hay_varios_productos_incompletos(self):
        productos = [
            {"producto": "Notebook tipo 1", "categoria": "equipo", "marca": None, "modelo": None},
            {"producto": "Notebook tipo 2", "categoria": "equipo", "marca": None, "modelo": None},
        ]
        textos = {"ficha.pdf": "Notebook HP ProBook 440 G11"}
        resultado = reglas.completar_desde_otros_documentos(productos, textos)
        self.assertTrue(all(producto.get("modelo") is None for producto in resultado))

    def test_cruce_identidad_conserva_archivo_y_linea(self):
        productos = [
            {"producto": "Notebook tipo 1", "categoria": "equipo", "marca": None, "modelo": None},
        ]
        linea = "Se oferta Notebook HP ProBook 440 G11, 16GB RAM, 512GB SSD"
        resultado = reglas.completar_desde_otros_documentos(
            productos, {"tecnico__01__ficha.pdf": linea}
        )
        self.assertEqual(resultado[0]["marca"], "HP")
        self.assertIn("ProBook", resultado[0]["modelo"])
        self.assertEqual(resultado[0]["fuente_producto"], "tecnico__01__ficha.pdf")
        self.assertEqual(resultado[0]["evidencia_producto"], linea)
        self.assertEqual(resultado[0]["metodo_identidad"], "cruce_unico_proveedor")

    def test_no_cruza_modelo_de_otra_marca(self):
        productos = [
            {"producto": "Impresora Brother HL-1202", "categoria": "impresora", "marca": "Brother", "modelo": None},
        ]
        resultado = reglas.completar_desde_otros_documentos(
            productos, {"oferta.pdf": "Impresora Epson EcoTank L4260"}
        )
        self.assertIsNone(resultado[0]["modelo"])


if __name__ == "__main__":
    unittest.main()
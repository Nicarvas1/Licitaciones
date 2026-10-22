import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("extractor_marca", ROOT / "3_extraer_ia.py")
E = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(E)


class MarcaCompuestaTest(unittest.TestCase):
    def test_varias_marcas_en_un_campo_no_son_una_marca(self):
        for valor in ("Lenovo/HP", "Asus/MSI/Logitech", "AMD/ASUS/Palit/Corsair", "Lenovo y HP", "HP + Dell"):
            self.assertIsNone(E.normalizar_marca(valor), valor)

    def test_generic_sin_sufijo_no_es_marca(self):
        self.assertIsNone(E.normalizar_marca("Generic"))

    def test_marcas_simples_se_conservan(self):
        self.assertEqual(E.normalizar_marca("HP"), "HP")
        self.assertEqual(E.normalizar_marca("OMA-X"), "OMA-X")
        self.assertEqual(E.normalizar_marca("Hewlett Packard"), "HP")




class LecturaDocAntiguoTest(unittest.TestCase):
    def test_doc_antiguo_se_lee_con_antiword(self):
        from unittest import mock
        import subprocess
        salida = subprocess.CompletedProcess([], 0, stdout="OFERTA ECONOMICA\nNotebook 5 unidades $ 500.000".encode("cp1252"))
        with mock.patch.object(E, "encontrar_antiword", return_value="antiword"), \
                mock.patch.object(E.subprocess, "run", return_value=salida):
            texto, estado = E.extraer_archivo(Path("oferta.doc"), type("A", (), {})())
        self.assertEqual(estado, "texto")
        self.assertIn("Notebook 5 unidades", texto)

    def test_doc_sin_antiword_queda_no_soportado(self):
        from unittest import mock
        with mock.patch.object(E, "encontrar_antiword", return_value=None):
            self.assertEqual(E.extraer_doc(Path("oferta.doc")), ("", "no_soportado"))



class OcrYRuidoAdministrativoTest(unittest.TestCase):
    def test_carnet_y_certificados_del_anexo_administrativo_son_ruido(self):
        for nombre in ("administrativo__20__2-carnetPamelaLegalizado.pdf", "administrativo__02__Certificado_de_habilidad.pdf",
                       "administrativo__05__Programa_Integridad_ACT.pdf"):
            self.assertTrue(E.es_ruido_administrativo(nombre), nombre)

    def test_cotizacion_dentro_del_anexo_administrativo_no_es_ruido(self):
        for nombre in ("administrativo__03__COTIZACION_ID_564162-16.pdf", "administrativo__01__ECONOMICO4291-4-LE26.pdf",
                       "administrativo__03__FT-HP_ProOne_240.pdf", "tecnico__01__certificado.pdf"):
            self.assertFalse(E.es_ruido_administrativo(nombre), nombre)

    def test_sin_tesseract_la_comprobacion_falla_con_mensaje(self):
        anterior = E.pytesseract.pytesseract.tesseract_cmd
        try:
            ok, mensaje = E.verificar_ocr(r"C:\no\existe\tesseract.exe")
        finally:
            E.pytesseract.pytesseract.tesseract_cmd = anterior
        self.assertFalse(ok)
        self.assertIn("Tesseract", mensaje)


if __name__ == "__main__":
    unittest.main()

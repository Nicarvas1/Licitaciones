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


if __name__ == "__main__":
    unittest.main()

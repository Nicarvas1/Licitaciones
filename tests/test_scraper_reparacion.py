import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("scraper", ROOT / "2_scraper_ofertas.py")
scraper = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(scraper)


def proveedor(carpeta, rut, enlaces, archivos=()):
    carpeta.mkdir(parents=True)
    (carpeta / "oferta.json").write_text(json.dumps({"rut": rut, "enlaces_anexos": enlaces}), encoding="utf-8")
    for nombre in archivos:
        (carpeta / nombre).write_bytes(b"x")


class ReparacionScraperTest(unittest.TestCase):
    def test_detecta_solo_los_anexos_economicos_o_tecnicos_que_faltan(self):
        with tempfile.TemporaryDirectory() as tmp:
            lic = Path(tmp) / "100-1-LE26"
            proveedor(lic / "1-1__completo", "1-1", {"economico": "u", "tecnico": "u"},
                      ["economico__01__a.pdf", "tecnico__01__b.pdf"])
            proveedor(lic / "2-2__sin_tecnico", "2-2", {"economico": "u", "tecnico": "u"}, ["economico__01__a.pdf"])
            # sin economico ni tecnico: puede haber subido su cotizacion como "administrativo"
            proveedor(lic / "3-3__solo_admin", "3-3", {"administrativo": "u"})
            proveedor(lic / "4-4__sin_anexos", "4-4", {})
            proveedor(lic / "5-5__admin_ya_bajado", "5-5", {"administrativo": "u"}, ["administrativo__01__a.pdf"])
            self.assertEqual(scraper.faltantes_en_disco(lic), {"2-2", "3-3"})

    def test_recontar_archivos_y_ofertas_en_metadata(self):
        with tempfile.TemporaryDirectory() as tmp:
            lic = Path(tmp) / "100-1-LE26"
            proveedor(lic / "1-1__a", "1-1", {"economico": "u"}, ["economico__01__a.pdf"])
            proveedor(lic / "2-2__b", "2-2", {"economico": "u"}, ["economico__01__a.pdf", "economico__02__b.xlsx"])
            (lic / "metadata.json").write_text(json.dumps({"codigo": "100-1-LE26", "ofertas": 0, "archivos": 0}))
            self.assertEqual(scraper.recontar_archivos(lic), 3)
            datos = json.loads((lic / "metadata.json").read_text())
            self.assertEqual((datos["ofertas"], datos["archivos"]), (2, 3))


if __name__ == "__main__":
    unittest.main()

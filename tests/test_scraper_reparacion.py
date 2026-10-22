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



class PaginacionDelPopupTest(unittest.TestCase):
    def test_detecta_las_paginas_del_popup_de_anexos(self):
        html = ("<a href=\"javascript:__doPostBack('DWNL$grdId','Page$2')\">2</a>"
                "<a href=\"javascript:__doPostBack('DWNL$grdId','Page$3')\">3</a>")
        self.assertEqual(sorted(int(n) for n in scraper.PATRON_PAGINA_POPUP.findall(html)), [2, 3])

    def test_anexo_con_seis_archivos_se_considera_posiblemente_truncado(self):
        with tempfile.TemporaryDirectory() as tmp:
            lic = Path(tmp) / "100-1-LE26"
            seis = [f"tecnico__{i:02d}__a.pdf" for i in range(1, 7)]
            proveedor(lic / "1-1__truncado", "1-1", {"tecnico": "u"}, seis)
            proveedor(lic / "2-2__completo", "2-2", {"tecnico": "u"}, seis[:4])
            proveedor(lic / "3-3__doce", "3-3", {"tecnico": "u"}, [f"tecnico__{i:02d}__a.pdf" for i in range(1, 13)])
            self.assertEqual(scraper.faltantes_en_disco(lic), {"1-1", "3-3"})

    def test_archivo_ya_bajado_se_reconoce_por_prefijo_e_indice(self):
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = Path(tmp)
            (carpeta / "tecnico__03__ficha (1).pdf").write_bytes(b"x")
            self.assertTrue(scraper.archivo_ya_bajado(carpeta, "tecnico", 3))
            self.assertFalse(scraper.archivo_ya_bajado(carpeta, "tecnico", 4))
            self.assertFalse(scraper.archivo_ya_bajado(carpeta, "economico", 3))



class HuecosEnLaNumeracionTest(unittest.TestCase):
    def test_un_hueco_en_los_indices_marca_al_proveedor_como_incompleto(self):
        with tempfile.TemporaryDirectory() as tmp:
            lic = Path(tmp) / "100-1-LE26"
            proveedor(lic / "1-1__con_hueco", "1-1", {"tecnico": "u"},
                      ["tecnico__01__a.pdf", "tecnico__02__b.pdf", "tecnico__04__d.pdf"])
            proveedor(lic / "2-2__sin_hueco", "2-2", {"tecnico": "u"},
                      ["tecnico__01__a.pdf", "tecnico__02__b.pdf", "tecnico__03__c.pdf"])
            self.assertEqual(scraper.faltantes_en_disco(lic), {"1-1"})



class LicitacionSinDatosTest(unittest.TestCase):
    def test_metadata_con_cero_ofertas_cuenta_como_no_descargada(self):
        with tempfile.TemporaryDirectory() as tmp:
            lic = Path(tmp) / "100-1-LE26"
            lic.mkdir()
            self.assertTrue(scraper.licitacion_sin_datos(lic))  # sin metadata
            (lic / "metadata.json").write_text(json.dumps({"ofertas": 0, "archivos": 0}), encoding="utf-8")
            self.assertTrue(scraper.licitacion_sin_datos(lic))
            (lic / "metadata.json").write_text(json.dumps({"ofertas": 7, "archivos": 30}), encoding="utf-8")
            self.assertFalse(scraper.licitacion_sin_datos(lic))



class FiltroAdministrativosTest(unittest.TestCase):
    OMITIR = ["8-ANEXO_N_1_IDENTIFICACIN_DEL_OFERENTE.pdf", "04.-_Ministerio_de_Hacienda_1618-257-SE25.pdf",
              "44.-_Servicio_Local_de_Educacion_Publica_Licancabur_1233623-66-SE.pdf",
              "Anexo_N__7_DECLARACION_JURADA_NO_PRESENTACION.pdf", "CERTIFICADO_DE_VIGENCIA_DE_LA_SOCIEDAD.pdf",
              "CARPETA_TRIBUTARIA_ALCA.pdf", "10-PROGRAMA_DE_INTEGRIDAD_Todotablet.pdf", "2022_Cedula_CM.pdf",
              "Anexo_05__Experiencia_del_oferente.pdf", "factura_electronica_5810.pdf", "OC.zip", "10.-_C.I_ALEJANDRA.pdf"]
    CONSERVAR = ["Anexo_03__Oferta_Economica_y_plazos.pdf", "Resumen_propuesta_SYNERGY.pdf", "COTIZACION_ID_564162-16-L126.pdf",
                 "FT-Impresora_EPSON_Multifuncional_EcoTank_L5590.pdf", "ANEXO_N_4_OFERTA_ECONOMICA.pdf", "PPTO.pdf",
                 "Ficha_tecnica_Brother_MFC-L6915DW.pdf", "anexo_4_(44).pdf", "Documentos_Empresa.zip",
                 "2-ANEXO_N__4_ESPECIFICACIONES_TECNICAS_-_OFERTA_ECONOMICA.docx"]

    def test_omite_declaraciones_certificados_y_ordenes_de_compra_pasadas(self):
        for nombre in self.OMITIR:
            self.assertTrue(scraper.es_administrativo_sin_oferta(nombre), nombre)

    def test_conserva_ofertas_fichas_y_nombres_dudosos(self):
        for nombre in self.CONSERVAR:
            self.assertFalse(scraper.es_administrativo_sin_oferta(nombre), nombre)


if __name__ == "__main__":
    unittest.main()

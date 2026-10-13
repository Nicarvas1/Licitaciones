"""Pruebas de las reglas deterministas que protegen la calidad del Excel: marcas, IVA, lineas
distintas, precios sin respaldo, tipo de producto y filtro de anexos tecnicos."""
import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("extractor_ia", ROOT / "3_extraer_ia.py")
extractor = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(extractor)


def producto(**campos):
    base = {"categoria": "equipo", "precio_total_tipo": "linea", "confianza": "alta", "archivo_fuente": "a.pdf"}
    base.update(campos)
    return base


class ReconciliacionSegurasTest(unittest.TestCase):
    def test_un_precio_que_esta_en_el_texto_del_anexo_no_se_anula(self):
        bloques = [{"archivo": "economico__01.pdf", "pagina": 1, "texto": chr(10).join([
            "FILA 2: Notebook HP 250 || 30 || 803.243 || 24.097.290",
            "Nota: oferta Notebook Lenovo ThinkPad E16 precio 854.700 por equipo"])}]
        productos = [producto(producto="Notebook Lenovo ThinkPad E16", marca="Lenovo", cantidad=30,
                              precio_unitario=854700, precio_total=25641000, fuente_precio="economico__01.pdf")]
        resultado = extractor.reconciliar_precios_con_anexos(productos, bloques, None, anular=True)
        self.assertEqual(resultado[0]["precio_unitario"], 854700)
        self.assertFalse(resultado[0].get("precio_fuente_no_coincide"))

    def test_un_precio_ausente_del_anexo_si_se_anula(self):
        bloques = [{"archivo": "economico__01.pdf", "pagina": 1, "texto":
                    "FILA 2: Notebook HP 250 || 30 || 803.243 || 24.097.290"}]
        productos = [producto(producto="Notebook Lenovo ThinkPad E16", marca="Lenovo", cantidad=30,
                              precio_unitario=999999, precio_total=29999970, fuente_precio="economico__01.pdf")]
        resultado = extractor.reconciliar_precios_con_anexos(productos, bloques, None, anular=True)
        self.assertIsNone(resultado[0]["precio_unitario"])
        self.assertTrue(resultado[0]["precio_fuente_no_coincide"])


class FilasConColumnasVaciasTest(unittest.TestCase):
    def test_fila_con_columnas_vacias_entre_unitario_y_total(self):
        bloques = [{"archivo": "economico__01.pdf", "pagina": 1, "texto": chr(10).join([
            "FILA 4: 3 || MONITOR 32\" || 214.555.- ||  ||  || 643.665.-",
            "FILA 5: 2 || COMPUTADOR DE ESCRITORIO (EQUIPO ARMADO) || 704.425 ||  ||  || 1.408.850",
            "FILA 9:  ||  || NETO ||  ||  || 7.998.477",
            "FILA 10:  ||  || IVA ||  ||  || 1.519.711"])}]
        filas = extractor.filas_precio_economico(bloques)
        self.assertEqual([(f["cantidad"], f["precio_unitario"], f["precio_total"]) for f in filas],
                         [(3, 214555, 643665), (2, 704425, 1408850)])

    def test_monto_con_sufijo_punto_guion(self):
        self.assertEqual(extractor.normalizar_numero("214.555.-"), 214555)
        self.assertEqual(extractor.normalizar_numero("$ 1.409.-"), 1409)
        self.assertEqual(extractor.normalizar_numero("471.111"), 471111)


class FichasTecnicasTest(unittest.TestCase):
    def test_la_ficha_aporta_marca_a_la_fila_con_precio_y_se_descarta(self):
        fila = producto(producto="LAPTOP PROBOOK 450 G10 I5-1334U 16GB B7KF9AT", cantidad=6, precio_unitario=593000,
                        precio_total=3558000, archivo_fuente="economico__01.pdf")
        ficha = producto(producto="HP ProBook 450 G10 B7KF9AT", marca="HP", modelo="B7KF9AT",
                         archivo_fuente="tecnico__01__ficha.pdf")
        validados = extractor.validar_productos([fila, ficha], [fila, ficha], None)
        self.assertEqual(len(validados), 1)
        self.assertEqual(validados[0]["marca"], "HP")
        self.assertIn("tecnico__01__ficha.pdf", validados[0]["fuentes_respaldo"])

    def test_ficha_huerfana_se_descarta_si_hay_productos_con_precio(self):
        fila = producto(producto="Notebook Tipo 1", marca="HP", cantidad=6, precio_unitario=567760,
                        precio_total=3406560, archivo_fuente="economico__01.pdf")
        huerfana = producto(producto="ThinkPad L14 Gen 6 AMD", marca="Lenovo", archivo_fuente="tecnico__02__ficha.pdf")
        self.assertEqual(len(extractor.validar_productos([fila, huerfana], [fila, huerfana], None)), 1)

    def test_si_no_hay_precios_las_fichas_se_conservan(self):
        ficha = producto(producto="HP ProBook 450 G10 B7KF9AT", marca="HP", archivo_fuente="tecnico__01__ficha.pdf")
        self.assertEqual(len(extractor.validar_productos([ficha], [ficha], None)), 1)

    def test_un_producto_economico_sin_precio_no_se_descarta(self):
        con_precio = producto(producto="Notebook HP 440 G10", marca="HP", cantidad=2, precio_unitario=786468,
                              precio_total=1572936, archivo_fuente="economico__01.pdf")
        sin_precio = producto(producto="Notebook HP 450 G10", marca="HP", archivo_fuente="economico__01.pdf")
        self.assertEqual(len(extractor.validar_productos([con_precio, sin_precio], [con_precio, sin_precio], None)), 2)


class LineasNoOfertadasTest(unittest.TestCase):
    def test_lineas_tipo_sin_precio_en_un_anexo_con_otras_lineas_con_precio_no_se_ofertaron(self):
        con = producto(producto="Notebook Tipo 1", cantidad=6, precio_unitario=593000, precio_total=3558000,
                       archivo_fuente="economico__02__anexo6.pdf")
        sin = producto(producto="Notebook Tipo 2", cantidad=10, archivo_fuente="economico__02__anexo6.pdf")
        self.assertEqual(len(extractor.validar_productos([con, sin], [con, sin], None)), 1)

    def test_si_ninguna_linea_tiene_precio_se_conservan(self):
        sin = producto(producto="Notebook Tipo 2", cantidad=10, archivo_fuente="economico__02__anexo6.pdf")
        self.assertEqual(len(extractor.validar_productos([sin], [sin], None)), 1)

    def test_marca_generica_no_es_una_marca(self):
        for texto in ("Marca Gen\u00e9rica", "Generico", "sin marca", "N/A"):
            self.assertIsNone(extractor.normalizar_marca(texto), texto)
        self.assertEqual(extractor.normalizar_marca("Lenovo"), "Lenovo")


class LineaEnDosFormulariosTest(unittest.TestCase):
    def test_requisito_y_oferta_con_los_mismos_numeros_en_archivos_distintos_son_una_linea(self):
        requisito = producto(producto="COMPUTADOR: PROCESADOR INTEL i5 3.0 GHZ 11 GEN, 8 GB RAM, SSD 500 GB",
                             cantidad=22, precio_unitario=450120, precio_total=9902640,
                             fuente_precio="F7_presupuesto.pdf", tipo_equipo="desktop")
        oferta = producto(producto="HP PRO MINI 260 G9 DESKTOP INTEL CORE I5-1334U", marca="HP", cantidad=22,
                          precio_unitario=450120, precio_total=9902640, fuente_precio="cotizacion.pdf",
                          tipo_equipo="desktop")
        validados = extractor.validar_productos([requisito, oferta], [requisito, oferta], None)
        self.assertEqual(len(validados), 1)
        self.assertEqual(validados[0]["marca"], "HP")

    def test_el_total_de_la_oferta_como_total_de_un_producto_se_marca(self):
        resumen = producto(producto="COMPUTADORES", cantidad=23, precio_unitario=638565.22, precio_total=14687000)
        otros = [producto(producto=f"Monitor LG 32U631A-B {i}", marca="LG", cantidad=1, precio_unitario=100000 + i,
                          precio_total=100000 + i) for i in range(3)]
        validados = extractor.validar_productos([resumen] + otros, [resumen] + otros, "$ 14.687.000")
        # una linea resumen generica cuyo total es el de la oferta no es un producto: se descarta
        self.assertEqual(len(validados), 3)
        self.assertNotIn("COMPUTADORES", [p["producto"] for p in validados])

    def test_un_producto_especifico_con_el_total_de_la_oferta_se_marca(self):
        raro = producto(producto="Notebook HP ProBook 440 G10", marca="HP", cantidad=23, precio_unitario=638565.22,
                        precio_total=14687000)
        otros = [producto(producto=f"Monitor LG 32U631A-B {i}", marca="LG", cantidad=1, precio_unitario=100000 + i,
                          precio_total=100000 + i) for i in range(3)]
        validados = extractor.validar_productos([raro] + otros, [raro] + otros, "$ 14.687.000")
        self.assertIn("total_general_asignado_a_producto", validados[0]["alertas"])

    def test_lineas_resumen_en_plural_no_se_ofertaron_si_no_tienen_precio(self):
        con = producto(producto="Monitor LG 32U631A-B", marca="LG", cantidad=5, precio_unitario=170000,
                       precio_total=850000, archivo_fuente="economico__02.pdf")
        resumen = producto(producto="MONITORES", cantidad=37, archivo_fuente="economico__02.pdf")
        self.assertEqual(len(extractor.validar_productos([con, resumen], [con, resumen], None)), 1)


class LineasGenericasTest(unittest.TestCase):
    def test_tipo_1_del_cuadro_de_ofertas_se_funde_con_la_cotizacion_real(self):
        generica = producto(producto="Notebook Tipo 1", cantidad=6, precio_unitario=593000, precio_total=3558000,
                            fuente_precio="anexo6.pdf")
        real = producto(producto="LAPTOP HP PROBOOK 450 G10 I5-1334U / 16GB RAM", marca="HP", cantidad=6,
                        precio_unitario=593000, precio_total=3558000, fuente_precio="cotizacion.pdf")
        validados = extractor.validar_productos([generica, real], [generica, real], None)
        self.assertEqual(len(validados), 1)
        self.assertIn("PROBOOK", validados[0]["producto"])

    def test_dos_lineas_genericas_distintas_no_se_funden(self):
        a = producto(producto="Notebook Tipo 1", cantidad=6, precio_unitario=567760, precio_total=3406560,
                     fuente_precio="anexo6.pdf")
        b = producto(producto="Notebook Tipo 2", cantidad=10, precio_unitario=815920, precio_total=8159200,
                     fuente_precio="anexo6.pdf")
        self.assertEqual(len(extractor.validar_productos([a, b], [a, b], None)), 2)

    def test_anexo_tecnico_con_cantidad_se_funde_con_la_fila_de_igual_tipo_y_cantidad(self):
        fila = producto(producto="ALL IN ONE INTEL CORE I5, 8GB RAM", cantidad=10, precio_unitario=675620,
                        precio_total=6756200, archivo_fuente="economico__01.pdf")
        tecnica = producto(producto="Computador All-in-One con procesador Intel Core i5", marca="Lenovo", cantidad=10,
                           archivo_fuente="tecnico__01__anexo.pdf", tipo_equipo="all-in-one")
        fila["tipo_equipo"] = "all-in-one"
        validados = extractor.validar_productos([fila, tecnica], [fila, tecnica], None)
        self.assertEqual(len(validados), 1)
        self.assertEqual(validados[0]["marca"], "Lenovo")


class MarcasTest(unittest.TestCase):
    def test_marca_nombrada_en_la_descripcion(self):
        self.assertEqual(extractor.marca_desde_descripcion("Impresora Multifuncional Epson EcoTank L6270"), "Epson")
        self.assertEqual(extractor.marca_desde_descripcion("LAPTOP HP PROBOOK 450 G10 I5"), "HP")
        self.assertEqual(extractor.marca_desde_descripcion("Notebook ThinkPad E16 Ryzen 7"), "Lenovo")

    def test_no_adivina_si_hay_dos_marcas_o_es_un_requisito(self):
        self.assertIsNone(extractor.marca_desde_descripcion("Notebook HP o Lenovo con 16GB"))
        self.assertIsNone(extractor.marca_desde_descripcion("NOTEBOOK, EQUIVALENTE A HP 14-EP0130LA"))
        self.assertIsNone(extractor.marca_desde_descripcion("Notebook Tipo 1"))

    def test_marca_unica_entre_las_fichas_tecnicas_del_mismo_tipo(self):
        genericos = [producto(producto="Notebook Tipo 1"), producto(producto="All in One Tipo 1")]
        parciales = [
            producto(producto="HP ProBook 440 G10", marca="HP", archivo_fuente="tecnico__01__ficha.pdf"),
            producto(producto="HP ProOne 245 G10 All-in-One", marca="HP", archivo_fuente="tecnico__02__ficha.pdf"),
        ]
        resultado = extractor.completar_marcas_por_descripcion_y_tipo(genericos, parciales)
        self.assertEqual([p["marca"] for p in resultado], ["HP", "HP"])
        self.assertTrue(all(p["metodo_marca"] == "tipo_con_marca_unica" for p in resultado))

    def test_dos_marcas_en_el_mismo_tipo_no_se_asignan(self):
        genericos = [producto(producto="Notebook Tipo 1")]
        parciales = [
            producto(producto="HP ProBook 440 G10", marca="HP", archivo_fuente="tecnico__01__a.pdf"),
            producto(producto="Lenovo ThinkPad L14 notebook", marca="Lenovo", archivo_fuente="tecnico__02__b.pdf"),
        ]
        resultado = extractor.completar_marcas_por_descripcion_y_tipo(genericos, parciales)
        self.assertIsNone(resultado[0].get("marca"))

    def test_marca_por_nombre_de_ficha_tecnica(self):
        genericos = [producto(producto="All in One Tipo 1")]
        archivos = ["economico__01__oferta.pdf", "tecnico__03__HP_ProOne_245_G10_AIO.pdf",
                    "tecnico__04__CERTIFICADO_DISTRIBUIDOR_DELL.pdf"]
        resultado = extractor.completar_marcas_por_nombre_de_archivo(genericos, archivos)
        self.assertEqual(resultado[0]["marca"], "HP")
        self.assertEqual(resultado[0]["metodo_marca"], "nombre_de_ficha_tecnica")

    def test_la_familia_corrige_la_marca_leida_del_renglon_vecino(self):
        corregido = extractor.corregir_marca_por_familia(
            [producto(producto="ASUS Notebook Gamer LOQ Gen 9 Intel Core i5", marca="ASUS")])
        self.assertEqual(corregido[0]["marca"], "Lenovo")
        self.assertEqual(corregido[0]["marca_declarada_por_modelo"], "ASUS")
        intacto = extractor.corregir_marca_por_familia([producto(producto="HP ProBook 440 G10", marca="HP")])
        self.assertEqual(intacto[0]["marca"], "HP")


class IvaYLineasTest(unittest.TestCase):
    def test_mismo_producto_en_dos_archivos_con_distinta_numeracion_se_fusiona(self):
        a = producto(item="1", producto="Notebook HP 15-FC0022LA AMD Ryzen 7 7730U 16GB", marca="HP", cantidad=2,
                     precio_unitario=549706, precio_total=1099412, fuente_precio="a.pdf")
        b = producto(item="5", producto="Notebook HP 15-FC0022LA AMD Ryzen 7 7730U 16GB", marca="HP", cantidad=2,
                     precio_unitario=549706, precio_total=1099412, fuente_precio="b.pdf")
        self.assertEqual(len(extractor.validar_productos([a, b], [a, b], None)), 1)

    def test_dos_items_del_mismo_archivo_no_se_fusionan(self):
        a = producto(item="4", producto="Notebook HP 14-EP0130LA i3 8GB", marca="HP", cantidad=1,
                     precio_unitario=402000, precio_total=402000, fuente_precio="a.pdf")
        b = producto(item="6", producto="Notebook HP 14-EP0130LA i3 8GB", marca="HP", cantidad=1,
                     precio_unitario=402000, precio_total=402000, fuente_precio="a.pdf")
        self.assertEqual(len(extractor.validar_productos([a, b], [a, b], None)), 2)

    def test_la_misma_linea_con_total_neto_y_con_iva_se_fusiona(self):
        neto = producto(producto="Apple MacBook Air 13 M1", marca="Apple", cantidad=2, precio_unitario=967455,
                        precio_total=1934910, archivo_fuente="cotizacion.pdf")
        con_iva = producto(producto="Apple MacBook Air 13 M1", marca="Apple", cantidad=2, precio_unitario=967455,
                           precio_total=2302542.9, archivo_fuente="anexos.pdf")
        validados = extractor.validar_productos([neto, con_iva], [neto, con_iva], None)
        self.assertEqual(len(validados), 1)
        self.assertEqual(validados[0]["precio_total"], 1934910)
        self.assertEqual(validados[0]["estado_validacion"], "ok")

    def test_total_con_iva_se_convierte_a_neto(self):
        filas = extractor.aplicar_logica_iva([producto(producto="AIO HP", cantidad=1, precio_unitario=580000,
                                                        precio_total=690200)])
        self.assertEqual(filas[0]["precio_total"], 580000)
        self.assertTrue(filas[0]["total_con_iva_convertido"])

    def test_duplicado_con_iva_se_descarta_y_deja_la_fuente_de_respaldo(self):
        neto = producto(producto="LAPTOP HP 245 G10", cantidad=4, precio_unitario=346000, precio_total=1384000,
                        fuente_precio="cotizacion.pdf")
        con_iva = producto(producto="NOTEBOOK, EQUIVALENTE A HP 14-EP0130LA", cantidad=4, precio_unitario=411740,
                           precio_total=1646960, fuente_precio="formulario.pdf")
        resultado = extractor.aplicar_logica_iva([neto, con_iva])
        self.assertEqual(len(resultado), 1)
        self.assertEqual(resultado[0]["precio_unitario"], 346000)
        self.assertIn("formulario.pdf", resultado[0]["fuentes_respaldo"])

    def test_misma_cantidad_distinta_no_se_descarta_como_iva(self):
        a = producto(producto="Notebook", cantidad=2, precio_unitario=500000, precio_total=1000000)
        b = producto(producto="Notebook", cantidad=3, precio_unitario=595000, precio_total=1785000)
        self.assertEqual(len(extractor.aplicar_logica_iva([a, b])), 2)

    def test_mismo_equipo_en_dos_items_del_mismo_anexo_no_es_conflicto(self):
        base = dict(producto="LAPTOP HP PROBOOK 450 G10 I5 16GB", modelo="ProBook 450 G10", marca="HP",
                    fuente_precio="cotizacion.pdf", precio_unitario=599000)
        a = producto(cantidad=2, precio_total=1198000, **base)
        b = producto(cantidad=3, precio_total=1797000, **base)
        validados = extractor.validar_productos([a, b], [a, b], None)
        self.assertFalse(any("conflicto" in alerta for p in validados for alerta in p["alertas"]))

    def test_el_exceso_sobre_el_total_solo_recae_en_los_sospechosos(self):
        bueno = producto(producto="Notebook HP ProBook 440 G10", marca="HP", cantidad=2, precio_unitario=500000,
                         precio_total=1000000)
        malo = producto(producto="Notebook Lenovo ThinkPad L14", marca="Lenovo", cantidad=1, precio_unitario=5,
                        precio_total=30_000_000)
        validados = extractor.validar_productos([bueno, malo], [bueno, malo], "$ 2.000.000")
        por_nombre = {p["producto"]: p for p in validados}
        self.assertNotIn("suma_productos_supera_total_oferta", por_nombre["Notebook HP ProBook 440 G10"]["alertas"])
        self.assertEqual(por_nombre["Notebook Lenovo ThinkPad L14"]["estado_validacion"], "inconsistente")


class TotalOfertaTest(unittest.TestCase):
    def test_un_total_de_oferta_corrupto_no_marca_inconsistentes_a_los_productos(self):
        buenos = [producto(producto="Notebook HP ProBook 440 G10", marca="HP", cantidad=6, precio_unitario=618078,
                           precio_total=3708468),
                  producto(producto="All in One HP 205 G8", marca="HP", cantidad=10, precio_unitario=579088,
                           precio_total=5790880)]
        validados = extractor.validar_productos(buenos, buenos, "$ 4")
        self.assertTrue(all(p["estado_validacion"] == "ok" for p in validados), [p["alertas"] for p in validados])
        self.assertIsNone(extractor.total_oferta_creible("$ 4"))
        self.assertEqual(extractor.total_oferta_creible("$ 24.836.767"), 24836767)


class UnitarioLeidoComoTotalTest(unittest.TestCase):
    def test_formulario_con_una_columna_de_precio(self):
        mal = producto(producto="Lenovo ThinkStation P8", marca="Lenovo", cantidad=6,
                       precio_unitario=2023301.5, precio_total=12139809)
        validado = extractor.validar_productos([mal], [mal], "$ 72.838.854")[0]
        self.assertEqual((validado["precio_unitario"], validado["precio_total"]), (12139809, 72838854))
        self.assertTrue(validado["total_leido_como_unitario"])

    def test_si_el_total_ya_es_el_de_la_linea_no_se_toca(self):
        bien = producto(producto="Computador", marca="HP", cantidad=6, precio_unitario=6450429.5,
                        precio_total=38702577)
        validado = extractor.validar_productos([bien], [bien], "$ 38.702.577")[0]
        self.assertEqual(validado["precio_total"], 38702577)
        self.assertFalse(validado.get("total_leido_como_unitario"))


class SumaSobreElTotalTest(unittest.TestCase):
    def test_el_precio_de_una_ficha_web_explica_el_exceso_y_no_se_castiga_a_las_filas_del_anexo(self):
        a = producto(producto="Notebook Tipo 1", marca="HP", cantidad=6, precio_unitario=567760, precio_total=3406560,
                     fuente_precio="economico__01__anexo6.pdf")
        b = producto(producto="All in One Tipo 1", marca="HP", cantidad=10, precio_unitario=714400,
                     precio_total=7144000, fuente_precio="economico__01__anexo6.pdf")
        web = producto(producto="HP ProOne 245 G10 AIO", marca="HP", cantidad=1, precio_unitario=929990,
                       precio_total=929990, fuente_precio="tecnico__03__ficha_web.pdf")
        validados = extractor.validar_productos([a, b, web], [a, b, web], "$ 10.550.560")
        por = {p["producto"]: p for p in validados}
        self.assertEqual(por["Notebook Tipo 1"]["estado_validacion"], "ok")
        self.assertEqual(por["HP ProOne 245 G10 AIO"]["estado_validacion"], "inconsistente")


class TotalGeneralTest(unittest.TestCase):
    def test_varios_productos_con_el_total_de_la_oferta_como_precio_son_inconsistentes(self):
        a = producto(producto="Notebook OZXEN NB15 S2154", marca="OZXEN", cantidad=1, precio_unitario=8306320,
                     precio_total=8306320)
        b = producto(producto="Computador All-in-One OZXEN AL24 SA45C", marca="OZXEN", cantidad=1,
                     precio_unitario=8306320, precio_total=8306320)
        validados = extractor.validar_productos([a, b], [a, b], "$ 8.306.320")
        self.assertTrue(all("total_general_asignado_a_producto" in p["alertas"] for p in validados))
        self.assertTrue(all(p["estado_validacion"] == "inconsistente" for p in validados))

    def test_una_oferta_de_un_solo_producto_puede_costar_el_total(self):
        unico = producto(producto="Notebook HP ProBook 440 G10", marca="HP", cantidad=1, precio_unitario=900000,
                         precio_total=900000)
        validado = extractor.validar_productos([unico], [unico], "$ 900.000")[0]
        self.assertEqual(validado["estado_validacion"], "ok")


class PreciosSinRespaldoTest(unittest.TestCase):
    def test_precio_que_no_esta_en_ningun_documento_se_marca(self):
        bloques = [{"archivo": "economico__01.pdf", "pagina": 1,
                    "texto": "FILA 1: 1 || Notebook HP 250 || 6 || 567.760 || 3.406.560"}]
        numeros = extractor.numeros_del_texto(bloques[0]["texto"])
        real = producto(producto="Notebook HP 250", cantidad=6, precio_unitario=567760, precio_total=3406560)
        inventado = producto(producto="Notebook Lenovo L14", cantidad=3, precio_unitario=500000, precio_total=1500000)
        marcados = extractor.marcar_precios_sin_respaldo([real, inventado], numeros, bloques)
        self.assertFalse(marcados[0].get("precio_sin_respaldo"))
        self.assertTrue(marcados[1].get("precio_sin_respaldo"))
        validado = extractor.validar_productos([inventado], [inventado], None)[0]
        self.assertIn("precio_no_aparece_en_documentos", validado["alertas"])
        self.assertEqual(validado["estado_validacion"], "inconsistente")

    def test_reconciliacion_corrige_cifras_partidas_aunque_la_cantidad_leida_difiera(self):
        bloques = [{"archivo": "economico__01.pdf", "pagina": 1, "texto": chr(10).join([
            "1   Notebook HP 15-FC0043LA                 3 89.318 6    2 .335.908",
            "2   All in one HP ProOne 245 G10            5 31.755 10   5 .317.550"])}]
        mal_leidos = [
            producto(producto="Notebook HP 15-FC0043LA", modelo="15-FC0043LA", cantidad=3, precio_unitario=89318,
                     precio_total=267954, fuente_precio="economico__01.pdf"),
            producto(producto="All in one HP ProOne 245 G10", cantidad=5, precio_unitario=31755,
                     precio_total=158775, fuente_precio="economico__01.pdf"),
        ]
        corregidos = extractor.reconciliar_precios_con_anexos(mal_leidos, bloques, None, anular=False)
        self.assertEqual([(p["cantidad"], p["precio_unitario"], p["precio_total"]) for p in corregidos],
                         [(6, 389318, 2335908), (10, 531755, 5317550)])


    def test_un_producto_mal_leido_no_roba_las_filas_de_los_demas(self):
        bloques = [{"archivo": "economico__01.pdf", "pagina": 1, "texto": chr(10).join([
            "1    ALL IN ONE INTEL CORE I5, 8GB RAM, 512GB SSD, W11P 6 75.620 1 0 6 .756.200",
            "2    ALL IN ONE INTEL CORE I7, 32GB RAM, 1TB SSD, W11P 1 .250.410 2 2 .500.820",
            "7    NOTEBOOK INTEL CORE I5, 16GB RAM, 512GB SSD, W11P 6 05.032 5 3 .025.160"])}]
        base = dict(fuente_precio="economico__01.pdf", categoria="equipo")
        productos = [
            producto(producto="ALL IN ONE INTEL CORE I5, 8GB RAM, 512GB SSD, W11P", cantidad=10,
                     precio_unitario=75620, precio_total=756200, **base),
            producto(producto="ALL IN ONE INTEL CORE I7, 32GB RAM, 1TB SSD, W11P", cantidad=2,
                     precio_unitario=1250410, precio_total=2500820, **base),
            producto(producto="NOTEBOOK INTEL CORE I5, 16GB RAM, 512GB SSD, W11P", cantidad=5,
                     precio_unitario=605032, precio_total=3025160, **base),
        ]
        corregidos = extractor.reconciliar_precios_con_anexos(productos, bloques, None, anular=False)
        self.assertEqual([(p["cantidad"], p["precio_unitario"], p["precio_total"]) for p in corregidos],
                         [(10, 675620, 6756200), (2, 1250410, 2500820), (5, 605032, 3025160)])


class HistorialTest(unittest.TestCase):
    def test_ordenes_de_compra_pasadas_se_reconocen_por_ruta_o_contenido(self):
        for ruta in ("tecnico__01__Experiencia.pdf", "_extraidos/tecnico__01__OCs_Varias/OC 123.pdf",
                     "tecnico__01__1232837-74-SE26.pdf", "tecnico__03__4388-207-AG25.pdf"):
            self.assertTrue(extractor.documento_de_historial(ruta, "texto"), ruta)
        self.assertTrue(extractor.documento_de_historial(
            "tecnico__09__doc.pdf", "ORDEN DE COMPRA" + chr(10) + "NOMBRE ORDEN DE COMPRA : ADQUISICION DE EQUIPOS"))

    def test_los_anexos_de_la_oferta_actual_no_son_historial(self):
        for ruta in ("economico__01__2464-17-LE26.pdf", "tecnico__01__Ficha_Tecnica_HP.pdf",
                     "economico__02__ANEXO_N_6.pdf", "tecnico__02__POLIZA.pdf"):
            self.assertFalse(extractor.documento_de_historial(ruta, "Oferta economica Notebook HP"), ruta)


class AlcanceAdicionalTest(unittest.TestCase):
    def test_escaner_con_marca_delante_y_lineas_de_armado_quedan_fuera(self):
        for descripcion in ("Brother Esc\u00e1ner ADS-3100 con alimentador automatico", "ARMADO COMPUTADOR SIN SISTEMA OPERATIVO"):
            self.assertIsNone(extractor.clasificar_producto({"producto": descripcion, "categoria": "equipo"}), descripcion)
        self.assertEqual(extractor.clasificar_producto(
            {"producto": "Impresora multifuncional con esc\u00e1ner Epson L5590", "categoria": "impresora"}), "impresora")

    def test_unitario_imposible_para_el_tipo_de_equipo(self):
        aio = producto(producto="All in One V400 Intel Core 5 23.8", marca="ASUS", cantidad=1, tipo_equipo="all-in-one",
                       precio_unitario=7857050, precio_total=7857050)
        validado = extractor.validar_productos([aio], [aio], None)[0]
        self.assertIn("precio_unitario_fuera_de_rango", validado["alertas"])


class TipoYFiltroTest(unittest.TestCase):
    def test_tipo_declarado_por_el_modelo_y_alias(self):
        self.assertEqual(extractor.subcategoria_producto(
            producto(producto="Asus B1503CVA-S75898X 15.6in", tipo_equipo="laptop")), "notebook")
        self.assertEqual(extractor.subcategoria_producto(
            producto(producto="TODO-EN-UNO HP PRO ONE 245 G10")), "all-in-one")

    def test_filtro_omite_certificados_y_conserva_fichas_y_anexos_con_equipos(self):
        self.assertTrue(extractor.anexo_tecnico_sin_equipos("Certificado ISO 9001 de gestion de calidad " * 5,
                                                              "tecnico__01__Dell_ISO_9001.pdf"))
        self.assertFalse(extractor.anexo_tecnico_sin_equipos("HP 250 G10 Notebook PC procesador i7 16GB RAM " * 5,
                                                               "tecnico__01__ficha.pdf"))
        historial = chr(10).join(["[PAGINA 1]", "FILA 2: 43211503 Notebook 3 Unidad Notebook HP 14 670.000,00 0,00"])
        self.assertTrue(extractor.anexo_tecnico_sin_equipos(historial, "tecnico__02__COMPROBANTE_DE_EXPERIENCIA.pdf"))
        ficha_thinkstation = "ThinkStation P5 Gen 2 30KBCTO1WWES1 " * 20
        self.assertFalse(extractor.anexo_tecnico_sin_equipos(ficha_thinkstation, "tecnico__01__FICHA_TECNICA.pdf"))
        self.assertTrue(extractor.anexo_tecnico_sin_equipos("Texto administrativo sin ningun producto " * 20,
                                                              "tecnico__01__bases.pdf"))
        garantias = chr(10).join(["[PAGINA 1]", "FILA 2: 1 || COMPUTADORES ALL-IN-ONE || 37 MESES || Dell Pro 24"])
        self.assertFalse(extractor.anexo_tecnico_sin_equipos(garantias, "tecnico__01__ANEXO_N_6_-_GARANTIAS.pdf"))

    def test_prompt_compacto_pide_menos_campos(self):
        import prompts_extraccion as prompts
        self.assertLess(len(prompts.PROMPT_ARCHIVO_COMPACTO), len(prompts.PROMPT_ARCHIVO))
        self.assertNotIn("evidencia", prompts.PROMPT_ARCHIVO_COMPACTO.split("REGLAS:")[0])
        self.assertIn('"tipo"', prompts.PROMPT_ARCHIVO_COMPACTO)


if __name__ == "__main__":
    unittest.main()

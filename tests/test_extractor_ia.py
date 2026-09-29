import importlib.util
import json
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("extractor_ia", ROOT / "3_extraer_ia.py")
extractor = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(extractor)


class ExtractorIATest(unittest.TestCase):
    def test_tipo_documental_con_tildes(self):
        self.assertEqual(extractor.tipo_documental("OFERTA_ECONÓMICA.pdf"), "economico")
        self.assertEqual(extractor.tipo_documental("FICHA_TÉCNICA.pdf"), "tecnico")

    def test_pagina_pdf_blanca_no_queda_pendiente_de_ocr(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "oferta.pdf"
            documento = extractor.fitz.open()
            pagina = documento.new_page()
            pagina.insert_text((72, 72), "Oferta de computadores: 80 unidades, precio neto 462490 pesos cada uno")
            documento.new_page()
            documento.save(extractor.ruta_sistema(ruta))
            documento.close()
            detalle = {}
            texto, estado = extractor.extraer_pdf(ruta, 5, detalle=detalle)
            self.assertEqual(estado, "texto")
            self.assertIn("462490", texto)
            self.assertEqual(detalle["paginas_vacias"], [1])
            self.assertNotIn(1, detalle["paginas_vision"])

    def test_pagina_pdf_con_imagen_no_se_confunde_con_blanca(self):
        with tempfile.TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "escaneada.pdf"
            documento = extractor.fitz.open()
            pagina = documento.new_page()
            pagina.draw_line((72, 72), (180, 72), color=(0, 0, 0), width=2)
            documento.save(extractor.ruta_sistema(ruta))
            documento.close()
            detalle = {}
            extractor.extraer_pdf(ruta, 5, detalle=detalle)
            self.assertEqual(detalle["paginas_vacias"], [])
            self.assertEqual(detalle["paginas_vision"].get(0), "sin_texto")

    @unittest.skipUnless(os.name == "nt", "Rutas largas de Windows")
    def test_archivo_pdf_largo_se_enumera_y_lee(self):
        with tempfile.TemporaryDirectory() as temporal:
            carpeta = Path(temporal) / "Proveedor"
            largo = carpeta / "_extraidos" / ("anexo_economico_" + "a" * 55) / ("oferta_" + "b" * 65)
            Path(extractor.ruta_sistema(largo)).mkdir(parents=True)
            pdf = largo / ("FORMATO_OFERTA_ECONÓMICA_" + "c" * 45 + ".pdf")
            self.assertGreater(len(str(pdf.resolve())), 260)
            try:
                documento = extractor.fitz.open()
                pagina = documento.new_page()
                pagina.insert_text((72, 72), "Desktop OzXen, 80 unidades, $462.490 por unidad, total $36.999.200")
                documento.save(extractor.ruta_sistema(pdf))
                documento.close()
                archivos = extractor.listar_archivos_oferta(carpeta)
                self.assertEqual(archivos, [pdf])
                self.assertNotIn("\\\\?\\", str(archivos[0]))
                texto, estado = extractor.extraer_pdf(pdf, 2)
                self.assertEqual(estado, "texto")
                self.assertIn("462.490", texto)
            finally:
                shutil.rmtree(extractor.ruta_sistema(carpeta / "_extraidos"))

    @unittest.skipUnless(os.name == "nt", "Rutas largas de Windows")
    def test_zip_con_ruta_larga_se_extrae_y_lee(self):
        with tempfile.TemporaryDirectory() as temporal:
            carpeta = Path(temporal) / "Proveedor"
            carpeta.mkdir()
            ruta_interna = ("anexo_" + "a" * 75 + "/" + "oferta_" + "b" * 75 +
                            "/economico_oferta.txt")
            with zipfile.ZipFile(carpeta / "economico.zip", "w") as compacto:
                compacto.writestr(ruta_interna, "Desktop OzXen 80 unidades $462.490 por unidad y $36.999.200 total")
            try:
                registros = extractor.descomprimir(carpeta, None)
                self.assertEqual(registros[0]["estado"], "ok")
                archivos = extractor.listar_archivos_oferta(carpeta)
                extraido = next(p for p in archivos if p.name == "economico_oferta.txt")
                self.assertGreater(len(str(extraido.resolve())), 260)
                texto, estado = extractor.extraer_texto_plano(extraido)
                self.assertEqual(estado, "texto")
                self.assertIn("462.490", texto)
            finally:
                shutil.rmtree(extractor.ruta_sistema(carpeta / "_extraidos"))

    def test_kit_unico_sin_fuente_precio_se_completa_con_anexo_verificado(self):
        bloques = [{"archivo": "economico.pdf", "pagina": 1,
                    "texto": "FILA 10: 1 || 80 || DESKTOP+MONITOR+KIT DE TECLADO || "
                             "80 || $462.490 || $ 36.999.200"}]
        pc = {"producto": "Desktop Tower OzXen A822-SX61V-121F-P0851", "categoria": "equipo",
              "cantidad": None, "precio_unitario": None, "precio_total": None, "fuente_precio": None}
        resultado = extractor.reconciliar_precios_con_anexos([pc], bloques, "$ 36.999.200")[0]
        self.assertEqual((resultado["cantidad"], resultado["precio_unitario"], resultado["precio_total"]),
                         (80, 462490, 36999200))
        self.assertEqual(resultado["fuente_precio"], "economico.pdf")
        con_fuente = extractor.reconciliar_precios_con_anexos(
            [{**pc, "fuente_precio": "economico.pdf"}], bloques, "$ 36.999.200")[0]
        self.assertEqual(con_fuente["precio_unitario"], 462490)

    def test_kit_ambiguo_no_se_asigna_por_proveedor(self):
        bloques = [{"archivo": "economico.pdf", "pagina": 1,
                    "texto": "FILA 10: 1 || 80 || DESKTOP+MONITOR+KIT || 80 || $462.490 || $ 36.999.200"}]
        pcs = [{"producto": f"Desktop OzXen {modelo}", "categoria": "equipo", "cantidad": None,
                "precio_unitario": None, "precio_total": None, "fuente_precio": None}
               for modelo in ("A822", "B900")]
        resultado = extractor.reconciliar_precios_con_anexos(pcs, bloques, "$ 36.999.200")
        self.assertTrue(all(p["precio_unitario"] is None for p in resultado))
        equivocado = extractor.reconciliar_precios_con_anexos(pcs[:1], bloques, "$ 40.000.000")
        self.assertIsNone(equivocado[0]["precio_unitario"])

    def test_descubre_licitaciones_en_una_carpeta_o_varios_meses(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal)
            licitaciones = [
                raiz / "2025-11" / "ofertas" / "1145-63-LE25",
                raiz / "2025-12" / "ofertas" / "1145-70-LE25",
            ]
            for licitacion in licitaciones:
                proveedor = licitacion / "12345678-9__Proveedor"
                proveedor.mkdir(parents=True)
                (proveedor / "oferta.json").write_text("{}", encoding="utf-8")
                (proveedor / "_extraidos").mkdir()
            self.assertEqual(extractor.detectar_licitaciones(licitaciones[0]), licitaciones[:1])
            self.assertEqual(extractor.detectar_licitaciones(raiz / "2025-11"), licitaciones[:1])
            self.assertEqual(extractor.detectar_licitaciones(raiz), licitaciones)

    def test_metadata_de_meses_y_csv_explicito(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal)
            licitacion = raiz / "2030-01" / "ofertas" / "9999-1-LE30"
            licitacion.mkdir(parents=True)
            csv_mes = licitacion.parent.parent / "para_scrapear.csv"
            csv_mes.write_text("codigo,nombre,fecha_publicacion\n"
                               "9999-1-LE30,Oferta de enero,03/01/2030 10:00:00\n", encoding="utf-8")
            metadata = extractor.cargar_metadata_para_corrida([licitacion])
            self.assertEqual(metadata["9999-1-LE30"]["fecha_publicacion"], "03/01/2030 10:00:00")
            csv_usuario = raiz / "metadata_personalizada.csv"
            csv_usuario.write_text("codigo,nombre,fecha_publicacion\n"
                                   "9999-1-LE30,Oferta corregida,04/01/2030\n", encoding="utf-8")
            metadata = extractor.cargar_metadata_para_corrida([licitacion], csv_usuario)
            self.assertEqual(metadata["9999-1-LE30"]["fecha_publicacion"], "04/01/2030")

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

    def test_reconcilia_precio_chileno_desde_fila_economica(self):
        bloques = [{"archivo": "economico.docx", "pagina": None,
                    "texto": "FILA 8: 1 || 80 || Computador Lenovo ThinkCentre M80s Gen3 "
                             "|| 80 || 471.111 || 37.688.880"}]
        producto = {"producto": "Computador Lenovo ThinkCentre M80s Gen3", "modelo": "ThinkCentre M80s Gen3",
                    "categoria": "equipo", "cantidad": 80, "precio_unitario": 471.111,
                    "precio_total": 37688.88, "fuente_precio": "economico.docx"}
        corregido = extractor.reconciliar_precios_con_anexos([producto], bloques)[0]
        self.assertEqual(corregido["precio_unitario"], 471111)
        self.assertEqual(corregido["precio_total"], 37688880)
        self.assertTrue(corregido["precio_reconciliado_fuente"])

    def test_kit_una_fila_no_recibe_dos_precios_inventados(self):
        bloques = [{"archivo": "economico.pdf", "pagina": 1,
                    "texto": "FILA 10: 1 || 80 || Lenovo ThinkCentre M75s Gen 2 SFF "
                             "AMD Ryzen 5 PRO 3350G / 8GB RAM / 512 SSD / Win11 Pro "
                             "+ Monitor ThinkVision E24-40 || 80 || 555.028 || 44.402.240"}]
        pc = {"producto": "Lenovo ThinkCentre M75s Gen 2 SFF AMD Ryzen 5 PRO 3350G / "
                          "8GB RAM / 512 SSD / Win11 Pro", "modelo": "ThinkCentre M75s Gen 2",
              "categoria": "equipo", "cantidad": 80, "precio_unitario": 692.53,
              "precio_total": 55402.4, "fuente_precio": "economico.pdf"}
        monitor = {"producto": "Lenovo ThinkVision E24-40 Monitor 23.8", "modelo": "ThinkVision E24-40",
                   "categoria": "monitor", "cantidad": 80, "precio_unitario": 59,
                   "precio_total": 4720, "fuente_precio": "economico.pdf"}
        corregidos = extractor.reconciliar_precios_con_anexos([pc, monitor], bloques)
        self.assertEqual((corregidos[0]["precio_unitario"], corregidos[0]["precio_total"]),
                         (555028, 44402240))
        self.assertIsNone(corregidos[1]["precio_unitario"])
        self.assertIsNone(corregidos[1]["precio_total"])
        self.assertTrue(corregidos[1]["precio_fuente_no_coincide"])

    def test_no_corrige_precio_si_fila_no_es_verificable(self):
        bloques = [{"archivo": "economico.pdf", "pagina": 1,
                    "texto": "FILA 3: Notebook HP || 2 || 100.000 || 999.999"}]
        producto = {"producto": "Notebook HP", "categoria": "equipo", "cantidad": 2,
                    "precio_unitario": 100000, "precio_total": 200000, "fuente_precio": "economico.pdf"}
        corregido = extractor.reconciliar_precios_con_anexos([producto], bloques)[0]
        self.assertEqual(corregido["precio_total"], 200000)
        self.assertNotIn("precio_reconciliado_fuente", corregido)

    def test_no_asigna_precio_por_coincidencia_generica(self):
        bloques = [{"archivo": "economico.pdf", "pagina": 1,
                    "texto": "FILA 3: 1 || 2 || Notebook HP EliteBook 840 G10 || "
                             "100.000 || 200.000"}]
        producto = {"producto": "Notebook Dell Latitude 5440", "modelo": "Latitude 5440",
                    "categoria": "equipo", "cantidad": 2, "precio_unitario": 100,
                    "precio_total": 200, "fuente_precio": "economico.pdf"}
        corregido = extractor.reconciliar_precios_con_anexos([producto], bloques)[0]
        self.assertIsNone(corregido["precio_unitario"])
        self.assertIsNone(corregido["precio_total"])
        self.assertTrue(corregido["precio_fuente_no_coincide"])

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

    def test_excel_simplificado_con_fuentes_por_campo(self):
        from openpyxl import load_workbook

        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "productos.xlsx"
            extractor.generar_excel(
                [{
                    "codigo": "1234-5-LE25",
                    "fecha_publicacion": "2025-11-01",
                    "proveedor": "Proveedor",
                    "producto": "Notebook HP",
                    "marca": "HP",
                    "cantidad": 2,
                    "precio_unitario": 500000,
                    "precio_total": 1000000,
                    "fuente_producto": "ficha.pdf",
                    "pagina_producto": 2,
                    "fuente_precio": "oferta.pdf",
                    "pagina_precio": 3,
                    "fuente_marca": "ficha.pdf",
                    "pagina_marca": 2,
                    "fuentes_respaldo": ["ficha.pdf", "oferta.pdf"],
                }],
                [], [], ruta,
            )
            libro = load_workbook(ruta, read_only=True, data_only=True)
            self.assertEqual(libro.sheetnames, ["Productos"])
            hoja = libro["Productos"]
            self.assertEqual(
                list(next(hoja.iter_rows(min_row=1, max_row=1, values_only=True))),
                [
                    "Licitación", "Fecha licitación", "Proveedor", "Producto", "Marca producto",
                    "Cantidad", "Precio unitario", "Precio total", "Fuente de información",
                ],
            )
            fila = list(next(hoja.iter_rows(min_row=2, max_row=2, values_only=True)))
            self.assertEqual(tuple(fila[0:8]), (
                "1234-5-LE25", "2025-11-01", "Proveedor", "Notebook HP", "HP", 2, 500000, 1000000,
            ))
            self.assertIn("Producto: ficha.pdf (p. 2)", fila[8])
            self.assertIn("Precio: oferta.pdf (p. 3)", fila[8])
            self.assertIn("Marca: ficha.pdf (p. 2)", fila[8])
            self.assertNotIn("Respaldo:", fila[8])
            libro.close()

    def test_prompt_proveedor_formatea_y_pide_reconstruir_tablas(self):
        prompt = extractor.PROMPT_PROVEEDOR.format(
            proveedor="Proveedor",
            total_oferta="1000",
            nota_imagenes="",
            documentos="tabla de prueba",
        )
        self.assertIn("Reconstruye cada fila", prompt)
        self.assertIn("archivo_precio", prompt)
        self.assertIn("tabla de prueba", prompt)
        self.assertIn("Conserva marcas", prompt)
        self.assertIn('"item":null', prompt)

    def test_filtro_conserva_marcas_desconocidas_y_fichas_sin_precio(self):
        for texto in (
            "Monitor OzXen Pro 24 pulgadas", "Notebook MarcaNueva modelo ABC123",
            "Item 2\nMarca: AOC\nModelo: 24B2XH", "Modelo: ZX123",
            "Panel IPS\nResolucion 1920x1080\nHDMI\nDisplayPort",
        ):
            with self.subTest(texto=texto):
                self.assertTrue(extractor.pagina_relevante(texto))
        self.assertFalse(extractor.pagina_relevante("Declaracion jurada de inhabilidades del proveedor"))

    def test_pdf_tecnico_desconocido_se_envia_junto_al_economico(self):
        args = SimpleNamespace(vision=False, vision_solo_escaneadas=False, ocr=False)
        carpeta = Path("proveedor")
        def lectura(path, args, detalle):
            texto = ("Monitor MarcaNueva ZX123\nMarca: MarcaNueva" if path.name == "tecnico.pdf"
                     else "Monitor 2 $100.000 $200.000")
            detalle.update(texto_paginas={0: texto}, paginas_vision={})
            return texto, "texto"
        with mock.patch.object(extractor, "extraer_archivo", side_effect=lectura):
            bloques, registros, _, _, _ = extractor.preparar_bloques_proveedor(
                carpeta, [carpeta / "economico.pdf", carpeta / "tecnico.pdf"], args)
        self.assertEqual({p["archivo"] for p in bloques}, {"economico.pdf", "tecnico.pdf"})
        self.assertTrue(all(r["paginas_enviadas"] == [1] for r in registros))

    def test_recupera_marca_por_modelo_exacto_y_conserva_fuente(self):
        economico = {"producto": "Monitor", "categoria": "monitor", "modelo": "ZX123", "marca": None}
        tecnico = {"producto": "Monitor MarcaNueva ZX123", "categoria": "monitor", "modelo": "ZX123",
                   "marca": "MarcaNueva", "archivo_fuente": "ficha.pdf", "pagina_producto": 2}
        producto = extractor.completar_marcas_desde_parciales([economico], [tecnico])[0]
        self.assertEqual(producto["marca"], "MarcaNueva")
        self.assertEqual(producto["fuente_marca"], "ficha.pdf")
        self.assertEqual(producto["pagina_marca"], 2)
        self.assertIsNone(economico["marca"])

    def test_no_recupera_marca_con_identidad_ambigua_o_contradictoria(self):
        base = {"producto": "Monitor", "categoria": "monitor", "modelo": "ZX123", "marca": None,
                "item": "2", "rut": "1"}
        tecnico = {**base, "marca": "MarcaNueva"}
        casos = [
            [{**tecnico, "modelo": "ZX124"}],
            [{**tecnico, "categoria": "equipo"}],
            [{**tecnico, "item": "3"}],
            [{**tecnico, "rut": "2"}],
            [tecnico, {**tecnico, "marca": "OtraMarca"}],
        ]
        for parciales in casos:
            with self.subTest(parciales=parciales):
                self.assertIsNone(extractor.completar_marcas_desde_parciales([base], parciales)[0]["marca"])
        generico = {**base, "modelo": None}
        self.assertIsNone(extractor.completar_marcas_desde_parciales([generico], [tecnico])[0]["marca"])
        declarado = {**base, "marca": "MarcaDeclarada"}
        self.assertEqual(extractor.completar_marcas_desde_parciales([declarado], [tecnico])[0]["marca"],
                         "MarcaDeclarada")

    def test_no_fusiona_productos_de_marcas_distintas(self):
        base = {"producto": "Monitor ZX123", "categoria": "monitor", "modelo": "ZX123"}
        self.assertEqual(len(extractor.deduplicar_productos([
            {**base, "marca": "MarcaUno"}, {**base, "marca": "MarcaDos"},
        ])), 2)

    def test_consolidacion_recupera_marca_omitida_sin_otra_llamada(self):
        parcial = {"producto": "Monitor MarcaNueva ZX123", "categoria": "monitor", "modelo": "ZX123",
                   "marca": "MarcaNueva", "archivo_fuente": "ficha.pdf"}
        args = SimpleNamespace(sin_consolidar=False, max_chars_consolidacion=24000, modelo="prueba",
                               timeout=5, num_ctx=8192, backend="lmstudio", max_tokens=1000,
                               reintentos_modelo=0, espera_reintento=0, pausa_archivo=0)
        with mock.patch.object(extractor, "consultar_modelo", return_value=(
                {"productos": [{**parcial, "marca": None}]}, "ok", "{}", {})) as llamada:
            productos, *_ = extractor.consolidar_parciales([parcial], "Proveedor", "1", "", args)
        llamada.assert_called_once()
        self.assertEqual(productos[0]["marca"], "MarcaNueva")
        self.assertEqual(productos[0]["fuente_marca"], "ficha.pdf")

    def test_agrupacion_conserva_ultima_fila_y_encabezados_de_excel_largo(self):
        encabezado = "FILA 1: Producto || Cantidad || Unitario || Total"
        filas = [f"FILA {i}: Notebook HP {i} || 2 || 100 || 200" for i in range(2, 402)]
        bloque = {
            "archivo": "economico.xlsx", "pagina": None, "imagen": None,
            "texto": "\n".join(["[HOJA Oferta]", encabezado, *filas]),
        }
        grupos = extractor.agrupar_bloques([bloque], 1500, 4)
        self.assertGreater(len(grupos), 1)
        textos = [parte["texto"] for grupo in grupos for parte in grupo]
        combinado = "\n".join(textos)
        for fila in filas:
            self.assertEqual(combinado.count(fila + "\n") + int(combinado.endswith(fila)), 1)
        self.assertTrue(all(encabezado in texto for texto in textos))
        self.assertTrue(all("[HOJA Oferta]" in texto for texto in textos))

    def test_no_corta_una_fila_individual_mayor_que_el_limite(self):
        fila = "FILA 2: Notebook HP " + "x" * 2000
        bloque = {"archivo": "oferta.pdf", "pagina": 1, "imagen": None, "texto": fila}
        grupos = extractor.agrupar_bloques([bloque], 500, 4)
        self.assertEqual(grupos[0][0]["texto"], fila)

    def test_ocr_por_proveedor_abre_un_lote_y_conserva_paginas(self):
        args = SimpleNamespace(max_chars_proveedor=24000, vision=False, vision_solo_escaneadas=False,
                               ocr=True, max_paginas=20, tesseract_cmd=None, idioma_ocr="spa+eng",
                               dpi_ocr=180)
        def lectura(path, args, detalle):
            detalle.update(texto_paginas={0: "[PAGINA 1]", 1: "[PAGINA 2]"},
                           paginas_vision={0: "sin_texto", 1: "sin_texto"})
            return "", "texto_vision"

        def ocr(*posicionales, paginas, texto_paginas):
            self.assertEqual(paginas, {0, 1})
            for indice in paginas:
                texto_paginas[indice] = f"[PAGINA {indice + 1} - OCR]\nNotebook HP 2 $100.000 $200.000"
            return "texto OCR", "texto_ocr"

        carpeta = Path("proveedor")
        with mock.patch.object(extractor, "extraer_archivo", side_effect=lectura), \
                mock.patch.object(extractor, "extraer_pdf_ocr", side_effect=ocr) as llamada:
            bloques, registros, pendientes, _, _ = extractor.preparar_bloques_proveedor(
                carpeta, [carpeta / "economico.pdf"], args)
        llamada.assert_called_once()
        self.assertEqual([bloque["pagina"] for bloque in bloques], [1, 2])
        self.assertEqual(registros[0]["paginas_enviadas"], [1, 2])
        self.assertEqual(pendientes, [])

    def test_consumo_acumula_multiples_llamadas_del_mismo_archivo(self):
        args = SimpleNamespace(solo_texto=False, debug=False, max_chars_proveedor=300,
                               max_imagenes_proveedor=4, modelo="qwen-prueba", timeout=5,
                               num_ctx=8192, backend="lmstudio", max_tokens=1024,
                               reintentos_modelo=0, espera_reintento=0, pausa_archivo=0,
                               sin_consolidar=True, vision=False)
        bloque = {"archivo": "economico.txt", "pagina": None, "imagen": None,
                  "texto": "\n".join(["Notebook HP 2 $100.000 $200.000"] * 20)}
        registro = {"archivo": "economico.txt", "estado_ia": "no_ejecutada", "productos_encontrados": 0}
        with tempfile.TemporaryDirectory() as temporal:
            carpeta = Path(temporal)
            respuesta = ({"productos": []}, "ok", "{}", {"total_tokens": 110})
            with mock.patch.object(extractor, "descomprimir", return_value=[]), \
                    mock.patch.object(extractor, "preparar_bloques_proveedor",
                                      return_value=([bloque], [registro], [], [], set())), \
                    mock.patch.object(extractor, "consultar_modelo", return_value=respuesta) as llamada:
                resultado = extractor.procesar_oferta_por_proveedor(carpeta, args, None)
        consumo = resultado["resultados_archivos"][0]["consumo_local"]
        self.assertGreater(llamada.call_count, 1)
        self.assertEqual(consumo["llamadas"], llamada.call_count)
        self.assertEqual(consumo["total_tokens"], 110 * llamada.call_count)
        self.assertEqual(resultado["llamadas_modelo"], llamada.call_count)

    def test_proveedor_cruza_dos_anexos_en_una_llamada_y_reanuda_sin_llamar(self):
        import contextlib
        import io
        import sys

        producto = {
            "producto": "Notebook HP ProBook 440", "marca": "HP", "modelo": "ProBook 440",
            "cantidad": 2, "precio_unitario": 100000, "precio_total": 200000, "categoria": "equipo",
            "archivo_producto": "tecnico.txt", "archivo_precio": "economico.txt",
        }
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal)
            licitacion = raiz / "1234-5-LE25"
            carpeta = licitacion / "1__Proveedor"
            carpeta.mkdir(parents=True)
            extractor.escribir_json_atomico(carpeta / "oferta.json",
                                           {"proveedor": "Proveedor", "rut": "1", "total": "$200.000"})
            (carpeta / "economico.txt").write_text(
                "Producto Cantidad Unitario Total\nNotebook 2 $100.000 $200.000", encoding="utf-8")
            (carpeta / "tecnico.txt").write_text("Notebook HP ProBook 440", encoding="utf-8")
            (carpeta / "administrativo.txt").write_text(
                "Declaracion jurada: el oferente no tiene conflictos de interes.", encoding="utf-8")
            argumentos = ["3_extraer_ia.py", "--dir", str(raiz), "--por-proveedor",
                          "--backend", "lmstudio", "--modelo", "qwen-prueba", "--sin-calentamiento",
                          "--rampa", "0", "--metadata-csv", str(raiz / "sin_metadata.csv"),
                          "--excel", str(raiz / "productos.xlsx")]
            with mock.patch.object(extractor, "consultar_modelo",
                                   return_value=({"productos": [producto]}, "ok", "{}", {"total_tokens": 110})) as llamada, \
                    mock.patch.object(sys, "argv", argumentos), contextlib.redirect_stdout(io.StringIO()):
                extractor.main()
                llamada.assert_called_once()
                prompt = llamada.call_args.args[0]
                self.assertIn("economico.txt", prompt)
                self.assertIn("tecnico.txt", prompt)
                self.assertNotIn("administrativo.txt", prompt)
                extractor.main()
                llamada.assert_called_once()
            checkpoint = json.loads((licitacion / "extraccion_ia.json").read_text(encoding="utf-8"))
            self.assertEqual(checkpoint[0]["llamadas_modelo"], 1)
            self.assertEqual(checkpoint[0]["productos"][0]["precio_unitario"], 100000)
            self.assertEqual(checkpoint[0]["productos"][0]["fuente_producto"], "tecnico.txt")
            self.assertTrue((raiz / "productos.xlsx").is_file())

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

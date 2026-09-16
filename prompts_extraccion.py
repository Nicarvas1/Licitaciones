PROMPT_ARCHIVO = r'''Eres un extractor de datos de ofertas de licitaciones publicas chilenas.
Analiza UN archivo de oferta. Puede ser economico o tecnico.

PROVEEDOR: {proveedor}
RUT: {rut}
ARCHIVO: {archivo}
TIPO DOCUMENTAL: {tipo_documental}
TOTAL MOSTRADO EN EL CUADRO DE OFERTAS: {total_oferta}

Devuelve SOLO JSON valido:
{{
  "productos": [
    {{
      "item": "numero o identificador de item, o null",
      "producto": "descripcion concreta del bien ofertado",
      "marca": "marca o null",
      "modelo": "modelo o null",
      "cantidad": null,
      "cantidad_fuente": "explicita|inferida_total_dividido_unitario|null",
      "precio_unitario": null,
      "precio_total": null,
      "moneda": "CLP|USD|UTM|null",
      "categoria": "equipo|monitor|impresora|consumible|complemento|null",
      "confianza": "alta|media|baja"
    }}
  ],
  "observaciones": "motivo breve si no existe informacion suficiente, o null"
}}

REGLAS:
- Producto es obligatorio cuando el archivo identifica algun bien, incluso sin marca.
- Extrae una fila por item o producto.
- En archivos tecnicos conserva producto, marca y modelo aunque no haya precio.
- En archivos economicos conserva cantidad y precios aunque la descripcion sea generica.
- Busca expresamente campos como "precio por equipo", "monto por equipo",
    "precio unitario", "oferta por equipo", "precio total", "monto total" y
    "total equipos". Si aparecen precio unitario y total de la misma oferta,
    devuelve ambos aunque la cantidad no este escrita.
- Si existe cantidad y precio total de linea, calcula precio_unitario = total/cantidad.
- Si la cantidad no aparece pero precio_total/precio_unitario produce una division
    entera positiva, devuelve ambos precios y deja que el programa infiera la cantidad.
- No uses IVA, subtotal ni total general como producto o precio unitario.
- No inventes. Usa null cuando el dato no aparece.
- Si no hay productos, devuelve {{"productos": [], "observaciones": "motivo"}}.
- El alcance comercial es SOLO: notebook/laptop/portatil, desktop/escritorio/PC,
    all-in-one/AIO, workstation, monitor, impresora/multifuncional y sus tintas o
    toner. Puedes incluir teclado, mouse u otro complemento SOLO si acompana a un
    equipo computacional relevante del mismo archivo/oferta.
- No devuelvas como productos independientes las especificaciones de un equipo:
    procesador, RAM, SSD, HDD, disco, puertos, conectividad, sistema operativo,
    fuente de poder, garantia o servicios. Tampoco devuelvas servidores, storage,
    switches, routers, redes, cables, racks ni servicios de instalacion.
- En cada producto agrega "categoria": "equipo|monitor|impresora|consumible|complemento".

PISTAS DE PRODUCTOS:
{pistas_producto}

PISTAS DE PRECIOS:
{pistas_precio}

TEXTO DEL ARCHIVO:
{texto}
'''

PROMPT_CONSOLIDAR = r'''Eres un reconciliador de resultados extraidos de documentos de una misma oferta.
Une registros que representan el mismo item o producto. Los datos tecnicos pueden
estar en un archivo y el precio/cantidad en otro.

PROVEEDOR: {proveedor}
RUT: {rut}
TOTAL DEL CUADRO DE OFERTAS: {total_oferta}

Devuelve SOLO JSON valido:
{{
  "productos": [
    {{
      "item": "item o null",
      "producto": "descripcion concreta",
      "marca": "marca o null",
      "modelo": "modelo o null",
      "cantidad": null,
      "cantidad_fuente": "explicita|inferida_total_dividido_unitario|null",
      "precio_unitario": null,
      "precio_total": null,
      "moneda": "CLP|USD|UTM|null",
      "categoria": "equipo|monitor|impresora|consumible|complemento|null",
      "fuente_producto": "archivo o null",
      "fuente_precio": "archivo o null",
      "confianza": "alta|media|baja"
    }}
  ],
  "observaciones": "texto breve o null"
}}

REGLAS:
- Une solamente cuando item, orden, descripcion, cantidad o modelo permitan una correspondencia razonable.
- No confundas el total general del proveedor con un precio unitario.
- Si no puedes unir un precio con seguridad, conserva el producto con precio null.
- Si hay total de linea y cantidad, calcula precio_unitario = total/cantidad.
- Busca la cantidad en todos los documentos de la oferta, incluso si aparece solo
    en el tecnico o en el nombre del item. Combina esa cantidad con el precio del
    economico cuando la correspondencia sea razonable.
- La cantidad puede estar en el anexo tecnico, el precio unitario y total en el
    economico, y la marca/modelo en otro documento del mismo proveedor: combina
    esas fuentes cuando describan el mismo equipo.
- Distingue equipo principal de complementos. Conserva un complemento solo si
    tiene precio propio o ayuda a explicar el paquete; no lo mezcles con el precio
    del equipo si el documento no dice que esta incluido.
- Si el total y el precio unitario corresponden a la misma oferta, infiere una
    cantidad entera solo cuando la division sea consistente y marca su origen.
- No inventes productos ni precios.

HALLAZGOS PARCIALES:
{parciales}
'''
PROMPT_ARCHIVO = r'''Eres un extractor de datos de ofertas de licitaciones publicas chilenas.
Analiza UN archivo de oferta. Puede ser economico o tecnico.
Las tablas pueden estar desordenadas o sin tabulaciones/separaciones. Analiza
encabezados, item, descripcion y alineacion para reconstruir correctamente cada
fila y asociar su producto, cantidad, precio unitario y total.

PROVEEDOR: {proveedor}
RUT: {rut}
ARCHIVO: {archivo}
TIPO DOCUMENTAL: {tipo_documental}
TOTAL MOSTRADO EN EL CUADRO DE OFERTAS: {total_oferta}

Devuelve SOLO JSON valido:
{{
  "productos": [
    {{
      "item": null,
      "producto": "descripcion concreta del bien ofertado",
      "marca": null,
      "modelo": null,
      "cantidad": null,
      "cantidad_fuente": "explicita|inferida_total_dividido_unitario|null",
      "precio_unitario": null,
      "precio_total": null,
      "precio_total_tipo": "linea|oferta|null",
      "moneda": "CLP|USD|UTM|null",
      "categoria": "equipo|monitor|impresora|null",
      "pagina": null,
      "fila_fuente": null,
      "evidencia": "fragmento literal breve de la fila",
      "confianza": "alta|media|baja"
    }}
  ],
  "observaciones": "motivo breve si no existe informacion suficiente, o null"
}}

REGLAS:
- Una fila por item ofertado. Conserva marca/modelo aunque falte precio, y
  cantidad/precios aunque la descripcion economica sea generica.
- precio_unitario es por unidad; precio_total es de esa linea. Con cantidad y
  total de la misma fila calcula total/cantidad; sin cantidad devuelve ambos
  precios disponibles para que el programa compruebe la division.
- Marca precio_total_tipo "linea" solo para un total de fila. Nunca asignes IVA,
  subtotal ni total general de la oferta a un producto. Usa precios netos cuando
  existan ambos y numeros sin simbolos ni separadores de miles.
- No inventes: usa null si falta un dato o no puedes asociarlo con seguridad.
- Las filas con || conservan columnas. Si la tabla esta aplanada, reconstruyela
  con encabezados y contexto, sin mezclar valores de items distintos. Si faltan
  elementos para distinguir las columnas, deja el dato dudoso en null.
- Indica pagina/fila y evidencia literal breve de los valores.
- Si no hay productos, devuelve {{"productos": [], "observaciones": "motivo"}}.
- Solo computadores, notebook/laptop/portatil, desktop/PC, all-in-one/AIO,
  workstation, monitores e impresoras/multifuncionales. Excluye TV, proyectores,
  tablets, celulares, servidores, redes, accesorios, consumibles y servicios.
- No separes procesador, RAM, disco, sistema operativo, licencias ni garantias
  como productos. Manten dentro de la descripcion los componentes de un kit
  que no tengan precio propio.

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
      "precio_total_tipo": "linea|oferta|null",
      "moneda": "CLP|USD|UTM|null",
      "categoria": "equipo|monitor|impresora|null",
      "fuente_producto": "archivo o null",
      "fuente_precio": "archivo o null",
      "pagina_producto": null,
      "pagina_precio": null,
      "evidencia": "texto literal breve que respalda la union",
      "confianza": "alta|media|baja"
    }}
  ],
  "observaciones": "texto breve o null"
}}

REGLAS:
- Une solamente cuando item, orden, descripcion, cantidad o modelo permitan una correspondencia razonable.
- No confundas el total general del proveedor con un precio unitario.
- Nunca uses un precio_total_tipo "oferta" para inferir cantidad o precio unitario.
- Si no puedes unir un precio con seguridad, conserva el producto con precio null.
- No cambies cifras para hacerlas coincidir. Conserva la evidencia literal y usa
  confianza baja cuando existan fuentes contradictorias.
- Cuando dos documentos repitan el mismo item, devuelve una sola fila. Prioriza
  la fuente con columnas separadas por || y valores que cumplan cantidad por
  precio unitario igual a total de linea.
- Conserva el nombre exacto de los archivos y las paginas que respaldan producto/marca y precio.
- La suma de productos no puede superar el total de la oferta. Si las fuentes
  no permiten resolver una contradiccion, devuelve una sola fila con los campos
  dudosos en null y confianza baja, en vez de conservar duplicados incompatibles.
- Si hay total DE LA MISMA LINEA y cantidad, calcula precio_unitario = total/cantidad.
- Busca la cantidad en todos los documentos de la oferta, incluso si aparece solo
    en el tecnico o en el nombre del item. Combina esa cantidad con el precio del
    economico cuando la correspondencia sea razonable.
- La cantidad puede estar en el anexo tecnico, el precio unitario y total en el
    economico, y la marca/modelo en otro documento del mismo proveedor: combina
    esas fuentes cuando describan el mismo equipo.
- Devuelve exclusivamente computadores, notebooks/laptops, desktops, all-in-one/AIO,
  workstations, monitores e impresoras/multifuncionales. Excluye siempre TV,
  Smart TV, accesorios, consumibles, instalaciones, servicios y cualquier otro
  producto, aunque tenga precio propio o forme parte de un paquete.
- Si el total y el precio unitario corresponden a la misma oferta, infiere una
    cantidad entera solo cuando la division sea consistente y marca su origen.
- No inventes productos ni precios.

HALLAZGOS PARCIALES:
{parciales}
'''

# Se antepone al texto cuando las paginas se envian como imagen (modo --vision).
REGLAS_VISION = r'''MODO VISION: junto a este mensaje se adjuntan imagenes de paginas del archivo.
- Las IMAGENES son la fuente principal para la estructura: lee cada tabla fila por fila
  y asocia producto, cantidad y precios que esten en la MISMA fila de la imagen.
- El texto extraido que viene abajo puede tener el orden de columnas mezclado. Usalo
  solo para confirmar digitos, modelos y nombres exactos; si contradice la estructura
  de la imagen, manda la imagen.
- El aviso "tabla(s) sin columnas recuperables" se refiere solo al texto extraido:
  esas tablas debes leerlas desde la imagen.
- En "pagina" indica el numero de pagina de la imagen donde esta el producto.
- En "fila_fuente" indica la fila visible de la tabla (por ejemplo "item 2" o "fila 3").
- En "evidencia" transcribe brevemente lo que se ve en esa fila de la imagen.
- Las mismas reglas de alcance, exclusiones y no inventar aplican igual.'''


# Modo --por-proveedor: una llamada con las paginas relevantes de los anexos de un proveedor.
PROMPT_PROVEEDOR = r"""Extrae los equipos ofertados por UN proveedor en una licitacion publica chilena.
Analiza las paginas de sus anexos economicos y tecnicos incluidas abajo. Las tablas pueden
estar desordenadas o sin tabulaciones/separaciones. Reconstruye cada fila usando encabezados,
item, descripcion, alineacion y contexto visual; asocia cantidad, unitario y total al mismo
producto, sin mezclar filas. {nota_imagenes}

PROVEEDOR: {proveedor}
TOTAL DE OFERTA (solo referencia; no asignarlo a productos): {total_oferta}

Devuelve SOLO este JSON, sin texto adicional:
{{"productos":[{{"producto":"","marca":null,"modelo":null,"cantidad":null,"precio_unitario":null,"precio_total":null,"categoria":"equipo|monitor|impresora","archivo_producto":null,"pagina_producto":null,"archivo_precio":null,"pagina_precio":null}}]}}

Extrae solo computadores, notebooks, desktop/PC, all-in-one, workstations, monitores e impresoras;
excluye accesorios, servicios, TV, tablets, servidores y redes. Combina anexos solo si el item
coincide claramente: usa el tecnico para marca/modelo y el economico para cantidad/precios. Si un
componente sin precio propio pertenece a un kit, incluyelo dentro de la descripcion del kit.
Devuelve precio unitario y total de la misma linea; calcula el unitario solo con cantidad y total
de esa linea. Usa precios netos, nunca IVA ni total general. Usa null si falta un dato; no inventes.
Indica archivo y pagina del producto/marca y del precio. Devuelve [] si no hay productos del alcance.

DOCUMENTOS:
{documentos}
"""

NOTA_IMAGENES_PROVEEDOR = r"""
Ademas se adjuntan como IMAGEN algunas paginas (tablas sin bordes o escaneadas), en este orden:
{lista_imagenes}
Usa la imagen para leer la estructura de la tabla (que precio va en que fila) y el texto,
si existe, para confirmar digitos."""

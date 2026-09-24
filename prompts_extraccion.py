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
      "precio_total_tipo": "linea|oferta|null",
      "moneda": "CLP|USD|UTM|null",
      "categoria": "equipo|monitor|impresora|null",
      "pagina": null,
      "fila_fuente": "identificador de fila o null",
      "evidencia": "texto literal breve que respalda producto, cantidad y precio",
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
- Si existe cantidad y precio total DE ESA LINEA, calcula precio_unitario = total/cantidad.
- Si la cantidad no aparece pero precio_total/precio_unitario produce una division
    entera positiva, devuelve ambos precios y deja que el programa infiera la cantidad.
- Si solo aparece un total general de la oferta, no lo asignes a ningun producto.
- Si solo aparece un precio total de linea pero no hay cantidad comprobable, conserva
  precio_total y devuelve precio_unitario null.
- Marca precio_total_tipo como "linea" solo cuando el total pertenece expresamente
  a esa fila; usa "oferta" para totales netos/finales generales.
- No uses IVA, subtotal ni total general como producto o precio unitario.
- No inventes. Usa null cuando el dato no aparece.
- Copia en evidencia el fragmento literal que respalda los valores. Si producto,
  cantidad y precio provienen de lugares distintos, indicalos brevemente.
- Las lineas marcadas como FILA dentro de una TABLA conservan columnas separadas
  por || y son la fuente preferente para asociar producto, cantidad y precios.
- Si aparece "tabla(s) sin columnas recuperables", no interpretes secuencias de
  digitos separadas por espacios como cantidad o precio salvo que otra parte
  estructurada del documento confirme exactamente esos valores.
- Si no hay productos, devuelve {{"productos": [], "observaciones": "motivo"}}.
- El alcance comercial es EXCLUSIVAMENTE: computadores, notebook/laptop/portatil,
  desktop/escritorio/PC, all-in-one/AIO, workstation, monitores e impresoras o
  multifuncionales.
- Excluye siempre televisores, TV, Smart TV, proyectores, tablets, celulares,
  servidores, storage, switches, routers, redes, accesorios, teclado, mouse,
  docking, cables, racks, tintas, toner, cartuchos, repuestos, licencias,
  garantias, instalaciones y servicios, aunque aparezcan junto a un equipo.
- No devuelvas como productos independientes las especificaciones de un equipo:
    procesador, RAM, SSD, HDD, disco, puertos, conectividad, sistema operativo,
  fuente de poder, garantia o servicios.
- En cada producto agrega "categoria": "equipo|monitor|impresora".

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


# Modo --por-proveedor: UNA llamada con las paginas relevantes de todos los anexos
# de un proveedor. Respuesta corta (9 campos) para generar menos tokens.
PROMPT_PROVEEDOR = r"""Extrae los equipos ofertados por UN proveedor en una licitacion publica chilena.
Abajo vienen las paginas relevantes de sus anexos (economicos y tecnicos).{nota_imagenes}

PROVEEDOR: {proveedor}
TOTAL DE SU OFERTA SEGUN EL PORTAL: {total_oferta}

Devuelve SOLO este JSON, sin texto adicional:
{{"productos":[{{"producto":"","marca":null,"modelo":null,"cantidad":null,"precio_unitario":null,"precio_total":null,"categoria":"equipo|monitor|impresora","archivo":"","pagina":null}}]}}

REGLAS:
- Una fila por equipo ofertado. Si el mismo equipo aparece en un anexo economico y en uno tecnico,
  devuelvelo UNA sola vez: cantidad y precios del economico; marca y modelo de donde aparezcan.
- Alcance: SOLO computadores, notebooks, desktop/PC, all-in-one, workstations, monitores e
  impresoras/multifuncionales. Ignora accesorios, licencias, garantias, servicios, despacho,
  instalacion, TV, tablets, servidores y redes, aunque tengan precio propio.
- No devuelvas especificaciones (procesador, RAM, disco, sistema operativo) como productos.
- precio_unitario: precio de UNA unidad. precio_total: total de ESA linea (cantidad x unitario).
  Nunca uses el total general de la oferta, el IVA ni subtotales como precio de un producto.
- Si hay precios con y sin IVA, usa los netos (sin IVA).
- Montos como numeros sin puntos ni simbolos: 650.000 -> 650000.
- archivo: nombre exacto del archivo de donde sale el precio (o el producto, si no hay precio).
  pagina: numero de pagina.
- No inventes: usa null si un dato no aparece. Si no hay equipos, devuelve {{"productos":[]}}.

DOCUMENTOS:
{documentos}
"""

NOTA_IMAGENES_PROVEEDOR = r"""
Ademas se adjuntan como IMAGEN algunas paginas (tablas sin bordes o escaneadas), en este orden:
{lista_imagenes}
Usa la imagen para leer la estructura de la tabla (que precio va en que fila) y el texto,
si existe, para confirmar digitos."""

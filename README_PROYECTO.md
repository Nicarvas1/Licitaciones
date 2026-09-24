# Inteligencia Competitiva de Licitaciones - Mercado Publico (Chile)

## Objetivo

Identificar las marcas, productos y precios ofertados en licitaciones publicas
de computo para comparar a HP con sus competidores.

El proyecto transforma una descarga CSV del buscador de Mercado Publico en un
catalogo auditable de productos ofertados. No consulta solamente al modelo:
primero filtra licitaciones, descarga los anexos reales de cada proveedor,
preserva tablas, aplica OCR cuando corresponde, extrae registros estructurados,
cruza documentos tecnicos y economicos y finalmente valida los resultados antes
de escribirlos en Excel.

## Estado actual

- La ruta principal validada es local: `3_extraer_ia.py` con LM Studio y
  `qwen/qwen3.6-35b-a3b`.
- LM Studio usa su API compatible con OpenAI en
  `http://127.0.0.1:1234/v1/chat/completions` y puede ejecutar el modelo mediante
  Vulkan en una GPU Intel Arc o AMD.
- Ollama sigue soportado en `http://127.0.0.1:11434/api/generate`, pero no es la
  ruta recomendada para Intel Arc en Windows.
- Tesseract esta integrado en el extractor local para PDF escaneado o paginas
  sin texto extraible.
- Existe un extractor separado, `3_extraer_openai_prueba.py`, para comparar
  contra la API de OpenAI. Este carga `OPENAI_API_KEY` desde `.env`.
- `automatizar_pipeline.py` ejecuta el extractor local `3_extraer_ia.py` y
  genera por mes `resultados_ia_local.xlsx`. Los lotes antiguos pueden tener el
  nombre historico `resultados_openai.xlsx`; el consolidado lo usa como respaldo
  si el mes no tiene el archivo nuevo. No significa que se use la API de OpenAI.
- Existe un modo vision selectivo (`--vision`): solo las paginas PDF con tablas
  sin bordes o colapsadas, las paginas escaneadas y los JPG/PNG adjuntos se envian
  como imagen (+ su texto) al modelo. El resto sigue por la ruta de texto.
- El extractor puede procesar varios proveedores a la vez (`--paralelo N`).
- Equipo recomendado: HP Z8 (2x RTX A4000). La Z2 Mini se descarto por
  sobrecalentamiento; las pausas `--pausa-archivo` y `--pausa-licitacion`
  existian por ese equipo y no se necesitan en la Z8.
- El lote descargado de `2025-11` contiene actualmente 75 licitaciones y 682
  carpetas de proveedores.

## Alcance comercial

El contrato de extraccion es deliberadamente estricto.

Se incluyen exclusivamente:

- computadores y notebooks/laptops/portatiles;
- desktop, PC y estaciones de trabajo;
- all-in-one/AIO;
- monitores;
- impresoras y multifuncionales.

Se excluyen, aunque aparezcan dentro de la misma oferta:

- televisores y Smart TV;
- proyectores, tablets, celulares y servidores;
- storage, switches, routers, redes, racks y cables;
- teclados, mouse, docking y otros accesorios;
- tintas, toner, cartuchos y repuestos;
- licencias, garantias, instalaciones y servicios;
- procesador, RAM, SSD, sistema operativo y otras especificaciones cuando el
  modelo las presenta incorrectamente como productos independientes.

La salida de negocio esperada contiene como minimo codigo de licitacion,
proveedor, RUT, producto, marca, cantidad y precio unitario. Los valores ausentes
se conservan como nulos y se marcan para revision; no se inventan.

## Flujo vigente

```text
CSV(s) descargado(s) desde el buscador de Mercado Publico
        |
        v
ListaLicitaciones2026.csv u otro CSV compatible
        |
        |  1_filtrar_licitaciones.py
        v
licitaciones_computo.csv + para_scrapear.csv
        |
        |  2_scraper_ofertas.py
        v
ofertas/<codigo>/<rut>__<proveedor>/
        |
        |  3_extraer_ia.py
        v
ofertas/<codigo>/extraccion_ia.json
        |
        v
Excel mensual + log JSONL + alertas
```

El mismo flujo puede ejecutarse etapa por etapa o mediante
`automatizar_pipeline.py`, que separa la entrada por mes de publicacion y procesa
cada lote de forma reanudable.

## Como funciona cada etapa

### 1. Filtrado del catalogo

`1_filtrar_licitaciones.py` acepta uno o varios CSV separados por coma o punto y
coma, prueba codificacion UTF-8 y Latin-1 y elimina duplicados por codigo.

Conserva licitaciones Publicadas, Cerradas y Adjudicadas cuyo nombre o descripcion
parezcan corresponder a computo. Descarta estados como desierta, cancelada,
revocada, suspendida o sin ofertas, y rubros explicitamente excluidos.

Genera dos archivos:

- `licitaciones_computo.csv`: catalogo completo de licitaciones relevantes;
- `para_scrapear.csv`: solo Cerradas y Adjudicadas, porque sus ofertas ya son
  visibles. Las Publicadas se conservan en el catalogo, pero no se scrapean.

Ambos CSV incluyen metadatos y columnas de control como `scraping_estado` e
`ia_estado`.

### 2. Descarga de ofertas y anexos

`2_scraper_ofertas.py` usa Playwright y Chromium para abrir la ficha de cada
licitacion, localizar el Cuadro de Ofertas, recorrer todas sus paginas y enumerar
todos los proveedores. Por cada proveedor guarda nombre, RUT, total, estado de la
oferta, enlaces de anexos y `oferta.json`.

```text
ofertas/
  <codigo-licitacion>/
    metadata.json
    resumen_ofertas.csv
    <rut>__<proveedor>/
      oferta.json
      economico__01__archivo.pdf
      tecnico__01__archivo.xlsx
```

Por defecto descarga anexos economicos. `--incluye-tecnicos` descarga economicos
y tecnicos en la primera pasada y es el modo usado por el orquestador mensual.
`--todos-anexos` agrega administrativos. El scraper actualiza el CSV despues de
cada licitacion, por lo que una interrupcion no pierde el avance.

### 3. Lectura documental

`3_extraer_ia.py` procesa PDF, DOCX, XLSX, XLSM, TXT, CSV, TSV, ZIP, RAR y 7Z.
Los comprimidos se expanden bajo `_extraidos`; RAR y 7Z requieren 7-Zip.

- PDF: `pdfplumber` preserva texto y tablas. Las filas se representan como
  `FILA N: columna || columna` para mantener la relacion entre producto,
  cantidad y precio.
- PDF escaneado: con `--ocr`, PyMuPDF renderiza solo las paginas sin texto y
  Tesseract ejecuta OCR en `spa+eng`.
- DOCX: se leen parrafos y tablas.
- XLSX/XLSM: se recorren hojas, filas y columnas con `openpyxl`.
- XLS antiguo: se registra como `xls_legacy_no_soportado`.
- Formatos desconocidos: se registran como `no_soportado`.

### 4. Extraccion con IA

Cada documento legible se divide en fragmentos de hasta
`--max-chars-archivo` caracteres, con solapamiento de lineas. Para cada fragmento
se construyen pistas de productos y precios y se llama al modelo (fragmentos en orden; proveedores en paralelo con `--paralelo`).
El modelo debe devolver JSON con item, producto, marca, modelo, categoria,
cantidad, precio unitario, total, moneda, evidencia y confianza.

Los prompts vigentes estan centralizados en `prompts_extraccion.py`. Tanto el
extractor local como el comparador OpenAI usan esas instrucciones.

### 5. Consolidacion por proveedor

Los hallazgos de anexos economicos y tecnicos se deduplican y luego se envian a
una segunda etapa de consolidacion. Esto permite obtener marca y modelo desde una
ficha tecnica, cantidad desde otro anexo y precio desde el economico.

Si los hallazgos exceden `--max-chars-consolidacion`, se dividen en lotes sin
cortar registros JSON. Cuando es posible se realiza una consolidacion final.
`--sin-consolidar` omite estas llamadas y conserva los parciales deduplicados.

### 6. Normalizacion y validacion

La salida del modelo no se acepta directamente. El codigo:

- normaliza formatos numericos chilenos e internacionales;
- clasifica y filtra nuevamente el alcance comercial;
- deduplica por categoria, item, modelo, descripcion y valores numericos;
- infiere cantidad o precio unitario solo si el total pertenece a la misma linea
  y la division es consistente;
- impide usar el total general de la oferta como precio de un producto;
- relaciona cada consolidado con sus hallazgos parciales y fuentes;
- detecta conflictos entre documentos;
- compara cantidad por precio unitario contra el total de linea;
- alerta si la suma de productos supera el total de la oferta.

Cada producto recibe `estado_validacion`:

- `ok`: completo, respaldado y sin alertas;
- `revisar`: presenta alertas no bloqueantes;
- `incompleto`: falta marca, cantidad o precio unitario;
- `inconsistente`: existen cantidades invalidas, conflictos o errores
  matematicos.

### 7. Checkpoints y reanudacion

Despues de cada proveedor se actualiza atomicamente `extraccion_ia.json` dentro
de la licitacion. Cada corrida escribe tambien un log JSONL con inicio y fin de
proveedores, productos, errores y uso del modelo.

Al repetir el comando sin `--rehacer`, el extractor reutiliza resultados
compatibles y repite solo lo necesario (ver "Reproceso selectivo").
`--rehacer` elimina ese beneficio y vuelve a procesar todo.

### Reproceso selectivo

Para cada proveedor ya guardado en `extraccion_ia.json` se decide, archivo por
archivo, que repetir:

- `archivo_nuevo`: el archivo no estaba en el resultado anterior (por ejemplo,
  anexos tecnicos descargados despues con `--solo-tecnicos`);
- `fallo_modelo`: la llamada al modelo para ese archivo fallo;
- `pendiente_ocr`: el archivo quedo en `necesita_ocr` y ahora se usa `--ocr` o
  `--vision`;
- `paginas_para_vision` / `imagen`: se activo `--vision` y el archivo tiene
  paginas que lo requieren, o es una imagen.

Los demas archivos reutilizan su registro y sus productos parciales, y luego se
vuelve a consolidar el proveedor con todo junto. Si solo fallo la consolidacion,
se repite solo esa llamada. Un cambio de `--modelo` o `--backend` sigue
rehaciendo el proveedor completo. En la consola, cada proveedor revisado muestra
`rehecho [motivo] archivo` o `sin cambios`, y el JSON guarda el detalle en
`reproceso`.

### 8. Excel final

El extractor local genera cinco hojas:

- `Productos`: todos los productos dentro del alcance, con evidencia y alertas;
- `Productos_validos`: solamente registros con validacion `ok`;
- `Alertas`: productos o proveedores que requieren revision;
- `Consumo`: llamadas y tokens reportados por el backend;
- `Resumen`: cobertura, productos, errores, OCR y no soportados por proveedor.

## Componentes conservados

- `1_filtrar_licitaciones.py`: filtra el CSV del buscador por computo y por
  estados Publicada, Cerrada y Adjudicada. Genera los dos CSV intermedios.
- `2_scraper_ofertas.py`: lee `para_scrapear.csv`, descarga anexos tecnicos y
  economicos, y actualiza `scraping_estado` para poder reanudar una corrida.
- `3_extraer_ia.py`: extractor principal local; soporta LM Studio y Ollama,
  fragmentacion, OCR, consolidacion, validacion, checkpoints y Excel.
- `3_extraer_openai_prueba.py`: comparador opcional contra la API de OpenAI.
- `automatizar_pipeline.py`: prepara lotes mensuales y coordina filtro, scraping,
  extraccion local y reportes.
- `prompts_extraccion.py`: contrato JSON e instrucciones compartidas de
  extraccion y consolidacion.
- `tests/test_extractor_ia.py`: regresiones de normalizacion, alcance,
  deduplicacion, validacion y reproceso.
- `ListaLicitaciones2026.csv`: entrada descargada manualmente desde el buscador.
- `licitaciones_computo.csv`, `para_scrapear.csv` y `ofertas/`: datos y estado
  generados por el flujo.

## Ejecucion

Instalar dependencias una vez:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
```

La version local comparte con la prueba OpenAI el filtro de productos, la
consolidacion economico-tecnica, la inferencia de cantidad por
`precio_total/precio_unitario`, el cruce de metadatos y el reporte con hojas
`Productos`, `Consumo` y `Resumen`.

Ejemplo recomendado y validado con LM Studio y Qwen:

```powershell
.\.venv\Scripts\python.exe .\3_extraer_ia.py `
  --dir ".\lotes\2025-11\ofertas" `
  --backend lmstudio `
  --modelo "qwen/qwen3.6-35b-a3b" `
  --metadata-csv ".\lotes\2025-11\para_scrapear.csv" `
  --excel ".\lotes\2025-11\resultado_local_qwen.xlsx" `
  --num-ctx 8192 `
  --max-tokens 4096 `
  --timeout 600 `
  --reintentos-modelo 2 `
  --ocr
```

Agrega `--rehacer` solamente si quieres descartar checkpoints previos. Para una
primera prueba controlada se recomienda agregar `--limite-licitaciones 1` y
`--limite-proveedores 1`.

Ollama se instala por separado para `3_extraer_ia.py`. 7-Zip tambien es una
instalacion externa y permite procesar anexos `.rar` y `.7z`.

Filtrar la descarga del buscador:

```powershell
python 1_filtrar_licitaciones.py --in ListaLicitaciones2026.csv
```

Probar el scraper con tres licitaciones, mostrando el navegador:

```powershell
python 2_scraper_ofertas.py --csv para_scrapear.csv --limite 3 --ver
```

Ejecutar el scraper completo o reintentar errores:

```powershell
python 2_scraper_ofertas.py --csv para_scrapear.csv
python 2_scraper_ofertas.py --csv para_scrapear.csv --reintentar-errores
```

Extraer datos con el modelo local configurado en Ollama:

```powershell
python 3_extraer_ia.py --dir ofertas --modelo llama3.2
```

Comparar localmente con la prueba OpenAI usando `gpt-oss:20b`:

```powershell
ollama list
ollama run gpt-oss:20b "Responde SOLO con JSON valido: {\"ok\": true}"

.\.venv\Scripts\python.exe .\3_extraer_ia.py `
  --dir ".\ofertas\3572-22-LE25" `
  --modelo "gpt-oss:20b" `
  --sin-consolidar `
  --rehacer `
  --metadata-csv ".\para_scrapear.csv" `
  --excel ".\resultado_local_3572.xlsx"
```

El extractor local no necesita `OPENAI_API_KEY`: se comunica con Ollama en
`http://127.0.0.1:11434` o con LM Studio en `http://127.0.0.1:1234`. Genera
`extraccion_ia.json` dentro de la licitacion y
un Excel con las mismas columnas principales del reporte OpenAI, incluyendo
nombre, fecha de publicacion, estado y organismo. `--sin-consolidar` hace que
la comparacion sea mas pareja, porque analiza los archivos sin una segunda
llamada de consolidacion.

Los extractores filtran el resultado exclusivamente a computadores, notebooks,
desktops/PC, AIO, workstations, monitores e impresoras/multifuncionales. TV,
Smart TV, accesorios, consumibles, instalaciones, servicios y productos no
relacionados quedan fuera.

El extractor procesa los archivos de un proveedor en orden y, con `--paralelo N`, varios proveedores a la vez. Los documentos
largos se dividen en fragmentos con solapamiento para no perder filas ni aumentar
el contexto de una sola llamada. El Excel contiene:

- `Productos`: todos los productos dentro del alcance, con evidencia y alertas.
- `Productos_validos`: solo registros completos y sin inconsistencias detectadas.
- `Alertas`: filas o proveedores que requieren revision.
- `Consumo`: llamadas y tokens reportados por LM Studio u Ollama.
- `Resumen`: cobertura por proveedor, OCR y errores del modelo.

Cada corrida también genera un archivo `.log.jsonl`. Si el equipo se reinicia,
ejecuta nuevamente el mismo comando sin `--rehacer`; se conservan los proveedores
terminados y se retoman los pendientes o fallidos.

Para activar OCR en Windows instala Tesseract y las dependencias Python:

```powershell
winget install --id UB-Mannheim.TesseractOCR -e
pip install -r requirements.txt
```

Si Tesseract no queda en `PATH`, usa
`--tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"`.

Prueba recomendada de 10 licitaciones, una a la vez, con LM Studio:

```powershell
python .\3_extraer_ia.py `
  --dir ".\lotes\2025-11\ofertas" `
  --backend lmstudio `
  --modelo "qwen/qwen3.6-35b-a3b" `
  --metadata-csv ".\lotes\2025-11\para_scrapear.csv" `
  --excel ".\lotes\2025-11\prueba_local_10.xlsx" `
  --limite-licitaciones 10 `
  --num-ctx 8192 `
  --max-tokens 4096 `
  --timeout 600 `
  --reintentos-modelo 2 `
  --pausa-archivo 2 `
  --ocr `
  --rehacer
```

Usa `--rehacer` solo en el primer intento de una prueba que deba reemplazar
resultados anteriores. Para reanudar después de una interrupcion, repite el
mismo comando quitando `--rehacer`.

## Modo por proveedor (`--por-proveedor`)

El modo original hace una llamada por cada fragmento de cada archivo, mas una
consolidacion por proveedor. En noviembre (563 proveedores, corrida secuencial)
tomo 10,2 horas: mediana de 42 s por proveedor, y el 16% mas lento consumio la
mitad del tiempo (muchos anexos, fichas tecnicas largas u OCR).

Con `--por-proveedor`:

- se leen todos los anexos del proveedor pagina por pagina y solo se envian las
  paginas con montos, o con un equipo junto a una marca; declaraciones, bases y
  formularios administrativos quedan fuera;
- si el filtro descarta TODAS las paginas de un proveedor, igual se envian sus
  anexos (economicos primero) hasta el limite de una llamada, marcados con
  `respaldo_sin_filtro`: el filtro ahorra tiempo pero no deja proveedores sin revisar;
- todo va en UNA llamada (anexos economicos primero). Si no cabe en
  `--max-chars-proveedor` (24.000 caracteres por defecto) se usan mas llamadas y
  solo entonces se consolida;
- el modelo cruza economico y tecnico en la misma respuesta, asi que desaparece
  la llamada de consolidacion;
- la respuesta es corta: 9 campos por producto en vez de 15;
- paginas escaneadas: como imagen con `--vision` (maximo
  `--max-imagenes-proveedor` por llamada) o por OCR solo esa pagina con `--ocr`.

El resultado tiene el mismo formato (validacion, Excel, checkpoints, revision).
Cada proveedor registra `modo`, `llamadas_modelo` y `segundos`, y el log agrega
`segundos` por proveedor. Cambiar de modo rehace el proveedor completo, por eso
para comparar modos conviene usar una copia del lote.

```powershell
.\.venv\Scripts\python.exe .\3_extraer_ia.py `
  --dir ".\copia_prueba\ofertas" --backend lmstudio --modelo "qwen/qwen3.6-35b-a3b" `
  --por-proveedor --paralelo 4 --vision --max-tokens 4096 --timeout 600
```

No usar `--pausa-licitacion` ni `--pausa-archivo` en la Z8.

### Vision solo para lo que el OCR no pudo leer (`--vision-solo-escaneadas`)

`--vision` envia como imagen las paginas con tablas dificiles y las escaneadas, y
al activarlo reprocesa todos los proveedores leidos sin vision. Para usar la
vision solo donde es imprescindible:

```powershell
.\.venv\Scripts\python.exe .\3_extraer_ia.py --dir "..." --backend lmstudio --modelo "..." `
  --por-proveedor --ocr --vision --vision-solo-escaneadas `
  --paralelo 1 --max-imagenes-proveedor 2 --dpi-vision 120
```

- las paginas escaneadas pasan primero por OCR; solo las que el OCR no pudo leer
  van como imagen;
- sobre un lote ya procesado, solo se reprocesan los proveedores con OCR pendiente;
- las imagenes van en JPEG (una pagina escaneada pesa ~6 veces menos que en PNG).

### Uso remoto con LM Link

Si el script corre en otro equipo y usa el modelo de la Z8 mediante LM Link, cada
peticion cruza un tunel cifrado (Tailscale). Al arrancar, varias peticiones
grandes simultaneas pueden provocar "Channel Error" en LM Studio. Por eso:

- antes de empezar se envia una peticion minima de calentamiento, que activa el
  enlace y confirma que el modelo responde (`--sin-calentamiento` la omite);
- los primeros proveedores en paralelo arrancan escalonados cada `--rampa`
  segundos (5 por defecto);
- "Channel Error" y errores de modelo cargandose se reintentan como transitorios
  (`--reintentos-modelo`).

Lo mas robusto es correr el script directamente en la Z8, con los documentos en
esa maquina: se elimina el tunel y el envio de imagenes por la red.

## Paralelismo y modo vision

### Paralelismo (`--paralelo N`)

Antes el extractor mandaba una sola peticion a la vez y la GPU quedaba esperando
entre llamadas. Con `--paralelo 4` se procesan 4 proveedores simultaneamente,
cada uno en su propio hilo. LM Studio atiende esas peticiones juntas mediante
continuous batching.

- En LM Studio, al cargar el modelo, activar "Manually choose model load
  parameters" y en "Show advanced settings" poner **Max Concurrent Predictions**
  igual a `--paralelo`. Si el script manda mas peticiones que ese valor, las
  sobrantes esperan en cola y cuentan contra `--timeout`.
- El context length del modelo se comparte entre las peticiones simultaneas.
  Para 4 en paralelo partir con 65536; si no carga o hay errores de memoria,
  usar 32768 y `--paralelo 2`.
- Los checkpoints, el log y el Excel no cambian de formato. Solo el hilo
  principal escribe `extraccion_ia.json`, despues de cada proveedor.
- Si LM Studio deja de responder, se cancelan los proveedores en cola, los que
  estaban en curso se abandonan en su siguiente llamada y la corrida se detiene.
  Repetir el comando retoma lo pendiente.
- **Ctrl+C** cancela la cola y abandona lo que esta en curso; solo se espera la
  respuesta de las llamadas activas. Un segundo Ctrl+C sale de inmediato. Los
  checkpoints se escriben de forma atomica, asi que siempre quedan validos. El
  extractor termina con codigo 130 y el orquestador espera su cierre ordenado.
- Un proveedor que se estaba reprocesando conserva su resultado anterior en
  `extraccion_ia.json` si el reproceso falla, se cancela o se interrumpe.
- Para usar LM Studio u Ollama en otro puerto o equipo, definir las variables de
  entorno `LMSTUDIO_URL` u `OLLAMA_URL`.
- `--pausa-licitacion` se ignora con `--paralelo` mayor que 1.

### Modo vision (`--vision`)

El texto extraido de tablas sin bordes puede mezclar el orden de las columnas.
Con `--vision`, cada pagina PDF se clasifica:

- `sin_texto`: pagina escaneada;
- `tabla_colapsada`: pdfplumber detecto una tabla pero no separo sus columnas, y
  la pagina menciona un equipo del alcance y precios;
- `tabla_sin_bordes`: no hay tabla detectada, pero la pagina tiene al menos
  `--vision-min-filas` (2 por defecto) lineas con forma de fila de precios (un
  monto y otro numero, como cantidad o item), una palabra de precio y un equipo
  del alcance (notebook, desktop, AIO, workstation, monitor, impresora, o
  "equipo" sin mencion de TV).

Montos sueltos (garantias, totales) y tablas de productos fuera del alcance no
activan vision. Las paginas que no califican no se pierden: siguen por texto.

Para calibrar la regla con documentos reales sin llamar al modelo ni tocar
checkpoints:

```powershell
.\.venv\Scripts\python.exe .\3_extraer_ia.py `
  --dir ".\lotes\2025-11\ofertas" `
  --excel ".\lotes\2025-11\resultado_local_qwen.xlsx" `
  --limite-licitaciones 5 `
  --diagnostico-vision
```

Muestra por archivo que paginas irian a vision y por que, un resumen con la
cantidad estimada de llamadas con imagen, y guarda `diagnostico_vision.csv`
junto al Excel. Si marca paginas que no deberian ir a vision, subir
`--vision-min-filas`.

Esas paginas se renderizan a PNG (`--dpi-vision`, 150 por defecto) y se envian
en grupos de `--vision-paginas-por-llamada` (3 por defecto), junto con su texto
extraido para confirmar cifras. Las demas paginas del mismo PDF siguen por texto.
Los JPG/PNG/WEBP adjuntos tambien se leen por vision. DOCX y Excel no cambian,
porque ya conservan su estructura de tabla. Las paginas enviadas como imagen no
pasan por Tesseract.

Las instrucciones extra para vision estan en `REGLAS_VISION` dentro de
`prompts_extraccion.py`. Cada archivo registra en `extraccion_ia.json` las
paginas enviadas como imagen (`paginas_vision`) y el motivo.

Al activar `--vision` sobre un lote ya procesado sin vision, NO se rehace todo:
se revisa cada PDF sin llamar al modelo y solo se repiten los archivos que tienen
paginas para vision, los JPG/PNG y los PDF que quedaron pendientes de OCR (ver
"Reproceso selectivo"). Para comparar ambos modos, probar sobre una copia del lote.

Comando recomendado en la Z8:

```powershell
.\.venv\Scripts\python.exe .\3_extraer_ia.py `
  --dir ".\lotes\2025-11\ofertas" `
  --backend lmstudio `
  --modelo "qwen/qwen3.6-35b-a3b" `
  --metadata-csv ".\lotes\2025-11\para_scrapear.csv" `
  --excel ".\lotes\2025-11\resultado_local_qwen.xlsx" `
  --paralelo 4 `
  --vision `
  --max-chars-archivo 20000 `
  --max-tokens 4096 `
  --timeout 600 `
  --reintentos-modelo 2 `
  --ocr
```

`--ocr` sigue sirviendo para PDF escaneados cuando no se usa `--vision`.

## Experimento: extraccion solo con reglas (sin IA)

`3_extraer_reglas.py` mide cuanto se puede extraer solo con codigo. No usa
modelo ni modifica `extraccion_ia.json`; escribe `extraccion_reglas.json` en cada
licitacion y un Excel aparte.

Tecnicas, de mas a menos confiable:

- **Tablas con encabezados**: en Excel, DOCX, CSV y PDF con tablas con bordes,
  identifica las columnas de descripcion, marca, modelo, cantidad, precio
  unitario y total (tambien encabezados en dos filas y tablas que siguen en la
  pagina siguiente). Prefiere columnas netas sobre con IVA y la descripcion
  "ofertada" sobre la "solicitada". Omite filas de total, neto e IVA.
- **Filas cuadradas**: en texto libre o tablas sin bordes, acepta una linea solo
  si tiene cantidad, precio unitario y total con cantidad x unitario = total.
- **Catalogo de marcas y modelos**: familias como ProBook, ThinkPad, Latitude,
  LaserJet, y respaldo "codigo despues de la marca" (LG 24MK430H). Si el anexo
  economico no trae marca, la toma del anexo tecnico del mismo proveedor cuando
  hay un solo modelo de esa categoria, y solo para filas que por su propio texto
  son equipos del alcance (nunca a discos, UPS, despacho, licencias, etc.).

Otras reglas: los montos se aislan de textos como "$ 650.000 c/u", "650.000 + IVA"
o "650.000.-"; si el total de una fila es exactamente 1,19 veces cantidad x
unitario, se interpreta como total con IVA y se deja neto (columna `nota_iva`);
en texto libre, la descripcion incluye las lineas sin montos que estan justo
encima de la fila con los montos. Cada producto trae `subcategoria`: notebook,
all-in-one, desktop, workstation, monitor o impresora.

Los productos pasan por la misma validacion y filtro de alcance que la IA. Un
proveedor queda `resuelto` si todos sus productos validan `ok`, `parcial` si se
encontro algo incompleto o inconsistente, y `sin_resultado` si no se encontro
nada. PDF escaneados, imagenes y cartas con un solo total no se intentan: quedan
para la IA.

```powershell
.\.venv\Scripts\python.exe .\3_extraer_reglas.py `
  --dir ".\lotes\2025-11\ofertas" `
  --metadata-csv ".\lotes\2025-11\para_scrapear.csv" `
  --excel ".\lotes\2025-11\resultado_reglas.xlsx"
```

Si la licitacion ya tiene `extraccion_ia.json`, el Excel agrega la hoja
`Comparacion` y la consola resume cuantos proveedores tienen los mismos precios
unitarios en ambos metodos. La IA no es la verdad: las filas `precios_distintos`
y `precios_parcialmente_iguales` hay que revisarlas contra el PDF. La hoja
`Archivos` muestra que metodo funciono en cada archivo, util para ver que
formatos dominan.

Pruebas: `python -m unittest tests.test_extraer_reglas -v`

### Revision con evidencia (trazabilidad)

`4_revisar_extraccion.py` genera una pagina HTML con una tarjeta por producto:

- PDF: recorte de la pagina con la fila marcada en rojo y las celdas de cantidad
  (azul), precio unitario (verde) y total (morado) resaltadas; el encabezado de
  la tabla en gris y, si la descripcion venia de la linea anterior, esa linea en
  naranja.
- Excel / Word / CSV: fragmento de la planilla con la fila destacada, sus
  encabezados y las filas vecinas.
- Enlace para abrir el documento original (en la pagina correcta si es PDF).
- Botones Correcto / Incorrecto / Dudoso y comentario. Se guardan en el navegador
  y se exportan con "Exportar CSV".

En PDF, la fila se ubica buscando los montos del producto con PyMuPDF (la misma
libreria que dibuja, asi que funciona tambien con paginas rotadas) y se verifica
palabra por palabra. Si el precio aparece varias veces, se elige la linea que
ademas contiene el total, la cantidad y palabras de la descripcion. Estados:

- `fila verificada`: precio unitario, total y cantidad estan en la misma linea;
- `datos no calzan con la fila`: se encontro el precio, pero la cantidad o el
  total extraidos no estan en esa linea (datos mezclados de otra fila);
- `precio NO encontrado`: el monto no aparece en el documento (error o PDF
  escaneado).

Los dos ultimos son los primeros que hay que revisar.

```powershell
# resultados por reglas (coordenadas exactas)
.\.venv\Scripts\python.exe .\4_revisar_extraccion.py --dir ".\lotes\2025-11\ofertas"
# resultados de la IA (ubicados buscando el precio en el documento)
.\.venv\Scripts\python.exe .\4_revisar_extraccion.py --dir ".\lotes\2025-11\ofertas" --fuente ia
```

Se crea `lotes\2025-11\revision_reglas\index.html` (o `revision_ia`), que se
abre con doble clic en Chrome o Edge. Las imagenes quedan en la subcarpeta `img`.

Forma sugerida de revisar:

1. Filtrar "Precio NO encontrado" y "Otro metodo: precio distinto": ahi se
   concentran los errores.
2. Activar "Solo muestra aleatoria" y revisar esas tarjetas (50 por defecto,
   `--muestra` lo cambia). La barra superior estima la precision de todo el lote
   con un intervalo de confianza del 95%.
3. Exportar el CSV con las decisiones.

Las decisiones se guardan en el almacenamiento del navegador: se conservan al
recargar, pero se pierden si se borra el historial o se usa otro navegador.
Conviene exportar el CSV al terminar cada sesion. La columna `id_revision` del
Excel de reglas coincide con el `id` del CSV.

Pruebas: `python -m unittest tests.test_revision -v`

## Comparacion opcional con OpenAI

`3_extraer_openai_prueba.py` no forma parte de la ruta local principal. Se usa
para comparar calidad y costo con `gpt-4.1-mini` u otro modelo compatible. El
script carga `.env` desde la raiz del proyecto mediante
`OPENAI_API_KEY=<clave>`; `.env` no debe versionarse ni compartirse.

El comparador recorre todas las licitaciones bajo `--dir`, pero por seguridad su
valor predeterminado es un solo proveedor. Para un mes completo debe indicarse
`--limite-proveedores 0`.

La implementacion OpenAI comparte prompts, alcance, normalizacion basica y
consolidacion, pero no posee toda la capa avanzada de validacion del extractor
local y actualmente no admite `--ocr`. Por eso sus resultados no deben
compararse como si ambos Excel tuvieran exactamente el mismo posprocesamiento.

Para regenerar un Excel OpenAI a partir de checkpoints existentes, sin volver a
llamar a la API para proveedores ya terminados, ejecuta sin `--rehacer`:

```powershell
.\.venv\Scripts\python.exe .\3_extraer_openai_prueba.py `
  --dir ".\lotes\2025-11\ofertas" `
  --modelo "gpt-4.1-mini" `
  --limite-proveedores 0 `
  --metadata-csv ".\lotes\2025-11\para_scrapear.csv" `
  --excel ".\lotes\2025-11\resultados_openai_filtrados.xlsx"
```

Procesar nuevamente todas las licitaciones de noviembre con OpenAI:

```powershell
Remove-Item Env:OPENAI_API_KEY -ErrorAction SilentlyContinue
.\.venv\Scripts\python.exe .\3_extraer_openai_prueba.py `
  --dir ".\lotes\2025-11\ofertas" `
  --modelo "gpt-4.1-mini" `
  --limite-proveedores 0 `
  --metadata-csv ".\lotes\2025-11\para_scrapear.csv" `
  --excel ".\lotes\2025-11\resultado_openai_mes_2025-11.xlsx" `
  --timeout 600 `
  --max-tokens 4096 `
  --rehacer
```

El `Remove-Item` evita que una variable temporal incorrecta de PowerShell tape
la clave real del `.env`. No imprime ni modifica la clave guardada.

La prueba OpenAI cruza cada codigo con `para_scrapear.csv` para incorporar el
nombre y la fecha de publicacion. Genera `extraccion_openai.json` dentro de cada
licitacion y un Excel con las hojas `Productos`, `Resumen` y `Consumo`. Si la
corrida se interrumpe, el mismo comando reanuda los proveedores pendientes. Usa
`--rehacer` solo cuando necesites descartar resultados previos y pagar nuevamente
las llamadas ya realizadas.

## Automatizacion local por mes

`automatizar_pipeline.py` ejecuta el flujo local completo con una sola orden. Usa
la fecha de publicacion para separar las licitaciones, reinicia el scraper entre
meses y conserva el estado de cada codigo para poder reanudar una corrida. Crea
un archivo `lotes/.pipeline.lock` para impedir dos ejecuciones simultaneas; si el
proceso termino abruptamente y no existe otra corrida activa, puede ser necesario
eliminar manualmente ese lock.

La descarga principal obtiene anexos economicos y tecnicos juntos. Esto permite
relacionar cantidad, especificaciones y precio aunque esten repartidos en varios
documentos del mismo proveedor. Los documentos administrativos siguen fuera,
salvo que se use `--todos-anexos`.

Primero se puede revisar la distribucion sin abrir el navegador ni ejecutar
ningun modelo:

```powershell
.\.venv\Scripts\python.exe .\automatizar_pipeline.py `
  --in ".\ListaLicitaciones2026.csv" `
  --salida ".\lotes" `
  --solo-preparar
```

Para ejecutar solamente noviembre de 2025 con LM Studio, anexos tecnicos, OCR y
reporte:

```powershell
.\.venv\Scripts\python.exe .\automatizar_pipeline.py `
  --in ".\ListaLicitaciones2026.csv" `
  --salida ".\lotes" `
  --desde 2025-11 `
  --hasta 2025-11 `
  --backend lmstudio `
  --modelo "qwen/qwen3.6-35b-a3b" `
  --num-ctx 8192 `
  --max-tokens 4096 `
  --timeout 600 `
  --reintentos-modelo 2 `
  --ocr
```

Se puede limitar el periodo con `--desde YYYY-MM --hasta YYYY-MM`. La salida
queda organizada asi:

```text
lotes/
  catalogo_licitaciones.csv
  catalogo_para_scrapear.csv
  pipeline.log
  resultados_consolidados.xlsx
  2025-11/
    licitaciones_computo.csv
    para_scrapear.csv
    ofertas/
    resultados_ia_local.xlsx  # antes resultados_openai.xlsx (nombre historico)
  2025-12/
    ...
```

`resultados_consolidados.xlsx` une los Excel mensuales con las hojas `Productos`,
`Productos_validos`, `Alertas`, `Resumen` (con fila TOTAL) y `Consumo`.

Si el proceso se interrumpe, se ejecuta la misma orden. El scraper omite las
licitaciones con estado `ok` y el extractor local reutiliza los proveedores ya
guardados en `extraccion_ia.json`. Para volver a intentar errores del scraper se
agrega `--reintentar-errores`. No se debe usar `--rehacer-ia` salvo que se quiera
repetir toda la inferencia local.

La primera pasada manual tambien puede ejecutarse por separado:

```powershell
python 2_scraper_ofertas.py --csv para_scrapear.csv --incluye-tecnicos
python 3_extraer_ia.py --dir ofertas --backend lmstudio --modelo "qwen/qwen3.6-35b-a3b" --ocr
```

Para completar tecnicos de un lote que ya descargo economicos, sin volver a
descargar los economicos:

```powershell
python 2_scraper_ofertas.py --csv para_scrapear.csv --solo-tecnicos
python 3_extraer_ia.py --dir ofertas --backend lmstudio --modelo "qwen/qwen3.6-35b-a3b" --ocr
```

## Notas operativas

- Los anexos de ofertas se obtienen por scraping web; la API de Mercado Publico
  no expone esos adjuntos de forma util para este flujo.
- Solo las licitaciones Cerrada y Adjudicada son scrapeables. Las Publicadas se
  mantienen en `licitaciones_computo.csv`, pero no pasan a `para_scrapear.csv`.
- El scraper guarda el estado despues de cada licitacion. No borres
  `para_scrapear.csv` ni `ofertas/` si necesitas reanudar el trabajo.
- Los PDF sin texto quedan registrados como `necesita_ocr` si no se usa OCR. Con
  `--ocr`, Tesseract procesa las paginas sin texto y el estado pasa a
  `texto_ocr` cuando obtiene contenido util.
- Por defecto el procesamiento es secuencial (`--paralelo 1`). El tiempo total
  depende principalmente del numero de fragmentos y llamadas de consolidacion;
  `--paralelo` y un `--max-chars-archivo` mayor reducen ese tiempo.
- El modo `--vision` no convierte todo a imagen: solo las paginas donde el texto
  pierde la estructura. Convertir todo seria mas lento y no mas preciso para
  tablas que ya se leen bien como texto.
- No ejecutes simultaneamente dos extractores sobre la misma carpeta de ofertas:
  ambos escriben checkpoints por licitacion y pueden competir por GPU o API.
- La etapa de dashboard consolidado para este flujo aun no esta implementada.

## Diagnostico rapido

- LM Studio no responde: confirmar que el servidor local este iniciado, que el
  modelo indicado por `--modelo` este cargado y que el puerto sea `1234`.
- Ollama no responde: ejecutar `ollama list` y comprobar el puerto `11434`.
- OCR no funciona: probar `tesseract --version` o usar
  `--tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"`.
- El pipeline indica que existe `.pipeline.lock`: confirmar que no haya otra
  corrida y eliminar `lotes\.pipeline.lock` solo si es un lock obsoleto.
- El Excel no se puede sobrescribir: cerrarlo en Excel y volver a generar.
- Un proveedor aparece sin productos: revisar las hojas `Alertas` y `Resumen`,
  los estados por archivo en `extraccion_ia.json` y el log JSONL.

## Pruebas

Las pruebas de regresion se ejecutan con:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_extractor_ia
```

Cubren reglas de alcance, normalizacion de numeros, deduplicacion, validacion,
conflictos y decisiones de reproceso.

Las pruebas del modo vision, paralelismo, reproceso selectivo, conservacion de
checkpoints, Ctrl+C y fallback tecnico se ejecutan con:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_paralelo_vision -v
```

Usan un servidor falso compatible con LM Studio, asi que no requieren GPU. La
prueba de Ctrl+C envia SIGINT y solo corre en Linux/macOS; en Windows se omite y
conviene verificarla a mano una vez (Ctrl+C durante una corrida con
`--paralelo 4` debe terminar en pocos segundos). No sustituyen una prueba real contra LM
Studio/Ollama, porque las respuestas del modelo y la lectura de anexos dependen
del entorno.

## Guia para otro agente o desarrollador

Para entender o modificar el proyecto sin recorrer todos los datos descargados,
leer en este orden:

1. `README_PROYECTO.md`: arquitectura, operacion y limitaciones.
2. `prompts_extraccion.py`: contrato que debe cumplir el modelo.
3. `3_extraer_ia.py`: lectura documental, inferencia, validacion y reportes.
4. `tests/test_extractor_ia.py`: comportamiento esperado de las reglas criticas.
5. `automatizar_pipeline.py`: coordinacion mensual y reanudacion.
6. `2_scraper_ofertas.py`: contrato de carpetas y descarga de anexos.
7. `1_filtrar_licitaciones.py`: seleccion inicial del universo de licitaciones.
8. `3_extraer_openai_prueba.py`: implementacion comparativa, no principal.

Invariantes que una modificacion debe preservar:

- nunca inventar marca, cantidad o precio;
- no asignar el total general de una oferta a un producto;
- mantener separados `precio_total_tipo=linea` y `oferta`;
- no perder filas al fragmentar documentos o lotes JSON;
- conservar evidencia, archivo y fuente de cada valor;
- excluir productos fuera del alcance incluso si el modelo los devuelve;
- guardar checkpoint despues de cada proveedor;
- permitir reanudar sin repetir proveedores correctos;
- mantener la validacion determinista despues de la respuesta del modelo.

Deuda tecnica y mejoras posibles:

- evaluar sobre un lote real la precision de `--vision` frente a solo texto;
- reordenar los prompts (reglas fijas primero, datos variables al final) para
  aprovechar la cache de prefijo de LM Studio;
- reducir llamadas agrupando paginas relevantes por proveedor;
- llevar la validacion avanzada del extractor local al comparador OpenAI;
- agregar un dashboard para analizar marcas, precios y participacion por mes.
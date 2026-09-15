# Inteligencia Competitiva de Licitaciones - Mercado Publico (Chile)

## Objetivo

Identificar las marcas, productos y precios ofertados en licitaciones publicas
de computo para comparar a HP con sus competidores.

## Flujo vigente

```text
CSV descargado desde el buscador de Mercado Publico
        |
        v
ListaLicitaciones2026.csv
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
ofertas/<codigo>/extraccion_ia.json + ofertas/resultado_productos.xlsx
```

## Componentes conservados

- `1_filtrar_licitaciones.py`: filtra el CSV del buscador por computo y por
  estados Publicada, Cerrada y Adjudicada. Genera los dos CSV intermedios.
- `2_scraper_ofertas.py`: lee `para_scrapear.csv`, descarga anexos tecnicos y
  economicos, y actualiza `scraping_estado` para poder reanudar una corrida.
- `3_extraer_ia.py`: extrae texto de anexos, procesa los hallazgos con un
  modelo local de Ollama y guarda datos estructurados por licitacion.
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
`http://127.0.0.1:11434`. Genera `extraccion_ia.json` dentro de la licitacion y
un Excel con las mismas columnas principales del reporte OpenAI, incluyendo
nombre, fecha de publicacion, estado y organismo. `--sin-consolidar` hace que
la comparacion sea mas pareja, porque analiza los archivos sin una segunda
llamada de consolidacion.

Los extractores filtran el resultado a productos relevantes para computo:
notebooks, desktops/PC, AIO, workstations, monitores, impresoras y
consumibles de impresion. RAM, SSD, procesadores, puertos, sistemas operativos,
servidores, switches y servicios quedan fuera. Teclados, mouse y docks solo se
conservan cuando acompañan a un equipo relevante de la misma oferta.

Para regenerar un Excel OpenAI ya existente con este filtro, sin volver a llamar
a la API, ejecuta el extractor sin `--rehacer` y usa otro nombre de salida:

```powershell
.\.venv\Scripts\python.exe .\3_extraer_openai_prueba.py `
  --dir ".\lotes\2025-11\ofertas" `
  --modelo "gpt-4.1-mini" `
  --limite-proveedores 0 `
  --excel ".\lotes\2025-11\resultados_openai_filtrados.xlsx"
```

Probar OpenAI sobre todas las licitaciones ya descargadas:

```powershell
python 3_extraer_openai_prueba.py --dir ofertas --modelo gpt-4.1-mini --limite-proveedores 0 --excel resultados_todas_openai.xlsx
```

La prueba OpenAI cruza cada codigo con `para_scrapear.csv` para incorporar el
nombre y la fecha de publicacion. Genera `extraccion_openai.json` dentro de cada
licitacion y un Excel con las hojas `Productos`, `Resumen` y `Consumo`. Si la
corrida se interrumpe, el mismo comando reanuda los proveedores pendientes. Usa
`--rehacer` solo cuando necesites descartar resultados previos y pagar nuevamente
las llamadas ya realizadas.

## Automatizacion por mes

`automatizar_pipeline.py` ejecuta el flujo completo con una sola orden. Usa la
fecha de publicacion para separar las licitaciones, reinicia el scraper entre
meses y conserva el estado de cada codigo para poder reanudar una corrida.

La descarga principal obtiene anexos economicos y tecnicos juntos. Esto permite
relacionar cantidad, especificaciones y precio aunque esten repartidos en varios
documentos del mismo proveedor. Los documentos administrativos siguen fuera,
salvo que se use `--todos-anexos`.

Primero se puede revisar la distribucion sin abrir el navegador ni llamar a
OpenAI:

```powershell
.\.venv\Scripts\python.exe .\automatizar_pipeline.py `
  --in ".\ListaLicitaciones2026.csv" `
  --salida ".\lotes" `
  --solo-preparar
```

Para ejecutar filtro, scraping, extraccion OpenAI y reporte consolidado:

```powershell
.\.venv\Scripts\python.exe .\automatizar_pipeline.py `
  --in ".\ListaLicitaciones2026.csv" `
  --salida ".\lotes" `
  --modelo "gpt-4.1-mini"
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
    resultados_openai.xlsx
  2025-12/
    ...
```

Si el proceso se interrumpe, se ejecuta la misma orden. El scraper omite las
licitaciones con estado `ok` y el extractor OpenAI omite los proveedores ya
guardados en `extraccion_openai.json`. Para volver a intentar errores del
scraper se agrega `--reintentar-errores`. No se debe usar `--rehacer-ia` salvo
que se quiera repetir y pagar nuevamente toda la extraccion.

La primera pasada manual tambien puede ejecutarse por separado:

```powershell
python 2_scraper_ofertas.py --csv para_scrapear.csv --incluye-tecnicos
python 3_extraer_openai_prueba.py --dir ofertas --limite-proveedores 0
```

Para completar tecnicos de un lote que ya descargo economicos, sin volver a
descargar los economicos:

```powershell
python 2_scraper_ofertas.py --csv para_scrapear.csv --solo-tecnicos
python 3_extraer_openai_prueba.py --dir ofertas --limite-proveedores 0 --rehacer
```

## Notas operativas

- Los anexos de ofertas se obtienen por scraping web; la API de Mercado Publico
  no expone esos adjuntos de forma util para este flujo.
- Solo las licitaciones Cerrada y Adjudicada son scrapeables. Las Publicadas se
  mantienen en `licitaciones_computo.csv`, pero no pasan a `para_scrapear.csv`.
- El scraper guarda el estado despues de cada licitacion. No borres
  `para_scrapear.csv` ni `ofertas/` si necesitas reanudar el trabajo.
- Los PDF sin texto quedan registrados como `necesita_ocr` en la extraccion.
- La etapa de dashboard consolidado para este flujo aun no esta implementada.
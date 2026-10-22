# Licitaciones de cómputo (Mercado Público, Chile)

Pipeline de inteligencia competitiva: descarga las ofertas de licitaciones de equipos de cómputo, extrae los
productos ofertados con un modelo local (Qwen en LM Studio) más reglas deterministas y genera un Excel con
ID y nombre de la licitación, fecha, proveedor, producto, tipo, marca, cantidad, precio unitario, precio total,
fuente de información, estado de validación y cobertura documental.

Documentación completa: [README_PROYECTO.md](README_PROYECTO.md). Resumen para otra IA o persona:
[BRIEFING_PARA_IA.md](BRIEFING_PARA_IA.md).

## Qué hay en el repositorio

| Archivo | Para qué sirve |
|---|---|
| `1_filtrar_licitaciones.py` | Filtra las licitaciones relevantes de un listado y genera `para_scrapear.csv` |
| `2_scraper_ofertas.py` | Descarga ofertas y anexos con Playwright (`--reparar` completa lo que falte) |
| `3_extraer_ia.py` | Extrae productos con LM Studio + reglas y genera el Excel |
| `4_revisar_extraccion.py` | HTML de revisión: de qué documento salió cada dato y si se mezcló con otro |
| `5_auditar_extraccion.py` | Auditoría sin IA: relee los documentos y verifica cada producto |
| `cobertura_documental.py` | Distingue datos que el proveedor no entregó de datos que el pipeline no encontró |
| `evaluar_calidad.py` | Mide una corrida contra valores verificados a mano (`calidad/gold_*`) |
| `automatizar_pipeline.py` | Orquesta todo el proceso por mes |
| `prompts_extraccion.py` | Prompts del modelo |
| `tests/` | Pruebas unitarias |
| `calidad/` | Gold de desarrollo y ciego, y scripts para repetir la prueba de calidad |

Los datos (`ofertas/`, `lotes_*/`, Excel, PDF, logs) no se versionan.

## Instalación en otra máquina (por ejemplo la Z8)

1. Clonar el repositorio y crear el entorno:

   ```powershell
   git clone <url-del-repositorio> Licitaciones
   cd Licitaciones
   python -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   .\.venv\Scripts\python.exe -m playwright install chromium
   ```

2. Herramientas externas:
   - **Tesseract OCR** con el idioma español (`spa`), para los PDF escaneados (`--ocr`).
   - **antiword** para los `.doc` antiguos. Viene con Git for Windows (`mingw64\bin\antiword.exe`).
   - **7-Zip** (opcional), para abrir `.rar` y `.7z`.
   - **LM Studio** con el modelo `qwen/qwen3.6-35b-a3b` cargado y *Max Concurrent Predictions = 4*,
     escuchando en `http://127.0.0.1:1234`.

3. Verificar la instalación:

   ```powershell
   .\.venv\Scripts\python.exe -m unittest discover -s tests
   ```

## Uso mínimo

```powershell
# 1. Descargar las ofertas de un mes (reanudable; --reparar completa lo que falte)
.\.venv\Scripts\python.exe .\2_scraper_ofertas.py --csv .\lotes\2025-11\para_scrapear.csv --salida .\lotes\2025-11\ofertas

# 2. Extraer (configuración recomendada)
.\.venv\Scripts\python.exe .\3_extraer_ia.py --dir .\lotes\2025-11\ofertas --backend lmstudio `
  --modelo "qwen/qwen3.6-35b-a3b" --metadata-csv .\lotes\2025-11\para_scrapear.csv `
  --excel .\lotes\2025-11\resultado.xlsx --por-proveedor --max-chars-proveedor 24000 `
  --paralelo 4 --max-tokens 4096 --timeout 600 --reintentos-modelo 2 --ocr

# 3. Auditar y revisar
.\.venv\Scripts\python.exe .\5_auditar_extraccion.py --dir .\lotes\2025-11\ofertas --excel .\lotes\2025-11\auditoria.xlsx
.\.venv\Scripts\python.exe .\4_revisar_extraccion.py --dir .\lotes\2025-11\ofertas --fuente ia
```

## Corrida desatendida (fin de semana)

1. Actualizar el código y verificar el entorno:

   ```powershell
   git pull
   .\.venv\Scripts\python.exe .\verificar_entorno.py --modelo "qwen/qwen3.6-35b-a3b" --salida D:\lotes
   ```

   Debe terminar sin errores. Avisos frecuentes y su solución:
   - **OCR sin español:** `.\.venv\Scripts\python.exe .\verificar_entorno.py --preparar-tessdata` (deja `tessdata\` en el proyecto; no requiere administrador).
   - **Paralelismo ~1:** en LM Studio, al cargar el modelo, subir *Max Concurrent Predictions* a 4. Con 1 la corrida tarda 3 a 4 veces más.
2. Lanzar todo:

   ```powershell
   .\correr_fin_de_semana.ps1 -Catalogo .\ListaLicitaciones.csv -Salida D:\lotes
   ```

   Filtra, descarga (mientras extrae el mes anterior), repara descargas incompletas, extrae, reintenta los
   proveedores que fallaron, audita cada mes y deja `D:\lotes\resultados_consolidados.xlsx`. Si algo se corta
   (LM Studio, internet), espera y retoma solo; volver a ejecutar el mismo comando nunca repite lo que ya está bien.
   El equipo no se suspende mientras corre.
3. Revisar al terminar: `D:\lotes\fin_de_semana.log`, `D:\lotes\pipeline.log` y, por mes,
   `descargas_faltantes.csv` (anexos que el portal ofrecía y no se pudieron bajar) y `auditoria.xlsx`.

## Corrida de varios meses (año completo)

`correr_anio.ps1` recorre todos los meses de una carpeta de datos (`<Datos>\<YYYY-MM>\{para_scrapear.csv, ofertas}`, también bajo `lotes_*\`) y por cada mes repara las descargas, extrae con LM Studio y audita sin IA. Es reanudable: si se corta, se vuelve a ejecutar el mismo comando.

```powershell
.\correr_anio.ps1 -Datos D:\licitaciones_datos
.\correr_anio.ps1 -Datos D:\licitaciones_datos -Desde 2025-11 -Hasta 2025-11 -SinScraping   # un solo mes
```

Deja en cada mes `resultado_<mes>.xlsx` y `auditoria_<mes>.xlsx`, los logs en `<Datos>\logs` y, al final, `resultado_anio.xlsx` (lo une `unir_resultados.py`).

Los detalles, las opciones y los resultados medidos están en [README_PROYECTO.md](README_PROYECTO.md).
Conviene guardar los datos fuera de OneDrive.

# Briefing del proyecto Licitaciones (para otra IA)

## 1. Qué es
Pipeline de inteligencia competitiva sobre **Mercado Público (Chile)**. Objetivo: para licitaciones de equipos de cómputo, obtener qué ofertó cada proveedor en un Excel con estas columnas:
ID y nombre de licitación, fecha, proveedor, producto ofertado, tipo (Notebook, AIO, Desktop, Monitor, Impresora, Workstation...), marca, cantidad, precio unitario, precio total, fuente de información (archivo del que salió el dato), más "Estado validación" y "Alertas".

Meta del usuario: al menos 80 % de fiabilidad (no exige 100 %) y velocidad razonable.

## 2. Arquitectura
| Paso | Script | Función |
|---|---|---|
| 1 | `1_filtrar` | Filtra licitaciones relevantes |
| 2 | `2_scraper` (Playwright) | Descarga ofertas y anexos (PDF, Excel, Word...) por licitación/proveedor |
| 3 | `3_extraer_ia.py` | Lee los documentos, llama a un LLM local y genera el Excel |
| 4 | `4_revisar_extraccion.py` | Genera un HTML de revisión: de qué documento salió cada dato, evidencia resaltada, aviso de mezcla de documentos |
| 5 | `5_auditar_extraccion.py` | Auditoría **sin LLM**: relee los documentos y verifica literalmente cada producto |
| - | `evaluar_calidad.py` | Evalúa una corrida contra un gold hecho a mano (`calidad/gold_desarrollo`, `calidad/gold_ciego`) |
| - | `automatizar_pipeline.py` | Orquesta todo |

- `prompts_extraccion.py` contiene los prompts. Las pruebas están en `tests/` (120 tests, `python -m unittest discover -s tests`).
- **LLM:** LM Studio en `http://127.0.0.1:1234`, conectado por LM Link a una workstation (Z8), modelo `qwen/qwen3.6-35b-a3b`. Los datos están en `lotes_z8` (nov-2025 a abr-2026) y `lotes_z4` (may a sep-2026, modelo 9B). No modificar esos lotes; trabajar sobre copias (`calidad/preparar_copia.py`).
- **Modos de extracción:** `--por-proveedor` (una llamada por proveedor con las páginas relevantes de todos sus archivos) o por archivo (una llamada por archivo y luego consolidación). `--rapido` equivale a `--prompt-compacto` más `--sin-consolidar`.

## 3. Cómo funciona 3_extraer_ia.py (resumen)
1. Convierte cada documento a texto (PDF, Excel, Word). Los escaneados o ilegibles pasan por OCR (`--ocr`).
2. Selecciona los fragmentos relevantes (económicos y técnicos), con topes de caracteres y fragmentos.
3. El LLM devuelve JSON con los productos: producto, marca, tipo, cantidad, unitario, total, archivo fuente.
4. **Reglas deterministas posteriores (lo que más pesa en la calidad):**
   - Rescate de filas económicas verificables (cantidad × unitario = total) que el modelo omitió.
   - Reconstrucción de cifras partidas por espacios ("6 75.620" → 675.620), montos con sufijo ".-".
   - Lógica de IVA (total con IVA → neto; duplicados ×1,19).
   - Conciliación de precios contra anexos económicos (anula precios sin respaldo en el documento).
   - Normalización e inferencia de marca (alias, familia de modelo, descripción, nombre de archivo).
   - Clasificación de alcance y tipo, con rangos de precio plausibles por tipo.
   - Descarte de historial/órdenes de compra pasadas (documentos de "experiencia"), líneas resumen, duplicados entre archivos y fuentes inventadas.
   - Estado: `ok`, `incompleto`, `inconsistente`, `revisar`.

## 4. Problemas encontrados
**Inconsistencias iniciales.** Productos sin precio aunque el documento lo tenía; datos mal leídos que el mismo modelo sí acertaba en el chat de LM Studio. La causa estaba en el código, no en el modelo.

**Causas específicas:**
- El filtro de alcance eliminaba MacBook, IdeaPad y ZBook, y dejaba pasar accesorios, UPS y servidores.
- "Conflictos" falsos de precio entre documentos.
- Duplicados por IVA.
- Dígitos partidos por espacios.
- Marcas mal normalizadas o inventadas.
- Nombres de archivo fuente inventados por el modelo.
- Órdenes de compra pasadas (experiencia previa del proveedor) extraídas como si fueran lo ofertado.
- Líneas resumen extraídas como productos.
- Totales leídos como precio unitario.
- Páginas con texto basura tratadas como si tuvieran texto.

**Entorno / herramientas:**
- Apagar el razonamiento de Qwen en LM Studio exige `reasoning_effort: "none"`. `enable_thinking: false` solo no basta.
- El tiempo lo domina la cantidad de tokens que escribe el modelo. Cuatro llamadas en paralelo dan unas 3 veces el throughput (cargar el modelo con Max Concurrent Predictions = 4).
- La CPU del notebook no debe usarse para trabajo pesado; la inferencia corre en la Z8.
- El proyecto está bajo OneDrive: las corridas grandes y copias de datos van fuera (p. ej. `C:\Users\CaNi517\licitaciones_nov_nuevo`).
- Windows: rutas largas, CRLF frente a LF (los `.sh` deben ir en LF) y la herramienta Bash que consume backslashes.

## 5. Metodología de evaluación
Se construyó un gold a mano: set de desarrollo (10 licitaciones, 102 productos) y set **ciego** (8 licitaciones al azar no usadas para diseñar reglas, 72 productos). Métricas: recall, precisión, exactitud de marca, tipo, cantidad y total, fiabilidad completa (todos los campos correctos) y exactitud de los productos marcados `ok`.

## 6. Resultados
| Prueba | Modo | Recall | Precisión | Todos los campos correctos | Seg/proveedor |
|---|---|---|---|---|---|
| Ciega (72 prod.) | `--por-proveedor` | 98,6 % | 91,0 % | 97,2 % | 44 |
| Ciega (72 prod.) | por archivo `--rapido` | 98,6 % | 82,6 % | 93,1 % | 61 |
| Desarrollo, 4 licitaciones (63 prod.) | `--por-proveedor` | 100 % | 98,4 % | 100 % | 45 |
| Desarrollo, 4 licitaciones (63 prod.) | por archivo completo | 100 % | 98,4 % | 100 % | 113 |
| Desarrollo, 4 licitaciones (63 prod.) | por archivo compacto | 100 % | 98,4 % | 100 % | 87 |
| Desarrollo, 4 licitaciones (63 prod.) | por archivo `--rapido` | 100 % | 85,1 % | 100 % | 57 |

**Configuración recomendada:** `--por-proveedor --max-chars-proveedor 24000 --paralelo 4 --ocr`.

**Corrida completa de noviembre 2025:**
- 763 proveedores en 69 licitaciones, 1 h 54 min.
- 1.128 productos: 821 `ok`, 222 `incompleto`, 77 `inconsistente`, 8 `revisar`.
- Tipos: notebook 474, all-in-one 221, impresora 201, monitor 141, desktop 66, workstation 20.
- La auditoría sin IA verificó literalmente el 88 % de los `ok` (719 de 821).
- La revisión visual marca 663 casos de "precio NO encontrado" y 29 de "descripción de otro producto": son los primeros a revisar.

## 7. Limitaciones conocidas
- Ofertas con un único precio global para varios equipos (kits) no se desglosan.
- Cantidades que solo aparecen en el texto de las bases pueden quedar vacías.
- Los documentos escaneados dependen de la calidad del OCR (39 archivos de noviembre pendientes de OCR).
- Si dos documentos del mismo proveedor traen precios distintos para un mismo equipo, el producto queda `inconsistente`.
- El gold es pequeño (174 productos en total), así que las métricas tienen margen de error.

## 8. Lecciones para quien continúe
- Antes de afirmar una mejora, medir con `evaluar_calidad.py` en el set de desarrollo **y** en el ciego.
- Las reglas de código posteriores al modelo pesan más que el prompt.
- Una sola llamada por proveedor, con pocas páginas relevantes, gana en velocidad y precisión al modo por archivo con consolidación.
- No correr trabajo pesado en el notebook del usuario ni tocar sus apps (explorador_demo) sin permiso.

## 9. Cómo reproducir
- **Comando recomendado:**
```
.\.venv\Scripts\python.exe .\3_extraer_ia.py --dir ".\lotes\2025-11\ofertas" --backend lmstudio --modelo "qwen/qwen3.6-35b-a3b" --metadata-csv ".\lotes\2025-11\para_scrapear.csv" --excel ".\lotes\2025-11\resultado.xlsx" --por-proveedor --max-chars-proveedor 24000 --paralelo 4 --max-tokens 4096 --timeout 600 --reintentos-modelo 2 --ocr
```
- **Repetir la prueba:** copiar licitaciones con `calidad/preparar_copia.py`, correr con `calidad/correr_prueba.sh`, medir con `evaluar_calidad.py --gold calidad/gold_desarrollo`. El detalle está en `README_PROYECTO.md`.
- **Resultados archivados** de cada corrida: `calidad/resultados_prueba/`.

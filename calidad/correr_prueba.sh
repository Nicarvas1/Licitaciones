#!/bin/bash
# Uso: bash calidad/correr_prueba.sh <carpeta> [opciones del extractor...]
# Ejecuta 3_extraer_ia.py sobre pruebas_calidad/<carpeta>/ofertas con la configuracion de prueba
# (LM Studio, qwen/qwen3.6-35b-a3b, 4 proveedores en paralelo, OCR) y guarda el log en la misma carpeta.
# Si se interrumpe, repetir el mismo comando: retoma lo pendiente.
#   bash calidad/correr_prueba.sh A --rapido          # configuracion recomendada
#   bash calidad/correr_prueba.sh A                   # modo completo (consolidacion con el modelo)
D="$1"; shift
cd "$(dirname "$0")/.."
META=pruebas_calidad/para_scrapear.csv
[ -f "pruebas_calidad/$D/para_scrapear.csv" ] && META="pruebas_calidad/$D/para_scrapear.csv"
PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe -u 3_extraer_ia.py --dir "pruebas_calidad/$D/ofertas" \
  --backend lmstudio --modelo "qwen/qwen3.6-35b-a3b" --metadata-csv "$META" \
  --excel "pruebas_calidad/$D/resultado.xlsx" --paralelo 4 --max-tokens 4096 --timeout 600 \
  --reintentos-modelo 2 --ocr "$@" > "pruebas_calidad/$D/corrida.log" 2>&1

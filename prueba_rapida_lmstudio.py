#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path

import pdfplumber
import requests


URL_LMSTUDIO = "http://127.0.0.1:1234/v1/chat/completions"


def extraer_texto_pdf(ruta, max_paginas):
    paginas = []
    with pdfplumber.open(ruta) as pdf:
        for numero, pagina in enumerate(pdf.pages):
            if numero >= max_paginas:
                break
            texto = (pagina.extract_text() or "").strip()
            if texto:
                paginas.append(f"[PAGINA {numero + 1}]\n{texto}")
    return "\n\n".join(paginas)


def limpiar_json(respuesta):
    texto = respuesta.strip()
    if texto.startswith("```"):
        texto = texto.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        inicio = texto.find("{")
        final = texto.rfind("}")
        if inicio >= 0 and final > inicio:
            return json.loads(texto[inicio:final + 1])
        raise


def main():
    parser = argparse.ArgumentParser(description="Prueba minima: un PDF y una llamada a LM Studio")
    parser.add_argument("archivo", type=Path)
    parser.add_argument("--modelo", default="qwen/qwen3.6-35b-a3b")
    parser.add_argument("--max-paginas", type=int, default=10)
    parser.add_argument("--max-caracteres", type=int, default=12000)
    parser.add_argument("--max-tokens", type=int, default=4096)
    parser.add_argument("--timeout", type=int, default=600)
    args = parser.parse_args()

    if not args.archivo.is_file():
        raise SystemExit(f"No existe el archivo: {args.archivo}")

    inicio_lectura = time.perf_counter()
    texto = extraer_texto_pdf(args.archivo, args.max_paginas)[:args.max_caracteres]
    segundos_lectura = time.perf_counter() - inicio_lectura
    if not texto.strip():
        raise SystemExit("El PDF no contiene texto extraible; probablemente necesita OCR.")

    prompt = f"""Analiza este documento de una oferta comercial.
Devuelve SOLO un objeto JSON valido con esta estructura:
{{
  "productos": [
    {{
      "producto": "descripcion concreta",
      "marca": "marca o null",
      "modelo": "modelo o null",
      "cantidad": "numero o null",
      "precio_unitario": "numero o null",
      "precio_total": "numero o null"
    }}
  ],
  "observaciones": "texto breve o null"
}}
Extrae los productos realmente ofertados. No incluyas especificaciones como productos separados.

DOCUMENTO:
{texto}
"""

    payload = {
        "model": args.modelo,
        "messages": [
            {"role": "system", "content": "Devuelve exclusivamente un objeto JSON valido."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1,
        "max_tokens": args.max_tokens,
        "reasoning_effort": "none",
        "stream": False,
        "response_format": {"type": "text"},
        "chat_template_kwargs": {"enable_thinking": False}
    }

    print(f"Archivo: {args.archivo}")
    print(f"Texto enviado: {len(texto):,} caracteres")
    print(f"Lectura PDF: {segundos_lectura:.2f} segundos")
    print("Consultando LM Studio...")
    inicio_modelo = time.perf_counter()
    respuesta = requests.post(URL_LMSTUDIO, json=payload, timeout=args.timeout)
    segundos_modelo = time.perf_counter() - inicio_modelo
    respuesta.raise_for_status()

    datos_respuesta = respuesta.json()
    contenido = ((datos_respuesta.get("choices") or [{}])[0]
                 .get("message", {}).get("content") or "")
    print(f"Respuesta LM Studio: {segundos_modelo:.2f} segundos")
    print(f"Tokens usados: {datos_respuesta.get('usage', {})}")
    print("\nRespuesta original:\n")
    print(contenido)

    try:
        datos = limpiar_json(contenido)
    except json.JSONDecodeError as exc:
        print(f"\nNo se pudo interpretar la respuesta como JSON: {exc}")
        return

    print("\nProductos detectados:")
    for producto in datos.get("productos", []):
        print(json.dumps(producto, ensure_ascii=False))


if __name__ == "__main__":
    main()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diagnostico de velocidad: por que el chat de LM Studio es rapido y el extractor no.

Procesa uno o mas archivos (idealmente el mismo documento que en el chat anda rapido)
de varias formas y mide, por cada una: llamadas, segundos, tokens de entrada, tokens
de salida y si el modelo "penso" (razonamiento oculto) antes de responder.

Formas:
  chat          -> como el chat: todo el texto en UNA llamada, pregunta corta,
                   sin parametros extra.
  chat_params   -> igual, pero con los parametros que usa el extractor para
                   desactivar el razonamiento (para ver si LM Studio los respeta).
  actual        -> exactamente lo que hace 3_extraer_ia.py: fragmentos de
                   --max-chars-archivo, PROMPT_ARCHIVO completo con pistas.
  vision        -> (con --vision, solo PDF) paginas como imagen + pregunta corta.

Uso (LM Studio abierto, servidor iniciado y el modelo cargado):
  python diagnostico_velocidad.py "ruta\\al\\anexo.pdf" --modelo "qwen/qwen3.6-35b-a3b"
  python diagnostico_velocidad.py anexo1.pdf anexo2.xlsx --modelo ... --vision

Genera diagnostico_velocidad.json con el detalle de cada llamada (incluido el
comienzo del razonamiento, si lo hubo).
"""

import argparse
import importlib.util
import json
import sys
import time
from pathlib import Path
from types import SimpleNamespace

import requests

RAIZ_PROYECTO = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ_PROYECTO))
_spec = importlib.util.spec_from_file_location("extractor_ia", RAIZ_PROYECTO / "3_extraer_ia.py")
ia = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ia)

PREGUNTA_CORTA = (
    "Extrae de este documento de oferta los equipos ofertados (computadores, notebooks, all-in-one, "
    "workstations, monitores e impresoras). Para cada uno indica producto, marca, modelo, cantidad, "
    "precio unitario y precio total. Responde solo con JSON: "
    '{"productos":[{"producto":"","marca":"","modelo":"","cantidad":0,"precio_unitario":0,"precio_total":0}]}'
)
PARAMETROS_SIN_RAZONAR = {
    "reasoning_effort": "none",
    "chat_template_kwargs": {"enable_thinking": False},
}


def llamar(url, modelo, mensajes, max_tokens, timeout, extra=None):
    """Una llamada a LM Studio. Devuelve un dict con tiempos, tokens y razonamiento."""
    payload = {"model": modelo, "messages": mensajes, "temperature": 0.0,
               "max_tokens": max_tokens, "stream": False, **(extra or {})}
    inicio = time.perf_counter()
    try:
        respuesta = requests.post(url, json=payload, timeout=timeout)
    except requests.exceptions.RequestException as exc:
        return {"error": str(exc), "segundos": round(time.perf_counter() - inicio, 1)}
    segundos = time.perf_counter() - inicio
    if respuesta.status_code >= 400:
        return {"error": f"http {respuesta.status_code}: {respuesta.text[:300]}", "segundos": round(segundos, 1)}
    datos = respuesta.json()
    mensaje = (datos.get("choices") or [{}])[0].get("message", {}) or {}
    contenido = mensaje.get("content") or ""
    razonamiento = mensaje.get("reasoning_content") or mensaje.get("reasoning") or ""
    if "<think>" in contenido:  # algunas plantillas dejan el razonamiento dentro del contenido
        razonamiento += contenido.split("</think>")[0]
        contenido = contenido.split("</think>")[-1]
    uso = datos.get("usage") or {}
    parseado, estado = ia.parsear_json(contenido)
    productos = len(parseado.get("productos", [])) if isinstance(parseado, dict) else 0
    salida = uso.get("completion_tokens") or 0
    return {
        "segundos": round(segundos, 1),
        "tokens_entrada": uso.get("prompt_tokens"),
        "tokens_salida": salida,
        "tokens_por_segundo": round(salida / segundos, 1) if segundos and salida else None,
        "razono": bool(razonamiento.strip()),
        "caracteres_razonamiento": len(razonamiento),
        "caracteres_respuesta": len(contenido),
        "json_valido": estado in ia.ESTADOS_IA_OK,
        "productos": productos,
        "inicio_razonamiento": razonamiento[:1500],
        "inicio_respuesta": contenido[:1500],
    }


def args_lectura(max_paginas):
    return SimpleNamespace(max_paginas=max_paginas, ocr=False, tesseract_cmd=None, idioma_ocr="spa+eng",
                           dpi_ocr=180, vision=False, vision_min_filas=2)


def forma_chat(url, modelo, texto, args, con_parametros):
    mensajes = [{"role": "user", "content": f"{PREGUNTA_CORTA}\n\nDOCUMENTO:\n{texto[:args.max_chars_chat]}"}]
    return [llamar(url, modelo, mensajes, args.max_tokens, args.timeout,
                   PARAMETROS_SIN_RAZONAR if con_parametros else None)]


def forma_actual(url, modelo, path, texto, args):
    """Replica 3_extraer_ia.py: fragmentos + PROMPT_ARCHIVO + mismo payload."""
    llamadas = []
    fragmentos = ia.dividir_texto(texto, args.max_chars_archivo)
    for numero, fragmento in enumerate(fragmentos, 1):
        productos_pista, precios_pista = ia.pistas(fragmento)
        prompt = ia.PROMPT_ARCHIVO.format(
            proveedor="(diagnostico)", rut="(diagnostico)", archivo=path.name,
            tipo_documental=ia.tipo_documental(path.name), total_oferta="no informado",
            pistas_producto="\n".join(productos_pista) or "(ninguna)",
            pistas_precio="\n".join(precios_pista) or "(ninguna)",
            texto=f"[FRAGMENTO {numero} DE {len(fragmentos)}]\n{fragmento}",
        )
        mensajes = [{"role": "system", "content": "Devuelve exclusivamente un objeto JSON valido."},
                    {"role": "user", "content": prompt}]
        llamadas.append(llamar(url, modelo, mensajes, args.max_tokens, args.timeout,
                               {**PARAMETROS_SIN_RAZONAR, "response_format": {"type": "text"}}))
    return llamadas


def forma_vision(url, modelo, path, args):
    with ia.FITZ_LOCK:
        documento = ia.fitz.open(path)
        paginas = list(range(min(len(documento), args.vision_paginas)))
        documento.close()
    imagenes = ia.renderizar_paginas_pdf(path, paginas, args.dpi)
    contenido = [{"type": "image_url", "image_url": {"url": f"data:{img['mime']};base64,{img['base64']}"}}
                 for img in imagenes.values()] + [{"type": "text", "text": PREGUNTA_CORTA}]
    return [llamar(url, modelo, [{"role": "user", "content": contenido}], args.max_tokens, args.timeout,
                   PARAMETROS_SIN_RAZONAR)]


def resumir(llamadas):
    validas = [l for l in llamadas if "error" not in l]
    return {
        "llamadas": len(llamadas),
        "errores": len(llamadas) - len(validas),
        "segundos": round(sum(l.get("segundos", 0) for l in llamadas), 1),
        "tokens_entrada": sum(l.get("tokens_entrada") or 0 for l in validas),
        "tokens_salida": sum(l.get("tokens_salida") or 0 for l in validas),
        "razono": any(l.get("razono") for l in validas),
        "caracteres_razonamiento": sum(l.get("caracteres_razonamiento", 0) for l in validas),
        "productos": sum(l.get("productos", 0) for l in validas),
    }


def main():
    parser = argparse.ArgumentParser(description="Diagnostico de velocidad del extractor IA.")
    parser.add_argument("archivos", nargs="+", help="Archivo(s) a medir (PDF, XLSX, DOCX...)")
    parser.add_argument("--modelo", required=True)
    parser.add_argument("--url", default=ia.LMSTUDIO_URL)
    parser.add_argument("--max-chars-archivo", type=int, default=7000, help="Igual que en 3_extraer_ia.py")
    parser.add_argument("--max-chars-chat", type=int, default=60000)
    parser.add_argument("--max-tokens", type=int, default=4096)
    parser.add_argument("--max-paginas", type=int, default=20)
    parser.add_argument("--timeout", type=int, default=900)
    parser.add_argument("--vision", action="store_true", help="Agrega la forma con paginas como imagen (PDF)")
    parser.add_argument("--vision-paginas", type=int, default=4)
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--salida", default="diagnostico_velocidad.json")
    args = parser.parse_args()

    try:
        requests.get(args.url.replace("/chat/completions", "/models"), timeout=5)
    except requests.exceptions.RequestException:
        sys.exit(f"No hay respuesta en {args.url}. Inicia el servidor de LM Studio (Developer > Start Server).")

    detalle = {}
    totales = {}
    for nombre in args.archivos:
        path = Path(nombre)
        if not path.is_file():
            print(f"No existe: {path}")
            continue
        texto, estado = ia.extraer_archivo(path, args_lectura(args.max_paginas))
        print(f"\n{path.name}: {len(texto):,} caracteres de texto ({estado})")
        formas = {}
        if texto.strip():
            print("  midiendo 'chat'...", flush=True)
            formas["chat"] = forma_chat(args.url, args.modelo, texto, args, con_parametros=False)
            print("  midiendo 'chat_params'...", flush=True)
            formas["chat_params"] = forma_chat(args.url, args.modelo, texto, args, con_parametros=True)
            print("  midiendo 'actual' (como 3_extraer_ia.py)...", flush=True)
            formas["actual"] = forma_actual(args.url, args.modelo, path, texto, args)
        if args.vision and path.suffix.lower() == ".pdf":
            print("  midiendo 'vision'...", flush=True)
            formas["vision"] = forma_vision(args.url, args.modelo, path, args)
        detalle[path.name] = formas
        for forma, llamadas in formas.items():
            totales.setdefault(forma, []).extend(llamadas)

    if not totales:
        sys.exit("No se midio nada.")
    Path(args.salida).write_text(json.dumps(detalle, ensure_ascii=False, indent=1), encoding="utf-8")

    print("\n" + "=" * 96)
    print(f"{'forma':12} {'llamadas':>8} {'segundos':>9} {'tok.entrada':>12} {'tok.salida':>11} "
          f"{'razono':>7} {'car.razon.':>11} {'productos':>9}")
    resumenes = {forma: resumir(llamadas) for forma, llamadas in totales.items()}
    for forma, r in resumenes.items():
        print(f"{forma:12} {r['llamadas']:>8} {r['segundos']:>9} {r['tokens_entrada']:>12,} {r['tokens_salida']:>11,} "
              f"{'SI' if r['razono'] else 'no':>7} {r['caracteres_razonamiento']:>11,} {r['productos']:>9}"
              + (f"   ({r['errores']} con error)" if r["errores"] else ""))
    print("=" * 96)

    # Interpretacion automatica
    actual, chat = resumenes.get("actual"), resumenes.get("chat")
    print("\nLECTURA DEL RESULTADO:")
    if actual and actual["razono"]:
        print("- El modelo RAZONA en el flujo del extractor: los parametros para desactivarlo no estan")
        print("  surtiendo efecto. Es la causa mas probable de la lentitud. Hay que desactivar el")
        print("  razonamiento en la configuracion del modelo en LM Studio (no solo en el chat).")
    elif actual:
        print("- El modelo no razona en el flujo del extractor (bien).")
    if chat and resumenes.get("chat_params") and chat["razono"] and not resumenes["chat_params"]["razono"]:
        print("- Sin parametros el modelo razona y con ellos no: los parametros funcionan.")
    if actual and chat and chat["segundos"]:
        print(f"- El flujo actual tarda {actual['segundos'] / chat['segundos']:.1f} veces lo que tarda la forma 'chat' "
              f"({actual['llamadas']} llamada(s) vs {chat['llamadas']}).")
        if actual["tokens_salida"] > 2 * max(1, chat["tokens_salida"]):
            print(f"- La salida del flujo actual es {actual['tokens_salida'] / max(1, chat['tokens_salida']):.1f} veces "
                  "mas larga: el formato de respuesta (15 campos por producto, evidencia, etc.) cuesta tiempo.")
        if actual["llamadas"] > chat["llamadas"]:
            print("- El flujo actual divide el documento en varias llamadas (fragmentos).")
        if actual["productos"] != chat["productos"]:
            print(f"- Productos devueltos: actual {actual['productos']} (sumados por fragmento, antes de deduplicar), "
                  f"chat {chat['productos']}. Revisa en el JSON que la forma rapida no pierda productos.")
    print(f"\nDetalle por llamada: {args.salida}")


if __name__ == "__main__":
    main()

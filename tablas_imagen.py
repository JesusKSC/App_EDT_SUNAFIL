# -*- coding: utf-8 -*-
"""
==============================================================================
TABLAS LEÍDAS DESDE LA IMAGEN
------------------------------------------------------------------------------
El OCR de Tesseract destroza las tablas (ej. Bases p. 32: la Tabla N° 01 de
equipos, GB y minutos queda sin cifras). Para las páginas escaneadas que tienen
una tabla, se envía la IMAGEN de la página a Gemini (PNG, zoom 3) y se pide la
transcripción con las tablas en JSON. Ese texto reemplaza al OCR de la página
antes del análisis. Cada respuesta queda en la caché de gemini_client.
==============================================================================
"""
import re
import pymupdf
import lectura_pdf
from gemini_client import llamar_json, ErrorGemini

DOCS_CON_TABLAS = ("Bases", "Pliego", "Contrato")   # agrega "Propuesta" si también lo necesitas
ZOOM = 3
PAGINAS_POR_LLAMADA = 3                  # varias páginas por llamada ahorran cuota de Gemini
MAX_BYTES_LLAMADA = 12 * 1024 * 1024     # el pedido a Gemini admite hasta ~20 MB en base64

# Señales de que la página tiene una tabla (Tabla N° 01, Cuadro N°, Gama A, columnas en GB...)
_MARCA_TABLA = re.compile(
    r"\b(tabla|cuadro)\s*(n\s*[°ºo*.:]?\s*\d|n\s*[°ºo*.:]|nro|\d)|\bgama\s+[abc]\b|\((gb|mb|minutos|soles|s/)\)"
    r"|\bitem\b.*\bcantidad\b|\bcantidad\b.*\bprecio\b", re.I)
_MARCA_PAGINA = re.compile(r"^\[([^\]]+), p\. (\d+)\]\n", re.M)

PROMPT_TABLAS = """Eres un analista de documentos de contrataciones públicas del Perú.
Te envío las imágenes de las páginas {paginas} del documento «{documento}», en ese orden.
Extrae SOLO las tablas (no copies los párrafos). Por cada página:
- tablas: cada tabla con su título, sus columnas y sus filas, con el valor de cada celda
  EXACTAMENTE como se ve (cifras, unidades, guiones). Si una fila ocupa varias líneas,
  junta su contenido en la misma fila. Si una celda no se puede leer, escribe "ilegible".
  No calcules, no completes y no corrijas ninguna cifra.
Responde solo con JSON válido:
{{"paginas": [{{"pagina": 32,
               "tablas": [{{"titulo": "Tabla N° 01: ...", "columnas": ["..."],
                           "filas": [["..."]], "notas": ""}}]}}]}}"""


def _por_pagina(texto):
    """{n.° de página: texto} a partir de las marcas "[Doc, p. N]"."""
    partes = _MARCA_PAGINA.split(texto or "")
    # partes = [antes, doc, pag, texto, doc, pag, texto, ...]
    return {int(partes[k + 1]): partes[k + 2] for k in range(1, len(partes) - 2, 3)}


def paginas_con_tablas(file_bytes, texto_doc):
    """Páginas leídas por OCR (no digitales) cuyo texto muestra una tabla."""
    h = lectura_pdf.hash_archivo(file_bytes)
    elegidas = []
    for n, txt in _por_pagina(texto_doc).items():
        dato = lectura_pdf._leer_cache(h, n)
        if dato and dato["metodo"] == "ocr" and _MARCA_TABLA.search(txt):
            elegidas.append(n)
    return sorted(elegidas)


def imagen_png(file_bytes, n):
    """Página n (1-based) como PNG en escala de grises (más liviano, igual de legible)."""
    with pymupdf.open(stream=file_bytes, filetype="pdf") as doc:
        pix = doc[n - 1].get_pixmap(matrix=pymupdf.Matrix(ZOOM, ZOOM),
                                    colorspace=pymupdf.csGRAY, alpha=False)
        return pix.tobytes("png")


def _grupos(paginas, imagenes):
    """Agrupa páginas consecutivas de la lista sin pasar el tamaño máximo por llamada."""
    grupos, actual, peso = [], [], 0
    for n in paginas:
        tam = len(imagenes[n])
        if actual and (len(actual) == PAGINAS_POR_LLAMADA or peso + tam > MAX_BYTES_LLAMADA):
            grupos.append(actual); actual, peso = [], 0
        actual.append(n); peso += tam
    if actual:
        grupos.append(actual)
    return grupos


def _celda(v):
    return str(v if v is not None else "").replace("|", "/").replace("\n", " ").strip()


def tabla_a_texto(t):
    """Tabla en formato Markdown (así la leen los pasos siguientes)."""
    cols = [_celda(c) for c in t.get("columnas") or []]
    filas = [[_celda(c) for c in f] for f in t.get("filas") or [] if isinstance(f, list)]
    ancho = max([len(cols)] + [len(f) for f in filas] + [1])
    cols = cols + [""] * (ancho - len(cols))
    lineas = [f"{_celda(t.get('titulo')) or 'Tabla'}",
              "| " + " | ".join(cols) + " |", "|" + "---|" * ancho]
    lineas += ["| " + " | ".join(f + [""] * (ancho - len(f))) + " |" for f in filas]
    if t.get("notas"):
        lineas.append(f"Notas: {_celda(t['notas'])}")
    return "\n".join(lineas)


def tareas_de_lectura(api_key, textos, pdfs, modelos):
    """Prepara las llamadas a Gemini (una por grupo de páginas).
    Devuelve [(documento, [páginas], función(al_esperar) -> ({página: texto nuevo}, error))]."""
    tareas = []
    for doc in DOCS_CON_TABLAS:
        if not textos.get(doc) or not pdfs.get(doc):
            continue
        paginas = paginas_con_tablas(pdfs[doc], textos[doc])
        if not paginas:
            continue
        imagenes = {n: imagen_png(pdfs[doc], n) for n in paginas}
        for grupo in _grupos(paginas, imagenes):
            prompt = PROMPT_TABLAS.format(paginas=", ".join(map(str, grupo)), documento=doc)

            # todo lo que usa la tarea va como argumento por defecto: se ejecuta después,
            # en otro hilo, cuando el bucle ya pasó al siguiente documento
            def leer(al_esperar, doc=doc, grupo=grupo, prompt=prompt, imagenes=imagenes):
                """Devuelve ({página: texto nuevo}, error). Si Gemini no pudo leer las
                imágenes, esas páginas se quedan con el OCR: el checklist no se detiene."""
                try:   # solo el modelo elegido: es una mejora opcional, no vale gastar la cuota de los respaldos
                    res = llamar_json(api_key, prompt, "", al_esperar, (modelos or [])[:1] or None,
                                      imagenes=[imagenes[n] for n in grupo])
                except ErrorGemini as e:
                    return {}, str(e)
                if isinstance(res, list):
                    res = {"paginas": res}
                nuevos = {}
                for p in res.get("paginas") or []:
                    if not isinstance(p, dict) or p.get("pagina") not in grupo:
                        continue
                    tablas = [tabla_a_texto(t) for t in p.get("tablas") or [] if isinstance(t, dict)]
                    if tablas:
                        nuevos[p["pagina"]] = "\n\n".join(tablas)
                return nuevos, None
            tareas.append((doc, grupo, leer))
    return tareas


def _celda_texto(v):
    return str(v or "").strip()


def reemplazar_paginas(texto_doc, documento, nuevos):
    """Pone primero las tablas leídas desde la imagen; el texto OCR de la página queda
    debajo, marcado como poco confiable para cifras (sus tablas salen destrozadas)."""
    if not nuevos:
        return texto_doc
    paginas = _por_pagina(texto_doc)
    for n, tablas in nuevos.items():
        if n in paginas:
            paginas[n] = ("TABLAS LEÍDAS DESDE LA IMAGEN DE LA PÁGINA (usa estas cifras):\n\n"
                          f"{tablas}\n\nTEXTO OCR DEL RESTO DE LA PÁGINA (sus cifras de tablas "
                          f"no son confiables; si difieren, valen las tablas de arriba):\n{paginas[n]}")
    return "\n\n".join(f"[{documento}, p. {n}]\n{t.strip()}" for n, t in paginas.items())

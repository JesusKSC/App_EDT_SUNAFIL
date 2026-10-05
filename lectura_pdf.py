# -*- coding: utf-8 -*-
"""
==============================================================================
LECTURA DE PDF  (texto digital + OCR)
------------------------------------------------------------------------------
Por cada página: si trae texto digital (> MIN_CARACTERES) se usa directo;
si no (página escaneada), se renderiza a imagen y se pasa por Tesseract (spa).
Cada página leída se guarda en .cache/ (hash del archivo + n.° de página) para
no repetir el OCR la próxima vez.
==============================================================================
"""
import os, re, json, shutil, hashlib, subprocess
import pymupdf

# Ruta del ejecutable de Tesseract (cámbiala si lo instalaste en otro lugar)
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if not os.path.exists(TESSERACT_CMD):   # Linux / servidor (p. ej. Streamlit Community Cloud)
    TESSERACT_CMD = shutil.which("tesseract") or TESSERACT_CMD

MIN_CARACTERES = 80   # menos que esto = página sin texto digital -> OCR
ZOOM_OCR = 2.5        # resolución del render para OCR
IDIOMA_OCR = "spa"
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cache")

# ---------------------------------------------------------------------------
# 1. RANGO DE PÁGINAS  ("24-48, 58-66"; vacío = todas)
# ---------------------------------------------------------------------------
_TOKEN = re.compile(r"^(\d+)(?:\s*-\s*(\d+))?$")

def validar_rango(texto):
    """Valida el formato. Devuelve (lista de (ini, fin), mensaje_error).
    Lista vacía = todas las páginas."""
    texto = (texto or "").strip()
    if not texto:
        return [], None
    rangos = []
    for parte in texto.split(","):
        parte = parte.strip()
        m = _TOKEN.match(parte)
        if not m:
            return None, (f"«{parte}» no es válido. Usa números y rangos separados "
                          f"por comas, por ejemplo: 24-48, 58-66")
        ini = int(m.group(1)); fin = int(m.group(2) or ini)
        if ini < 1:
            return None, "Las páginas empiezan en 1."
        if ini > fin:
            return None, f"En «{parte}» el inicio es mayor que el fin."
        rangos.append((ini, fin))
    return rangos, None

def paginas_a_leer(rangos, total):
    """Convierte los rangos (1-based) en la lista de páginas a leer.
    Lanza ValueError si alguna página no existe en el PDF."""
    if not rangos:
        return list(range(1, total + 1))
    paginas = []
    for ini, fin in rangos:
        if fin > total:
            raise ValueError(f"el PDF tiene {total} páginas y pediste hasta la {fin}")
        paginas.extend(range(ini, fin + 1))
    return sorted(set(paginas))

def contar_paginas(file_bytes):
    with pymupdf.open(stream=file_bytes, filetype="pdf") as doc:
        return doc.page_count

# ---------------------------------------------------------------------------
# 2. CACHÉ EN DISCO
# ---------------------------------------------------------------------------
def hash_archivo(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()

def _ruta_cache(h, n):
    return os.path.join(CACHE_DIR, f"{h}_p{n}.json")

def _leer_cache(h, n):
    try:
        with open(_ruta_cache(h, n), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None

def _guardar_cache(h, n, dato):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(_ruta_cache(h, n), "w", encoding="utf-8") as f:
        json.dump(dato, f, ensure_ascii=False)

# ---------------------------------------------------------------------------
# 3. LECTURA PÁGINA A PÁGINA
# ---------------------------------------------------------------------------
def _ocr(pagina):
    """OCR llamando directo a tesseract.exe (la imagen va por stdin, el texto sale por
    stdout). No se usa pytesseract: importa pandas, y pandas carga una DLL de pyarrow
    que Smart App Control de Windows bloquea."""
    pix = pagina.get_pixmap(matrix=pymupdf.Matrix(ZOOM_OCR, ZOOM_OCR), alpha=False)
    try:
        r = subprocess.run([TESSERACT_CMD, "stdin", "stdout", "-l", IDIOMA_OCR],
                           input=pix.tobytes("png"), capture_output=True, timeout=300,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except FileNotFoundError:
        raise RuntimeError(f"No se encontró Tesseract en {TESSERACT_CMD}")
    if r.returncode != 0:
        raise RuntimeError("Tesseract falló: " + r.stderr.decode("utf-8", "ignore").strip()[:300])
    return r.stdout.decode("utf-8", "ignore")

def leer_pdf(file_bytes, etiqueta, paginas, al_avanzar=None):
    """Lee las páginas indicadas (1-based) del PDF.
    Devuelve (texto marcado con "[etiqueta, p. N]", resumen dict).
    al_avanzar() se llama tras cada página (para la barra de progreso)."""
    h = hash_archivo(file_bytes)
    partes = []
    nativas = ocr = 0
    with pymupdf.open(stream=file_bytes, filetype="pdf") as doc:
        for n in paginas:
            dato = _leer_cache(h, n)
            if dato is None:
                pagina = doc[n - 1]
                texto = pagina.get_text().strip()
                if len(texto) > MIN_CARACTERES:
                    dato = {"texto": texto, "metodo": "nativo"}
                else:
                    dato = {"texto": _ocr(pagina).strip(), "metodo": "ocr"}
                _guardar_cache(h, n, dato)
            if dato["metodo"] == "nativo":
                nativas += 1
            else:
                ocr += 1
            partes.append(f"[{etiqueta}, p. {n}]\n{dato['texto']}")
            if al_avanzar:
                al_avanzar()
    resumen = {"Documento": etiqueta, "Páginas leídas": len(paginas),
               "Texto nativo": nativas, "OCR": ocr}
    return "\n\n".join(partes), resumen

# -*- coding: utf-8 -*-
"""
==============================================================================
CLIENTE GEMINI (REST con urllib)
------------------------------------------------------------------------------
No usa google-generativeai: esa librería carga grpc y Smart App Control de
Windows la bloquea. Pide siempre JSON (responseMimeType), reintenta ante
429/500/503 con espera creciente y guarda cada respuesta en .cache/gemini/
(por hash de modelo + prompt + texto) para no repetir llamadas.
Si un modelo sigue saturado o agotó su cuota, prueba con los modelos de
respaldo que se le pasen.
==============================================================================
"""
import os, re, json, time, base64, hashlib, threading
import urllib.request, urllib.error, http.client

MODELO_GEMINI = "gemini-3.8-flash"   # el que recomienda Google: los 2.5 ya no se dan a cuentas nuevas
BASE = "https://generativelanguage.googleapis.com/v1beta"
INTENTOS = 5
ESPERAS = [10, 20, 40, 60]   # segundos entre reintentos (≈ 2 min por modelo)
MAX_ESPERA = 120             # si Google pide esperar más que esto, es cuota diaria: no se espera
TIMEOUT = 300                # segundos por llamada (los documentos son largos)
INTENTOS_LENTO = 2           # si el modelo no responde a tiempo 2 veces, se pasa al siguiente
INTENTOS_JSON = 2            # si devuelve JSON inválido 2 veces, se pasa al siguiente
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cache", "gemini")


class ErrorGemini(RuntimeError):
    pass


class _NoDisponible(ErrorGemini):
    """El modelo está saturado o sin cuota: conviene probar con otro."""
    def __init__(self, tipo, mensaje):
        super().__init__(mensaje); self.tipo = tipo


# Cuántas respuestas salieron de la caché y cuántas de Gemini (se usa desde varios hilos)
_estad = {"cache": 0, "api": 0}
_estad_lock = threading.Lock()


def _contar(tipo):
    with _estad_lock:
        _estad[tipo] += 1


def estadisticas():
    """{"cache": n, "api": n} acumulados desde que arrancó la app."""
    with _estad_lock:
        return dict(_estad)


def _clave(modelo, prompt, texto, imagenes=()):
    h = hashlib.sha256()
    for parte in (modelo, prompt, texto):
        h.update(parte.encode("utf-8")); h.update(b"\0")
    for img in imagenes:   # sin imágenes la clave es la misma de siempre
        h.update(hashlib.sha256(img).digest())
    return h.hexdigest()


def _leer_cache(clave):
    try:
        with open(os.path.join(CACHE_DIR, clave + ".json"), encoding="utf-8") as f:
            dato = json.load(f)
        return dato if isinstance(dato, (dict, list)) and dato else None   # ignora "null" guardados
    except (OSError, ValueError):
        return None


def _guardar_cache(clave, dato):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(os.path.join(CACHE_DIR, clave + ".json"), "w", encoding="utf-8") as f:
        json.dump(dato, f, ensure_ascii=False)


def _mayus(s):
    return s[:1].upper() + s[1:]


def _mensaje(detalle):
    """Extrae el "message" del JSON de error de Google (o el texto recortado)."""
    try:
        return json.loads(detalle)["error"]["message"]
    except (ValueError, KeyError, TypeError):
        return detalle[:300]


def _segundos_espera(detalle):
    """Segundos que Google pide esperar (retryDelay), o None."""
    m = re.search(r'"retryDelay":\s*"([\d.]+)s"', detalle)
    return float(m.group(1)) if m else None


class _JSONInvalido(ErrorGemini):
    """La respuesta llegó, pero cortada o con un JSON que no se pudo reparar."""
    def __init__(self, mensaje, raw):
        super().__init__(mensaje); self.raw = raw


def _reparar_json(raw):
    """Intenta leer el JSON corrigiendo los defectos típicos de Gemini: saltos de
    línea o tabulaciones dentro de los textos (strict=False), barras invertidas sin
    escapar (frecuentes al citar texto OCR), comas sobrantes y texto alrededor."""
    candidatos = [raw]
    ini, fin = raw.find("{"), raw.rfind("}")
    if 0 <= ini < fin:
        candidatos.append(raw[ini:fin + 1])
    for c in candidatos:
        for arreglo in (lambda s: s,
                        lambda s: re.sub(r",\s*([}\]])", r"\1", s),
                        lambda s: re.sub(r'\\(?!["\\/bfnrtu])', r"\\\\", re.sub(r",\s*([}\]])", r"\1", s))):
            try:
                return json.loads(arreglo(c), strict=False)
            except ValueError:
                pass
    raise ValueError("JSON no reparable")


def modelo_actual(respuesta):
    """'gemini-x ' si la respuesta trae la versión del modelo (para los mensajes)."""
    v = respuesta.get("modelVersion")
    return f"{v} " if v else "Gemini "


def _parsear(respuesta):
    """Extrae el JSON del texto de la respuesta de Gemini."""
    cands = respuesta.get("candidates") or []
    if not cands:
        motivo = respuesta.get("promptFeedback", {}).get("blockReason", "sin respuesta")
        raise ErrorGemini(f"Gemini no devolvió resultado ({motivo}).")
    cand = cands[0]
    partes = (cand.get("content") or {}).get("parts") or []
    raw = "".join(p.get("text", "") for p in partes if not p.get("thought")).strip()
    motivo = cand.get("finishReason", "")
    if motivo == "MAX_TOKENS":
        raise _JSONInvalido("devolvió una respuesta cortada por longitud (MAX_TOKENS)", raw)
    if not raw:   # p. ej. RECITATION o SAFETY: Google bloquea la respuesta y llega vacía
        if motivo in ("RECITATION", "SAFETY", "PROHIBITED_CONTENT", "BLOCKLIST", "SPII"):
            # repetir la misma pregunta al mismo modelo da el mismo bloqueo y gasta cuota
            raise _NoDisponible("bloqueado", f"{modelo_actual(respuesta)}bloqueó la respuesta ({motivo}).")
        raise _JSONInvalido(f"devolvió una respuesta vacía ({motivo or 'sin motivo'})", raw)
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    try:
        dato = _reparar_json(raw)
    except ValueError:
        raise _JSONInvalido("devolvió un JSON mal formado", raw)
    if not isinstance(dato, (dict, list)) or not dato:   # "null", un número o {} vacío
        raise _JSONInvalido("devolvió una respuesta sin datos", raw)
    return dato


def _guardar_error(modelo, raw):
    """Guarda la respuesta defectuosa para poder revisarla después."""
    carpeta = os.path.join(CACHE_DIR, "errores")
    os.makedirs(carpeta, exist_ok=True)
    nombre = time.strftime("%Y%m%d-%H%M%S") + f"_{time.time_ns() % 10**6:06d}_{modelo}.txt"
    with open(os.path.join(carpeta, nombre), "w", encoding="utf-8") as f:
        f.write(raw)


def listar_modelos(api_key):
    """Modelos Gemini de texto que esta API key puede usar (generateContent)."""
    req = urllib.request.Request(f"{BASE}/models?pageSize=1000",
                                 headers={"x-goog-api-key": api_key})
    with urllib.request.urlopen(req, timeout=30) as r:
        datos = json.loads(r.read().decode("utf-8"))
    excluir = ("embedding", "tts", "image", "audio", "live", "robotics", "computer-use", "aqa")
    nombres = []
    for m in datos.get("models", []):
        nombre = m.get("name", "").removeprefix("models/")
        if ("generateContent" in m.get("supportedGenerationMethods", [])
                and nombre.startswith("gemini") and not any(x in nombre for x in excluir)):
            nombres.append(nombre)
    return nombres


def modelos_de_respaldo(elegido, disponibles, maximo=3):
    """Otros modelos 'flash' para usar si el elegido falla: primero los estables
    y no-lite (más capaces), luego los lite (más cuota gratuita)."""
    candidatos = [m for m in disponibles if m != elegido and "flash" in m]
    def version(m):
        v = re.search(r"(\d+(?:\.\d+)?)", m)
        return float(v.group(1)) if v else 0.0
    def orden(m):
        # primero la versión "lite" del modelo elegido (otra cuota, mismo comportamiento),
        # luego los demás estables, del más nuevo al más antiguo
        return ("preview" in m or "exp" in m, version(m) != version(elegido), -version(m), "lite" in m, m)
    return sorted(candidatos, key=orden)[:maximo]


def _llamar_modelo(api_key, modelo, prompt, texto, al_esperar, imagenes=()):
    partes = [{"text": prompt + ("\n\n" + texto if texto else "")}]
    partes += [{"inline_data": {"mime_type": "image/png",
                                "data": base64.b64encode(img).decode("ascii")}} for img in imagenes]
    cuerpo = json.dumps({
        "contents": [{"parts": partes}],
        "generationConfig": {"temperature": 0, "responseMimeType": "application/json"},
    }).encode("utf-8")
    req = urllib.request.Request(f"{BASE}/models/{modelo}:generateContent", data=cuerpo,
                                 headers={"Content-Type": "application/json",
                                          "x-goog-api-key": api_key})
    lentas = [0]     # respuestas que superaron el TIMEOUT
    invalidas = [0]  # respuestas con JSON cortado o mal formado
    for i in range(INTENTOS):
        ultimo = i == INTENTOS - 1
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return _parsear(json.loads(r.read().decode("utf-8")))
        except _JSONInvalido as e:
            _guardar_error(modelo, e.raw)
            invalidas[0] += 1
            if invalidas[0] >= INTENTOS_JSON:
                raise _NoDisponible("json", f"{modelo} {e} {invalidas[0]} veces.")
            if al_esperar:
                al_esperar(f"{modelo} {e}. Reintentando…")
        except urllib.error.HTTPError as e:
            detalle = e.read().decode("utf-8", "ignore")
            if e.code == 429:
                pedida = _segundos_espera(detalle)
                if (pedida and pedida > MAX_ESPERA) or "PerDay" in detalle:
                    m = re.search(r"retry in ([\dhms.]+)", detalle)
                    cuando = f" Se renueva en {m.group(1).split('.')[0]}." if m else ""
                    raise _NoDisponible("cuota", f"se agotó la cuota diaria de {modelo}.{cuando}")
                espera = pedida + 1 if pedida else ESPERAS[min(i, len(ESPERAS) - 1)]
            elif e.code in (500, 503):
                espera = ESPERAS[min(i, len(ESPERAS) - 1)]
            elif e.code == 404:
                raise _NoDisponible("no existe", f"el modelo {modelo} no está disponible para tu API key "
                                                 f"(Google dice: {_mensaje(detalle)[:200]}).")
            elif e.code in (400, 403) and "API key" in detalle:
                raise ErrorGemini("La API key de Gemini no es válida. Revísala en la barra lateral.")
            else:
                raise ErrorGemini(f"La API de Gemini respondió {e.code}: {_mensaje(detalle)}")
            if ultimo:
                raise _NoDisponible("saturado", f"{modelo} sigue saturado (error {e.code}).")
            if al_esperar:
                al_esperar(f"{modelo} ocupado (error {e.code}). Reintento {i + 2} de "
                           f"{INTENTOS} en {espera:.0f} s…")
            time.sleep(espera)
        except (urllib.error.URLError, TimeoutError, ConnectionError, http.client.HTTPException) as e:
            # Respuesta que no llega a tiempo o conexión cortada a mitad: es temporal
            lento = isinstance(e, TimeoutError) or isinstance(getattr(e, "reason", None), TimeoutError)
            if lento:
                lentas[0] += 1
                if lentas[0] >= INTENTOS_LENTO:
                    raise _NoDisponible("lento", f"{modelo} no respondió en {TIMEOUT // 60} minutos.")
            elif ultimo:
                raise _NoDisponible("lento", f"se cortó la conexión con {modelo} ({getattr(e, 'reason', e)}).")
            if al_esperar:
                al_esperar(f"{modelo} " + ("tardó demasiado en responder" if lento else "cortó la conexión")
                           + f". Reintentando…")
            time.sleep(5)


def llamar_json(api_key, prompt, texto="", al_esperar=None, modelos=None, imagenes=()):
    """Envía prompt + texto (+ imágenes PNG) a Gemini y devuelve el JSON ya parseado.
    modelos: lista en orden de preferencia (el primero es el elegido; los demás,
    respaldo si el anterior está saturado o sin cuota).
    al_esperar(mensaje) informa reintentos y cambios de modelo."""
    modelos = modelos or [MODELO_GEMINI]
    imagenes = tuple(imagenes or ())
    for modelo in modelos:   # si algún modelo ya respondió esto antes, se reutiliza
        dato = _leer_cache(_clave(modelo, prompt, texto, imagenes))
        if dato is not None:
            _contar("cache")
            return dato

    fallas = []
    for k, modelo in enumerate(modelos):
        try:
            dato = _llamar_modelo(api_key, modelo, prompt, texto, al_esperar, imagenes)
            _guardar_cache(_clave(modelo, prompt, texto, imagenes), dato)
            _contar("api")
            return dato
        except _NoDisponible as e:
            fallas.append(e)
            if k + 1 < len(modelos) and al_esperar:
                al_esperar(f"{_mayus(str(e))} Probando con {modelos[k + 1]}…")

    detalle = " ".join(_mayus(str(e)) for e in fallas)
    if all(e.tipo == "bloqueado" for e in fallas):
        raise ErrorGemini(f"Google bloqueó la respuesta ({detalle}) Suele pasar cuando se pide "
                          f"copiar textos largos de un documento. Lo ya procesado queda guardado.")
    if all(e.tipo == "json" for e in fallas):
        raise ErrorGemini(f"Gemini devolvió respuestas cortadas o mal formadas ({detalle}) "
                          f"Vuelve a presionar el botón o elige otro modelo en la barra lateral. "
                          f"Lo que ya se procesó queda guardado y no se repite.")
    if all(e.tipo in ("saturado", "lento", "json") for e in fallas):
        raise ErrorGemini(f"Gemini está saturado o lento en este momento ({detalle}) Es un problema "
                          f"temporal de Google: espera unos minutos y vuelve a presionar el botón. "
                          f"Lo que ya se procesó queda guardado y no se repite.")
    raise ErrorGemini(f"No hay modelos de Gemini disponibles ahora. {detalle} Elige otro modelo "
                      f"en la barra lateral, espera a que se renueve la cuota o activa la "
                      f"facturación en Google AI Studio. Lo ya procesado queda guardado.")

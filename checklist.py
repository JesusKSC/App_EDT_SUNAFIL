# -*- coding: utf-8 -*-
"""
==============================================================================
CHECKLIST DEL TDR  (análisis en 3 pasos + reglas fijas)
------------------------------------------------------------------------------
Paso 1a  Inventario de puntos del Capítulo III + factores del Capítulo IV
Paso 1b  Detalle de cada punto (Bases + Pliego), en lotes de 20
Paso 2   Evaluación contra la Propuesta y el Contrato, en lotes de 20
Paso 3   Síntesis: ficha resumen, inconsistencias y recomendaciones
Después de la IA se aplican las reglas fijas R1–R3 y se calculan en Python
todos los conteos y puntajes.
==============================================================================
"""
import re, json, time, queue, unicodedata
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from gemini_client import llamar_json, estadisticas
import tablas_imagen

LOTE = 30
PARALELO = 3   # lotes del paso 2 (y lecturas de imágenes) que se piden a Gemini a la vez
VEREDICTOS = ("SI", "NO", "PARCIAL", "NO EVIDENCIADO", "NO APLICA")
VERIFICAR_PDF = "Verificar en el PDF"
CAMPOS_FICHA = ["Entidad", "Objeto", "Procedimiento", "Contratista", "Monto contractual",
                "Plazo de ejecución", "Plazo de activación", "Garantías", "Penalidades"]

ROL = """Actúa como analista experto en contrataciones públicas del Estado peruano y
Project Manager certificado PMP. Nunca inventes datos: si algo no está en los
documentos, déjalo vacío. Cita siempre la fuente con el formato [Bases, p. N],
[Pliego, consulta N], [Propuesta, p. N] o [Contrato, cláusula N]. En el texto,
las marcas [Documento, p. N] indican de qué página viene cada fragmento.
Responde exclusivamente con JSON válido, con la estructura indicada."""

PROMPT_1A = ROL + """

PASO 1a — INVENTARIO DE PUNTOS DEL TDR
Con las Bases y el Pliego, lista TODOS los numerales y literales del Capítulo III
(Requerimiento / Términos de Referencia) con su título, en el orden en que aparecen
(ej. 3.1, 3.3.c, 7.d, 8.2.b). No te saltes ninguno y no resumas: debe salir una
lista larga (normalmente 80 o más puntos). Si un literal tiene partes que se
evalúan por separado (por ejemplo Gama A, Gama B y Gama C), crea una fila por
parte, con el mismo numeral y un título que las distinga. Ignora lo administrativo
de la convocatoria (calendario del procedimiento, registro de participantes,
forma de presentar ofertas, etc.).
Extrae también los factores de evaluación del Capítulo IV con su puntaje máximo
y, si las Bases lo indican, el puntaje técnico mínimo.

JSON:
{"puntos": [{"numeral": "3.3.c", "titulo": "Plazo de prestación del servicio"}],
 "factores": [{"factor": "Factor C", "titulo": "Plazo de activación", "puntaje_max": 10,
               "criterio": "cómo se asigna el puntaje", "fuente": "[Bases, p. 60]"}],
 "puntaje_minimo": 70,
 "fuente_puntaje_minimo": "[Bases, p. 58]"}
Si las Bases no indican el puntaje mínimo, usa null."""

PROMPT_1B = ROL + """

PASO 1b — DETALLE DE CADA PUNTO
Para cada punto de la lista (usa las Bases y el Pliego):
- clase: solo "Producto" o "Proyecto". "Producto" = el bien o servicio que recibe la
  Entidad (líneas, planes, equipos, cobertura, plataformas, funcionalidades).
  "Proyecto" = todo lo que es gestión del contrato: plazos, entregables (planes,
  informes, actas), personal clave, capacitación, penalidades, garantías,
  conformidad y pagos
- no_aplica: true si las Bases dicen expresamente que el punto "No aplica" (por
  ejemplo adelanto directo, fórmulas de reajuste, sistema de entrega, prestaciones
  accesorias); en otro caso false
- alcance_bases: qué exige el TDR, en una o dos frases (A)
- fuente_bases: [Bases, p. N]
- consulta_pliego: la consulta u observación del Pliego que lo modifica o precisa,
  con su cita [Pliego, consulta N] (B); si no hay, "Sin consultas"
- alcance_vigente: el alcance que rige después de la consulta (A+B)
- valor_exigido: el valor medible exigido con unidad y condición (ej. "mínimo 30 GB",
  "hasta 30 d.c.", "1145 líneas"); si no hay cifra, el requisito verificable en pocas palabras
Devuelve exactamente un elemento por cada "n" recibido.

JSON:
{"detalle": [{"n": 1, "clase": "", "no_aplica": false, "alcance_bases": "", "fuente_bases": "",
              "consulta_pliego": "", "alcance_vigente": "", "valor_exigido": ""}]}

PUNTOS:
"""

PROMPT_2 = ROL + """

PASO 2 — EVALUACIÓN DE LA OFERTA
Compara cada punto (lo exigido según Bases + Pliego) contra la Propuesta técnica y,
si se cargó, el Contrato. Por cada punto:
- valor_ofertado: el dato concreto con unidad (ej. "14 d.c.", "30 GB") tomado de la
  Propuesta o, si la Propuesta no lo trae, del Contrato; vacío si ninguno lo indica
- fuente_propuesta: la cita del dato: [Propuesta, p. N] y/o [Contrato, cláusula N]
- fuente_veredicto: "Propuesta", "Contrato" o "Ambos", según qué documento acredita el
  punto; vacío si no está acreditado
- alcance_final: el alcance que rige al final (A+B+C: Bases + Pliego + lo ofertado)
- veredicto: "SI", "NO", "PARCIAL" o "NO EVIDENCIADO"
- motivo: compara exigido vs. ofertado con las cifras (ej. "Ofrece 14 d.c.; el TDR
  exige hasta 30 d.c.")
- reflejo_contrato: cómo queda en el Contrato, con [Contrato, cláusula N];
  "No se cargó el contrato" si no hay contrato; "No figura en el contrato" si no aparece
- mejora: si ofrece algo por encima del mínimo exigido, descríbelo con la cifra; si no, vacío
- mejora_valor: cuánto supera el mínimo (ej. "+16 d.c.", "+10 GB"); vacío si no hay mejora
- riesgo: riesgo para la ejecución del contrato; vacío si no hay
- puntaje_obtenido: SOLO si el punto trae "puntaje_max" (factor de evaluación): el puntaje
  que corresponde según la evidencia y el criterio de las Bases; en otro caso null
- experiencia_total_anios: SOLO en factores de experiencia: la experiencia TOTAL acreditada
  en años, como número (ej. 5 años y 9 meses = 5.75), sin restar el mínimo exigido; si no
  aplica, null
Reglas:
- Una declaración genérica del tipo "cumplimos con todo lo solicitado" o una
  declaración jurada de cumplimiento NO es evidencia: el veredicto es "NO EVIDENCIADO".
- CONTRATO COMO EVIDENCIA: si la Propuesta no menciona el punto pero el Contrato lo
  formaliza con un dato concreto (con su cláusula), el veredicto es "SI",
  fuente_veredicto "Contrato" y el motivo empieza con "Acreditado en el Contrato:".
- "NO" solo si el valor ofertado existe y no alcanza lo exigido. "PARCIAL" si cubre
  solo una parte.
- En los requisitos, "más de X", "mínimo X" o "a partir de X" son pisos (ofrecer más
  es mejor); solo "hasta X", "máximo X" o "no mayor a X" son topes. El "puntaje
  máximo" de un factor no es un tope del requisito.
- Si el texto OCR de una tabla es ilegible (cifras que no suman o que contradicen el
  Contrato), no saques conclusiones de esas cifras: escribe en riesgo
  "Verificar en el PDF: " y qué cifra revisar.
Devuelve exactamente un elemento por cada "n" recibido.

JSON:
{"evaluacion": [{"n": 1, "valor_ofertado": "", "fuente_propuesta": "", "fuente_veredicto": "",
                 "alcance_final": "", "veredicto": "", "motivo": "", "reflejo_contrato": "",
                 "mejora": "", "mejora_valor": "", "riesgo": "", "puntaje_obtenido": null,
                 "experiencia_total_anios": null}]}

PUNTOS:
"""

PROMPT_3 = ROL + """

PASO 3 — SÍNTESIS
Con los resultados de la evaluación (abajo) y los documentos, arma:
1) ficha: Entidad, Objeto, Procedimiento, Contratista, Monto contractual, Plazo de
   ejecución, Plazo de activación, Garantías, Penalidades. Si hay Contrato, toma los
   datos de ahí; si no, de las Bases y la Propuesta. Cada campo con su fuente; si no
   aparece, valor vacío.
2) inconsistencias: inconsistencias y riesgos priorizados (Alta / Media / Baja), sobre
   todo Bases vs. Contrato (plazos, cantidades de líneas o equipos, GB, modelos de
   equipos), además de los puntos en NO o PARCIAL. Si la inconsistencia puede venir
   de un OCR ilegible (cifras de una tabla que no suman o que contradicen el
   Contrato), la prioridad es "Verificar en el PDF", no "Alta".
3) recomendaciones para administrar el contrato: el ítem a monitorear, el entregable
   que lo acredita y el detalle.
4) mejoras: las mejoras ofrecidas (campo "mejora" de los resultados), consolidadas:
   si la misma mejora aparece en varios numerales (por ejemplo en un punto del TDR y
   en un factor de evaluación), va UNA sola vez, con todos esos numerales separados
   por comas. Solo mejoras con cifra.

JSON:
{"ficha": [{"campo": "Entidad", "valor": "", "fuente": ""}],
 "mejoras": [{"mejora": "", "numerales": "3.4.VII.i, Factor I", "exigido": "", "ofrecido": "",
              "valor_sobre_minimo": "+20 GB", "reflejo_contrato": "", "evidencia": "[Propuesta, p. N]"}],
 "inconsistencias": [{"prioridad": "Alta", "descripcion": "",
                      "documentos": "[Bases, p. 26] vs. [Contrato, cláusula 5]"}],
 "recomendaciones": [{"item": "", "entregable": "", "detalle": ""}]}

RESULTADOS DE LA EVALUACIÓN:
"""


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _docs(textos, nombres):
    return "\n\n".join(f"===== DOCUMENTO: {k} =====\n{textos[k]}"
                       for k in nombres if textos.get(k))


def _como_dict(res, clave):
    """Gemini a veces devuelve directamente la lista en vez de {clave: lista}."""
    return {clave: res} if isinstance(res, list) else res


def _txt(v):
    return "" if v is None else str(v).strip()


def _sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def _veredicto(v):
    v = _sin_tildes(_txt(v)).upper()
    return v if v in VEREDICTOS else "NO EVIDENCIADO"


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    m = re.search(r"\d+(?:[.,]\d+)?", _txt(v))
    return float(m.group().replace(",", ".")) if m else None


def _anotar(fila, nota):
    fila["motivo"] = (fila["motivo"] + " " if fila["motivo"] else "") + nota


# ---------------------------------------------------------------------------
# REGLAS FIJAS (R1–R3)
# ---------------------------------------------------------------------------
_GENERICO = re.compile(
    r"cumpl|conforme|de acuerdo (a|con)|segun (lo |el |la |las |los )?(solicitad|requerid|exigid|bases|tdr|terminos)"
    r"|declaracion jurada|se compromete|\bacepta(mos)?\b|no (se )?(especifica|indica|encontr|evidencia|menciona|precisa|consigna)"
    r"|sin dato|no aplica|^n/?a$|^-+$|^—$", re.I)


def es_concreto(valor):
    """¿El valor ofertado es un dato concreto (no vacío ni declaración genérica)?"""
    v = _sin_tildes(_txt(valor)).lower()
    if not v:
        return False
    if re.search(r"\d", v):
        return True
    return not _GENERICO.search(v)


# unidad normalizada <- patrón
_UNIDADES = [
    ("dias_habiles", r"d[ií]as?\s+h[aá]biles|d\.?\s?h\.?"),
    ("dias", r"d[ií]as?(\s+calendarios?)?|d\.?\s?c\.?"),
    ("horas", r"horas?|hrs?\.?|h\b"),
    ("meses", r"meses|mes\b"),
    ("anios", r"a[nñ]os?"),
    ("gb", r"gb\b|gigas?(bytes?)?"),
    ("mb", r"mb\b|megas?(bytes?)?"),
    ("tb", r"tb\b"),
    ("mp", r"mpx?\b|megap[ií]xel(es)?"),
    ("mah", r"mah\b"),
    ("pulgadas", r"pulgadas?|\""),
    ("lineas", r"l[ií]neas?"),
    ("equipos", r"equipos?"),
    ("porcentaje", r"%"),
]
_NUM = r"(\d{1,3}(?:[.,]\d{3})+|\d+(?:[.,]\d+)?)"


def _a_float(s):
    if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})+", s):      # 1,145 / 1.145 -> miles
        return float(re.sub(r"[.,]", "", s))
    return float(s.replace(",", "."))


def _cifras(texto):
    """{unidad: [valores]} encontrados en el texto."""
    t = _sin_tildes(_txt(texto)).lower()
    res = {}
    for unidad, patron in _UNIDADES:
        for m in re.finditer(_NUM + r"\s*(?:" + _sin_tildes(patron) + r")", t):
            res.setdefault(unidad, []).append(_a_float(m.group(1)))
        t = re.sub(_NUM + r"\s*(?:" + _sin_tildes(patron) + r")", " ", t)  # no contar dos veces
    return res


# Pisos: ofrecer más es mejor. Topes: ofrecer más incumple. ("desde" no se usa: suele
# referirse a fechas, p. ej. "desde la firma del contrato".)
_PISO = re.compile(r"\bmas de\b|minim|a partir de|o superior|\bo mas\b|al menos|no menor|igual o mayor", re.I)
_TOPE = re.compile(r"\bhasta\b|maxim|no mayor a", re.I)
_PUNTAJE_MAX = re.compile(r"puntajes?\s+maxim[oa]s?", re.I)


def regla_r2(exigido, ofertado):
    """Devuelve la nota si lo ofertado incumple claramente lo exigido; si hay duda, None."""
    ex_txt = _PUNTAJE_MAX.sub(" ", _sin_tildes(_txt(exigido)).lower())   # no es un tope
    es_piso, es_tope = bool(_PISO.search(ex_txt)), bool(_TOPE.search(ex_txt))
    if es_piso == es_tope:          # ninguna condición o ambas: duda
        return None
    ex, of = _cifras(ex_txt), _cifras(ofertado)
    comunes = [u for u in ex if u in of]
    if len(comunes) != 1:
        return None
    u = comunes[0]
    if len(set(ex[u])) != 1 or len(set(of[u])) != 1:   # varias cifras: duda
        return None
    e, o = ex[u][0], of[u][0]
    if es_piso and o < e:
        return f"[Regla R2: ofrece {o:g} y el mínimo exigido es {e:g} → NO]"
    if es_tope and o > e:
        return f"[Regla R2: ofrece {o:g} y el máximo permitido es {e:g} → NO]"
    return None


# Tramos de un factor por años: "Más de 4 años: 30 pts; Más de 3 a 4 años: 20 pts"
_TRAMO = re.compile(r"mas de\s*(\d+(?:[.,]\d+)?)(?:\s*(?:a|hasta)\s*(\d+(?:[.,]\d+)?))?\s*anos?"
                    r"\s*[:=\-–(]?\s*(\d+(?:[.,]\d+)?)\s*(?:pts|puntos|ptos)", re.I)


def tramos_de(texto):
    """[(desde, hasta o None, puntaje)] leídos del criterio del factor."""
    t = _sin_tildes(_txt(texto)).lower()
    return [(float(a.replace(",", ".")), float(b.replace(",", ".")) if b else None,
             float(p.replace(",", "."))) for a, b, p in _TRAMO.findall(t)]


def anios_de(texto):
    """'5 años, 9 meses y 24 días' -> 5.81 (None si no hay años ni meses)."""
    t = _sin_tildes(_txt(texto)).lower()
    a = re.search(r"(\d+(?:[.,]\d+)?)\s*anos?", t)
    m = re.search(r"(\d+)\s*mes", t)
    d = re.search(r"(\d+)\s*dias?", t)
    if not a and not m:
        return None
    total = float(a.group(1).replace(",", ".")) if a else 0.0
    return total + (int(m.group(1)) / 12 if m else 0) + (int(d.group(1)) / 365 if d else 0)


def regla_tramos(f):
    """Factores por tramos de años (p. ej. Factor B.1): se compara la EXPERIENCIA TOTAL
    acreditada con cada tramo, sin restar el mínimo exigido. Lo decide Python."""
    tramos = tramos_de(f.get("criterio", "") + " " + f.get("valor_exigido", ""))
    if not tramos:
        return
    x = f.get("experiencia_total")
    if x is None:
        x = anios_de(f["valor_ofertado"])
    if x is None:
        return
    puntaje, elegido = 0.0, None
    for desde, hasta, pts in tramos:
        if x > desde and (hasta is None or x <= hasta) and pts >= puntaje:
            puntaje, elegido = pts, (desde, hasta)
    f["puntaje_obt"] = puntaje
    f["veredicto"] = ("SI" if puntaje >= f["puntaje_max"] else "PARCIAL" if puntaje > 0 else "NO")
    if not es_concreto(f["valor_ofertado"]):
        f["valor_ofertado"] = f"{x:.2f} años"
    if elegido:
        desde, hasta = elegido
        tramo = f"más de {desde:g}" + (f" a {hasta:g}" if hasta is not None else "") + " años"
        _anotar(f, f"[Regla de tramos: experiencia total {x:.2f} años → tramo «{tramo}» → {puntaje:g} pts]")
    else:
        _anotar(f, f"[Regla de tramos: experiencia total {x:.2f} años no alcanza ningún tramo → 0 pts]")


def aplicar_reglas(filas):
    for f in filas:
        if f["veredicto"] == "NO APLICA":
            continue
        if f["es_factor"]:
            regla_tramos(f)
        # R1: SI sin valor ofertado concreto -> NO EVIDENCIADO
        if f["veredicto"] == "SI" and not es_concreto(f["valor_ofertado"]):
            f["veredicto"] = "NO EVIDENCIADO"
            f["fuente_veredicto"] = ""
            _anotar(f, "[Regla R1: SI sin valor ofertado concreto → NO EVIDENCIADO]")
        # R2: comparación de cifras de la misma unidad (no se aplica a factores de evaluación)
        if not f["es_factor"] and f["veredicto"] != "NO" and es_concreto(f["valor_ofertado"]):
            nota = regla_r2(f["valor_exigido"], f["valor_ofertado"])
            if nota:
                f["veredicto"] = "NO"
                _anotar(f, nota)
        # R3: mejora sin cifra que la sustente -> se descarta
        if f["mejora"] and not re.search(r"\d", f["mejora"] + " " + f["mejora_valor"]):
            _anotar(f, f"[Regla R3: mejora sin cifra descartada: «{f['mejora']}»]")
            f["mejora"], f["mejora_valor"] = "", ""
        # Puntaje del factor: dentro de [0, máx.] y 0 si no hay evidencia
        if f["es_factor"]:
            p = max(0.0, min(f["puntaje_obt"] or 0.0, f["puntaje_max"]))
            if f["veredicto"] == "NO EVIDENCIADO" and p > 0:
                _anotar(f, "[Puntaje en 0: sin evidencia en la propuesta ni en el contrato]")
                p = 0.0
            f["puntaje_obt"] = p
    return filas


def calcular_conteos(filas, factores, mejoras, puntaje_minimo):
    """filas = hoja B (sin factores). El puntaje sale de los factores (hoja B2)."""
    total = len(filas)
    c = {v: sum(1 for f in filas if f["veredicto"] == v) for v in VEREDICTOS}
    aplicables = [f for f in filas if f["veredicto"] != "NO APLICA"]
    con_dato = sum(1 for f in aplicables if es_concreto(f["valor_ofertado"]))
    return {
        "total": total, "si": c["SI"], "no": c["NO"], "parcial": c["PARCIAL"],
        "no_evidenciado": c["NO EVIDENCIADO"], "no_aplica": c["NO APLICA"],
        "aplicables": len(aplicables),
        "cumplimiento": (c["SI"] / len(aplicables)) if aplicables else 0.0,
        "cobertura": (con_dato / len(aplicables)) if aplicables else 0.0,
        "mejoras": len(mejoras),
        "puntaje_obt": sum(f["puntaje_obt"] for f in factores),
        "puntaje_max": sum(f["puntaje_max"] for f in factores),
        "puntaje_minimo": puntaje_minimo,
        "items_no": [f"{f['numeral']} {f['titulo']}" for f in filas if f["veredicto"] == "NO"],
    }


# ---------------------------------------------------------------------------
# AJUSTES DE PYTHON SOBRE LO QUE DEVUELVE LA IA
# ---------------------------------------------------------------------------
def quitar_duplicados(puntos):
    """Las Bases resumen las condiciones en 3.3.* y las repiten en el TDR (3.4.*).
    Si existe la numeración del TDR, se quitan las filas 3.3.*."""
    if not any(re.match(r"3\.4(\.|$)", p["numeral"]) for p in puntos):
        return puntos, []
    quitados = [p["numeral"] for p in puntos if re.match(r"3\.3(\.|$)", p["numeral"])]
    return [p for p in puntos if not re.match(r"3\.3(\.|$)", p["numeral"])], quitados


def _bool(v):
    return v is True or _sin_tildes(_txt(v)).lower() in ("true", "si", "1")


def _clase(v):
    return "Producto" if _sin_tildes(_txt(v)).lower() == "producto" else "Proyecto"


def _fuente_veredicto(f, valor_ia):
    """Propuesta / Contrato / Ambos (vacío si el punto no quedó acreditado)."""
    if f["veredicto"] in ("NO EVIDENCIADO", "NO APLICA"):
        return ""
    v = _sin_tildes(_txt(valor_ia)).lower()
    if not v:   # la IA no lo dijo: se deduce de las citas
        v = _sin_tildes(f["fuente_propuesta"] + " " + f["motivo"][:30]).lower()
    prop, cont = "propuesta" in v, "contrato" in v
    if "ambos" in v or (prop and cont):
        return "Ambos"
    return "Contrato" if cont else "Propuesta" if prop else ""


def _marcar_no_aplica(f):
    f.update(veredicto="NO APLICA", fuente_veredicto="", valor_ofertado="", fuente_propuesta="",
             alcance_final=f["alcance_vigente"], reflejo_contrato="", mejora="", mejora_valor="",
             riesgo="", puntaje_obt=None,
             motivo="Las Bases indican que no aplica" +
                    (f" {f['fuente_bases']}." if f["fuente_bases"] else "."))


def _normalizar_mejoras(lista, numerales_validos):
    """Mejoras consolidadas por la IA en el paso 3: una fila por mejora, con todos
    los numerales que impacta. Se descartan las que no tienen cifra (R3)."""
    mejoras, vistas = [], set()
    for x in lista or []:
        if not isinstance(x, dict):
            continue
        texto = _txt(x.get("mejora"))
        valor = _txt(x.get("valor_sobre_minimo"))
        if not texto or not re.search(r"\d", texto + " " + valor):
            continue
        clave = re.sub(r"\W+", " ", _sin_tildes(texto).lower()).strip()
        if clave in vistas:
            continue
        vistas.add(clave)
        nums = [n.strip() for n in re.split(r"[,;]", _txt(x.get("numerales"))) if n.strip()]
        nums = [n for n in nums if n in numerales_validos] or nums
        mejoras.append({"mejora": texto, "numerales": ", ".join(dict.fromkeys(nums)),
                        "exigido": _txt(x.get("exigido")), "ofrecido": _txt(x.get("ofrecido")),
                        "valor_sobre_minimo": valor,
                        "reflejo_contrato": _txt(x.get("reflejo_contrato")),
                        "evidencia": _txt(x.get("evidencia"))})
    return mejoras


def _mejoras_por_texto(filas):
    """Respaldo si la IA no consolidó: junta las mejoras con el mismo texto."""
    grupos = {}
    for f in filas:
        if not f.get("mejora"):
            continue
        clave = re.sub(r"\W+", " ", _sin_tildes(f["mejora"]).lower()).strip()
        g = grupos.setdefault(clave, {"mejora": f["mejora"], "numerales": [], "exigido": [],
                                      "ofrecido": [], "valor_sobre_minimo": [],
                                      "reflejo_contrato": [], "evidencia": []})
        g["numerales"].append(f["numeral"])
        for k, campo in (("exigido", "valor_exigido"), ("ofrecido", "valor_ofertado"),
                         ("valor_sobre_minimo", "mejora_valor"),
                         ("reflejo_contrato", "reflejo_contrato"), ("evidencia", "fuente_propuesta")):
            if f[campo] and f[campo] not in g[k]:
                g[k].append(f[campo])
    return [{k: (v if k == "mejora" else (", " if k == "numerales" else "; ").join(v))
             for k, v in g.items()} for g in grupos.values()]


# ---------------------------------------------------------------------------
# CONTROL DE COHERENCIA (después de la IA)
# ---------------------------------------------------------------------------
_MEJORA_MAS = re.compile(r"\+\s*" + _NUM + r"\s*(?:gb|gigas?|mb|minutos?|lineas?|equipos?)\b", re.I)


def control_coherencia(filas):
    """Si una mejora declara "+X <unidad> sobre el mínimo", el exigido debe ser
    (valor del Contrato − X). Si no cuadra, no se inventa la diferencia: la fila
    queda con riesgo "Verificar en el PDF" y la mejora no se considera.
    Devuelve los numerales marcados."""
    marcados = set()
    for f in filas:
        if not f.get("mejora"):
            continue
        texto = _sin_tildes(f["mejora"] + " " + f["mejora_valor"]).lower()
        m = _MEJORA_MAS.search(texto)
        if not m:
            continue
        mas = _cifras(m.group().lstrip("+ "))
        if len(mas) != 1:
            continue
        unidad, (x, *_) = next(iter(mas.items()))
        ex = _cifras(f["valor_exigido"]).get(unidad, [])
        co = _cifras(f["reflejo_contrato"]).get(unidad, [])
        if len(set(ex)) != 1 or len(set(co)) != 1:   # sin una cifra clara no se puede verificar
            continue
        e, c = ex[0], co[0]
        if abs(e + x - c) < 0.01:
            continue
        f["riesgo"] = (f"{VERIFICAR_PDF}: exigido {e:g} + mejora {x:g} = {e + x:g}, pero el "
                       f"Contrato dice {c:g} ({unidad.upper()}). Revisar las cifras en el PDF."
                       + (f" {f['riesgo']}" if f["riesgo"] else ""))
        _anotar(f, f"[Coherencia: la mejora «{f['mejora']}» no cuadra con el Contrato; no se considera]")
        f["mejora"], f["mejora_valor"] = "", ""
        if f["veredicto"] == "NO":
            f["veredicto"] = "NO EVIDENCIADO"
            _anotar(f, "[Coherencia: no se afirma un incumplimiento con cifras que no cuadran]")
        marcados.add(f["numeral"])
    return marcados


def _menciona(texto, numerales):
    return any(re.search(r"(?<![\w.])" + re.escape(n) + r"(?![\w])", texto) for n in numerales)


# ---------------------------------------------------------------------------
# ANÁLISIS EN 3 PASOS
# ---------------------------------------------------------------------------
def _lotes(filas):
    return [filas[i:i + LOTE] for i in range(0, len(filas), LOTE)]


CAMPOS_EVAL = ("valor_ofertado", "fuente_propuesta", "alcance_final", "motivo",
               "reflejo_contrato", "mejora", "mejora_valor", "riesgo")


def _en_paralelo(tareas, etiqueta, avisar):
    """Corre tareas fn(al_esperar) con hasta PARALELO a la vez y devuelve sus
    resultados en orden. Solo este hilo actualiza la pantalla (Streamlit no admite
    llamadas desde otros hilos): los avisos de reintento llegan por una cola."""
    if not tareas:
        return []
    mensajes = queue.Queue()
    resultados = [None] * len(tareas)
    ex = ThreadPoolExecutor(max_workers=PARALELO)
    futuros = {ex.submit(fn, mensajes.put): i for i, fn in enumerate(tareas)}
    pendientes, listos_n, ultimo = set(futuros), 0, ""
    avisar(f"{etiqueta}: 0 de {len(tareas)} listos…", 0)
    try:
        while pendientes:
            listos, pendientes = wait(pendientes, timeout=1, return_when=FIRST_COMPLETED)
            for fut in listos:
                resultados[futuros[fut]] = fut.result()   # si falló, se propaga aquí
                listos_n += 1
            while not mensajes.empty():
                ultimo = mensajes.get_nowait()
            avisar(f"{etiqueta}: {listos_n} de {len(tareas)} listos"
                   + (f" · {ultimo}" if ultimo and pendientes else "…"), listos_n)
    except BaseException:
        ex.shutdown(wait=False, cancel_futures=True)   # lo ya terminado quedó en caché
        raise
    ex.shutdown(wait=True)
    return resultados


def generar_checklist(api_key, textos, progreso=lambda texto, frac: None, modelos=None, pdfs=None):
    """Corre los pasos. progreso(texto, fraccion) informa el avance.
    pdfs: {documento: bytes del PDF} para leer las tablas desde la imagen (opcional).
    Devuelve dict con filas (hoja B), factores (hoja B2), mejoras, ficha,
    inconsistencias, recomendaciones, conteos y tiempos por paso."""
    estado = {"hechos": 0, "total": 3, "frac": 0.0}
    tiempos = []

    def avisar(texto, extra=0):
        # el total de pasos se conoce recién después del inventario: la barra nunca retrocede
        frac = min((estado["hechos"] + extra) / max(estado["total"], 1), 1.0)
        estado["frac"] = max(estado["frac"], frac)
        progreso(texto, estado["frac"])

    def medir(nombre, inicio, est0):
        est1 = estadisticas()
        tiempos.append({"Paso": nombre, "Tiempo (s)": round(time.time() - inicio, 1),
                        "Llamadas a Gemini": est1["api"] - est0["api"],
                        "Desde la caché": est1["cache"] - est0["cache"]})

    # Paso 0 — tablas leídas desde la imagen de la página
    t0, e0 = time.time(), estadisticas()
    tareas_img = tablas_imagen.tareas_de_lectura(api_key, textos, pdfs or {}, modelos)
    estado["total"] = len(tareas_img) + 3
    textos = dict(textos)
    paginas_imagen, avisos = {}, []
    for (doc, grupo, _), (nuevos, error) in zip(tareas_img, _en_paralelo(
            [fn for _, _, fn in tareas_img], "Paso 0: tablas leídas desde la imagen", avisar)):
        if error or not nuevos:
            avisos.append(f"{doc} p. {', '.join(map(str, grupo))}: no se pudo leer desde la imagen; "
                          f"se usó el texto OCR. ({error or 'Gemini no devolvió esas páginas'})")
            continue
        textos[doc] = tablas_imagen.reemplazar_paginas(textos[doc], doc, nuevos)
        paginas_imagen.setdefault(doc, []).extend(sorted(nuevos))
    estado["hechos"] += len(tareas_img)
    if tareas_img:
        medir("0. Tablas desde la imagen", t0, e0)

    texto_bp = _docs(textos, ["Bases", "Pliego"])
    texto_pc = _docs(textos, ["Propuesta", "Contrato"])
    hay_contrato = bool(textos.get("Contrato"))

    # Paso 1a — inventario
    t0, e0 = time.time(), estadisticas()
    avisar("Paso 1a: inventario de puntos del TDR…")
    inv = _como_dict(llamar_json(api_key, PROMPT_1A, texto_bp, avisar, modelos), "puntos")
    puntos, duplicados = quitar_duplicados(
        [{"numeral": _txt(p.get("numeral")), "titulo": _txt(p.get("titulo"))}
         for p in inv.get("puntos") or [] if isinstance(p, dict)])
    filas = [dict(p, es_factor=False, puntaje_max=None, criterio="") for p in puntos]
    factores = [{"numeral": _txt(fz.get("factor")), "titulo": _txt(fz.get("titulo")),
                 "es_factor": True, "puntaje_max": _num(fz.get("puntaje_max")) or 0.0,
                 "criterio": _txt(fz.get("criterio"))}
                for fz in inv.get("factores") or [] if isinstance(fz, dict)]
    if not filas:
        raise ValueError("Gemini no encontró puntos del TDR. Revisa que el rango de páginas "
                         "de las Bases incluya el Capítulo III.")
    todos = filas + factores
    for i, f in enumerate(todos, 1):
        f["n"] = i
    puntaje_minimo = _num(inv.get("puntaje_minimo"))
    lotes_1b = _lotes(todos)
    estado["hechos"] += 1
    estado["total"] = estado["hechos"] + 2 * len(lotes_1b) + 1   # estimado
    medir("1a. Inventario", t0, e0)

    # Paso 1b — detalle por lotes
    t0, e0 = time.time(), estadisticas()
    for k, lote in enumerate(lotes_1b, 1):
        avisar(f"Paso 1b: detalle de los puntos (lote {k} de {len(lotes_1b)})…")
        entrada = [{"n": f["n"], "numeral": f["numeral"], "titulo": f["titulo"]}
                   | ({"factor_de_evaluacion": True, "criterio": f["criterio"]}
                      if f["es_factor"] else {})
                   for f in lote]
        res = _como_dict(llamar_json(api_key, PROMPT_1B + json.dumps(entrada, ensure_ascii=False),
                                     texto_bp, avisar, modelos), "detalle")
        por_n = {d.get("n"): d for d in res.get("detalle") or [] if isinstance(d, dict)}
        for f in lote:
            d = por_n.get(f["n"], {})
            for campo in ("alcance_bases", "fuente_bases", "consulta_pliego",
                          "alcance_vigente", "valor_exigido"):
                f[campo] = _txt(d.get(campo))
            f["clase"] = "Factor de evaluación" if f["es_factor"] else _clase(d.get("clase"))
            f["no_aplica"] = (not f["es_factor"]) and _bool(d.get("no_aplica"))
            if not f["consulta_pliego"]:
                f["consulta_pliego"] = "Sin consultas"
        estado["hechos"] += 1
    medir("1b. Detalle de los puntos", t0, e0)

    # Paso 2 — evaluación por lotes, en paralelo (los puntos que no aplican no se evalúan)
    t0, e0 = time.time(), estadisticas()
    for f in todos:
        if f["no_aplica"]:
            _marcar_no_aplica(f)
    lotes_2 = _lotes([f for f in todos if not f["no_aplica"]])
    estado["total"] = estado["hechos"] + len(lotes_2) + 1

    def evaluar(lote):
        entrada = [{"n": f["n"], "numeral": f["numeral"], "titulo": f["titulo"],
                    "alcance_vigente": f["alcance_vigente"], "valor_exigido": f["valor_exigido"]}
                   | ({"puntaje_max": f["puntaje_max"], "criterio": f["criterio"]}
                      if f["es_factor"] else {})
                   for f in lote]
        prompt = PROMPT_2 + json.dumps(entrada, ensure_ascii=False)
        return lambda al_esperar: _como_dict(
            llamar_json(api_key, prompt, texto_pc, al_esperar, modelos), "evaluacion")

    respuestas = _en_paralelo([evaluar(l) for l in lotes_2],
                              f"Paso 2: evaluación de la oferta ({len(lotes_2)} lotes, "
                              f"{PARALELO} a la vez)", avisar)
    for lote, res in zip(lotes_2, respuestas):
        por_n = {d.get("n"): d for d in res.get("evaluacion") or [] if isinstance(d, dict)}
        for f in lote:
            d = por_n.get(f["n"])
            if d is None:
                d = {"veredicto": "NO EVIDENCIADO",
                     "motivo": "La IA no devolvió evaluación para este punto."}
            for campo in CAMPOS_EVAL:
                f[campo] = _txt(d.get(campo))
            v = _veredicto(d.get("veredicto"))
            f["veredicto"] = v if v != "NO APLICA" else "NO EVIDENCIADO"  # NO APLICA lo deciden las Bases
            f["puntaje_obt"] = _num(d.get("puntaje_obtenido")) if f["es_factor"] else None
            f["experiencia_total"] = _num(d.get("experiencia_total_anios")) if f["es_factor"] else None
            f["fuente_veredicto"] = _fuente_veredicto(f, d.get("fuente_veredicto"))
            if (f["veredicto"] == "SI" and f["fuente_veredicto"] == "Contrato"
                    and not f["motivo"].startswith("Acreditado en el Contrato")):
                f["motivo"] = "Acreditado en el Contrato: " + f["motivo"]
            if not hay_contrato:
                f["reflejo_contrato"] = "No se cargó el contrato"
    estado["hechos"] += len(lotes_2)
    medir("2. Evaluación de la oferta", t0, e0)

    aplicar_reglas(todos)
    marcados = control_coherencia(todos)

    # Paso 3 — síntesis (incluye la consolidación de mejoras)
    t0, e0 = time.time(), estadisticas()
    avisar("Paso 3: ficha resumen, mejoras, inconsistencias y recomendaciones…")
    resumen_eval = [{k: f[k] for k in ("numeral", "titulo", "valor_exigido", "valor_ofertado",
                                       "fuente_propuesta", "veredicto", "motivo",
                                       "reflejo_contrato", "mejora", "mejora_valor", "riesgo")}
                    for f in todos if f["veredicto"] != "NO APLICA"]
    sint = _como_dict(llamar_json(api_key, PROMPT_3 + json.dumps(resumen_eval, ensure_ascii=False),
                                  _docs(textos, ["Contrato", "Bases", "Pliego", "Propuesta"]),
                                  avisar, modelos), "inconsistencias")
    estado["hechos"] += 1
    medir("3. Síntesis", t0, e0)

    ficha_ia = {_sin_tildes(_txt(x.get("campo"))).lower(): x
                for x in sint.get("ficha") or [] if isinstance(x, dict)}
    ficha = []
    for campo in CAMPOS_FICHA:
        x = ficha_ia.get(_sin_tildes(campo).lower(), {})
        ficha.append({"campo": campo, "valor": _txt(x.get("valor")) or "No consignado",
                      "fuente": _txt(x.get("fuente"))})

    orden = {"alta": 0, "media": 1, "baja": 2, _sin_tildes(VERIFICAR_PDF).lower(): 3}
    inconsistencias = []
    for x in sint.get("inconsistencias") or []:
        if not isinstance(x, dict):
            continue
        prio = _sin_tildes(_txt(x.get("prioridad"))).lower()
        prio = (VERIFICAR_PDF if "verificar" in prio else prio.capitalize() if prio in orden
                else "Media")
        desc, docs = _txt(x.get("descripcion")), _txt(x.get("documentos"))
        if prio == "Alta" and _menciona(desc + " " + docs, marcados):
            prio = VERIFICAR_PDF   # no hay hallazgos Alta con cifras que no cuadran
        inconsistencias.append({"prioridad": prio, "descripcion": desc, "documentos": docs})
    inconsistencias.sort(key=lambda x: orden[_sin_tildes(x["prioridad"]).lower()])
    recomendaciones = [{"item": _txt(x.get("item")), "entregable": _txt(x.get("entregable")),
                        "detalle": _txt(x.get("detalle"))}
                       for x in sint.get("recomendaciones") or [] if isinstance(x, dict)]
    mejoras = (_normalizar_mejoras(sint.get("mejoras"), {f["numeral"] for f in todos})
               or _mejoras_por_texto(todos))
    mejoras = [m for m in mejoras if not _menciona(m["numerales"], marcados)]

    progreso("Listo.", 1.0)
    return {"filas": filas, "factores": factores, "mejoras": mejoras, "ficha": ficha,
            "inconsistencias": inconsistencias, "recomendaciones": recomendaciones,
            "duplicados_quitados": duplicados, "paginas_imagen": paginas_imagen,
            "avisos_imagen": avisos,
            "verificar_pdf": sorted(marcados), "tiempos": tiempos,
            "puntaje_minimo": puntaje_minimo,
            "fuente_puntaje_minimo": _txt(inv.get("fuente_puntaje_minimo")),
            "conteos": calcular_conteos(filas, factores, mejoras, puntaje_minimo)}

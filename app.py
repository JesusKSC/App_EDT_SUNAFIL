# -*- coding: utf-8 -*-
"""
==============================================================================
ANALIZADOR DE TDR Y PLAN DEL PROYECTO — SUNAFIL
Caso 1 — Administración de Proyectos Informáticos (URP)
------------------------------------------------------------------------------
Una sola app en 4 pasos:
    1) Documentos   -> lectura de PDF (texto digital u OCR)            lectura_pdf.py
    2) Checklist    -> análisis con Gemini + Excel                     checklist.py, excel_checklist.py
    3) EDT          -> estructura de desglose del trabajo              edt_builder.py
    4) Cronograma   -> actividades, fechas y costos                    cronograma_builder.py
(edt.py ya no se usa en la interfaz)
==============================================================================
"""
import datetime
import streamlit as st
import lectura_pdf
import checklist, excel_checklist, gemini_client
from gemini_client import ErrorGemini

st.set_page_config(page_title="Analizador TDR · SUNAFIL", page_icon="🧭", layout="wide")

# ---------------------------------------------------------------------------
# Estilos propios (los colores base vienen de .streamlit/config.toml)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.6rem; max-width: 1250px;}
.hero {background: linear-gradient(120deg, #0F766E 0%, #155E75 55%, #1E3A5F 100%);
       color: #FFFFFF; border-radius: 18px; padding: 26px 32px 22px; margin-bottom: 18px;}
.hero .eyebrow {text-transform: uppercase; letter-spacing: .14em; font-size: .78rem; opacity: .85;}
.hero h1 {color: #FFFFFF; font-size: 2.05rem; margin: 4px 0 6px; padding: 0;}
.hero p {margin: 0; opacity: .92; font-size: .98rem;}
.chips {display: flex; flex-wrap: wrap; gap: 8px; margin-top: 16px;}
.chip {background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.28);
       border-radius: 999px; padding: 5px 13px; font-size: .84rem;}
.chip.ok {background: #FFFFFF; color: #0F766E; font-weight: 600; border-color: #FFFFFF;}
.paso {display: flex; align-items: center; gap: 14px; margin: 2px 0 10px;}
.paso .num {background: #0F766E; color: #FFFFFF; width: 38px; height: 38px; border-radius: 12px;
            display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 1.1rem;
            flex-shrink: 0;}
.paso .num.ok {background: #16A34A;}
/* colores heredados del tema: se leen bien en modo claro y en modo oscuro */
.paso .txt b {font-size: 1.22rem; color: inherit;}
.paso .txt span {display: block; opacity: .68; font-size: .9rem;}
.doc-tag {font-weight: 600; color: inherit; font-size: .95rem;}
.side-titulo {font-weight: 700; color: #14B8A6; text-transform: uppercase; letter-spacing: .08em;
              font-size: .74rem; margin: 6px 0 2px;}
</style>
""", unsafe_allow_html=True)

def tabla_md(filas):
    """Tabla en Markdown. No usamos st.dataframe/st.table porque cargan pyarrow,
    cuya DLL bloquea Smart App Control de Windows."""
    if not filas:
        return ""
    cols = list(filas[0].keys())
    lineas = ["| " + " | ".join(cols) + " |", "|" + "---|" * len(cols)]
    for f in filas:
        lineas.append("| " + " | ".join(str(f[c]) for c in cols) + " |")
    return "\n".join(lineas)

def encabezado_paso(numero, titulo, descripcion, listo=False):
    st.markdown(f'<div class="paso"><div class="num{" ok" if listo else ""}">{"✓" if listo else numero}</div>'
                f'<div class="txt"><b>{titulo}</b><span>{descripcion}</span></div></div>',
                unsafe_allow_html=True)

# El banner y el progreso de la barra lateral se dibujan al final (cuando ya se sabe qué
# pasos se completaron en esta ejecución); aquí solo se reserva su lugar.
banner = st.empty()

# ---------------------------------------------------------------------------
# Barra lateral: conexión con Gemini (la API key solo vive en la sesión)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🧭 Analizador TDR")
    st.markdown('<div class="side-titulo">Conexión con Gemini</div>', unsafe_allow_html=True)
    st.text_input("🔑 API Key de Gemini", type="password", key="api_key",
                  help="Gratis en https://aistudio.google.com/apikey")

def _clave_del_servidor():
    """Clave guardada en los Secrets del servidor (GEMINI_API_KEY); nunca se muestra."""
    try:
        return str(st.secrets.get("GEMINI_API_KEY", "") or "")
    except Exception:   # sin archivo de secrets (uso local)
        return ""

# La que se escribe en la barra lateral manda; si está vacía, se usa la del servidor
API_KEY = st.session_state.get("api_key") or _clave_del_servidor()
if API_KEY and not st.session_state.get("api_key"):
    st.sidebar.caption("🔒 Usando la API key configurada en el servidor.")

# Modelos que la API key puede usar (se consulta una vez por clave)
_clave_api = API_KEY
if _clave_api and st.session_state.get("_modelos_de") != _clave_api:
    try:
        st.session_state["_modelos"] = gemini_client.listar_modelos(_clave_api)
    except Exception:
        st.session_state["_modelos"] = []
    st.session_state["_modelos_de"] = _clave_api
_disponibles = st.session_state.get("_modelos") or []
# solo los modelos que la API key tiene; el de por defecto va primero si está
_opciones = (sorted(_disponibles, key=lambda m: m != gemini_client.MODELO_GEMINI)
             or [gemini_client.MODELO_GEMINI])
if st.session_state.get("modelo") not in _opciones:
    st.session_state.pop("modelo", None)

with st.sidebar:
    st.selectbox("🤖 Modelo de Gemini", _opciones, key="modelo",
                 help="Solo aparecen los modelos que tu API key puede usar. "
                      "Cada modelo tiene su propia cuota gratuita diaria.")
    st.markdown('<div class="side-titulo">Opciones de análisis</div>', unsafe_allow_html=True)
    st.toggle("Probar otros modelos si el elegido está saturado o sin cuota",
              value=True, key="respaldo")
    st.toggle("Leer las tablas desde la imagen de la página (≈4 llamadas más)",
              value=True, key="tablas_imagen",
              help="Corrige las cifras que el OCR destroza en las tablas. "
                   "Desactívalo si te queda poca cuota de Gemini.")
    progreso_lateral = st.empty()
    st.caption("🌗 Modo claro / oscuro: menú ⋮ (arriba a la derecha) → Settings → "
               "Choose app theme. Por defecto sigue el modo de tu sistema.")
    with st.expander("❓ Cómo usar"):
        st.markdown("1. Sube las Bases, la Propuesta y el Pliego (el Contrato es opcional).\n"
                    "2. Genera el checklist con Gemini.\n"
                    "3. Genera la EDT desde ese checklist o desde un Excel que ya tengas.\n"
                    "4. Genera el cronograma con la fecha de firma y los plazos del contrato.")

# ---------------------------------------------------------------------------
# PASO 1 — DOCUMENTOS
# ---------------------------------------------------------------------------
# (etiqueta para las marcas "[Bases, p. 26]", título del campo, obligatorio)
DOCUMENTOS = [
    ("Bases",     "Bases Integradas — alcance / TDR", True),
    ("Propuesta", "Propuesta técnica del postor",     True),
    ("Pliego",    "Pliego de consultas y observaciones", True),
    ("Contrato",  "Contrato",                          False),
]
ICONO_DOC = {"Bases": "📘", "Propuesta": "📗", "Pliego": "📙", "Contrato": "📄"}

with st.container(border=True):
    encabezado_paso(1, "Documentos del procedimiento",
                    "Bases, Propuesta y Pliego son obligatorios; el Contrato es opcional. En PDFs "
                    "escaneados largos indica solo las páginas útiles (ej. el Capítulo III) y ahorras OCR.",
                    listo="textos" in st.session_state)
    entradas = {}   # etiqueta -> (archivo, rangos o None si el formato está mal)
    filas_docs = [st.columns(2), st.columns(2)]
    for i, (etq, titulo, oblig) in enumerate(DOCUMENTOS):
        with filas_docs[i // 2][i % 2].container(border=True):
            st.markdown(f'<span class="doc-tag">{ICONO_DOC[etq]} {titulo}'
                        f'{" · obligatorio" if oblig else " · opcional"}</span>', unsafe_allow_html=True)
            archivo = st.file_uploader(titulo, type=["pdf"], key=f"pdf_{etq}", label_visibility="collapsed")
            rango_txt = st.text_input(f"Páginas de {etq} a leer", key=f"pag_{etq}",
                                      placeholder="ej. 24-48, 58-66 · vacío = todas")
            rangos, error = lectura_pdf.validar_rango(rango_txt)
            if error:
                st.error(f"Páginas de {etq}: {error}")
            entradas[etq] = (archivo, rangos)

    if st.button("📖 Leer documentos", type="primary"):
        faltan = [etq for etq, _, oblig in DOCUMENTOS if oblig and not entradas[etq][0]]
        mal_formato = [etq for etq, (_, rangos) in entradas.items() if rangos is None]
        if faltan:
            st.error("Falta subir: " + ", ".join(faltan) + ".")
        elif mal_formato:
            st.error("Corrige el rango de páginas de: " + ", ".join(mal_formato) + ".")
        else:
            # Primero calculamos qué páginas leer de cada PDF (y validamos contra su total)
            plan, error = [], None
            for etq, (archivo, rangos) in entradas.items():
                if not archivo:
                    continue
                datos = archivo.getvalue()
                try:
                    paginas = lectura_pdf.paginas_a_leer(rangos, lectura_pdf.contar_paginas(datos))
                except ValueError as e:
                    error = f"Páginas de {etq}: {e}."
                    break
                except Exception as e:
                    error = f"No se pudo abrir el PDF de {etq}: {e}"
                    break
                plan.append((etq, datos, paginas))

            if error:
                st.error(error)
            else:
                total = sum(len(p) for _, _, p in plan)
                barra = st.progress(0.0, text="Leyendo documentos…")
                hechas = [0]   # lista para poder actualizarlo desde avanzar()
                textos, resumen = {}, []
                try:
                    for etq, datos, paginas in plan:
                        def avanzar(etq=etq):
                            hechas[0] += 1
                            barra.progress(hechas[0] / total,
                                           text=f"Leyendo {etq}… ({hechas[0]}/{total} páginas)")
                        textos[etq], fila = lectura_pdf.leer_pdf(datos, etq, paginas, avanzar)
                        resumen.append(fila)
                    barra.progress(1.0, text=f"Listo: {total} páginas leídas.")
                    st.session_state["textos"] = textos
                    st.session_state["resumen_lectura"] = resumen
                except Exception as e:
                    barra.empty()
                    st.error(f"Error leyendo los PDF: {e}. Si el error menciona Tesseract, "
                             f"revisa TESSERACT_CMD en lectura_pdf.py.")

    if "resumen_lectura" in st.session_state:
        resumen = st.session_state["resumen_lectura"]
        cols_doc = st.columns(len(resumen))
        for col, fila in zip(cols_doc, resumen):
            col.metric(f"{ICONO_DOC.get(fila['Documento'], '📄')} {fila['Documento']}",
                       f"{fila['Páginas leídas']} págs.",
                       f"{fila['OCR']} con OCR · {fila['Texto nativo']} digitales",
                       delta_color="off", border=True)

# ---------------------------------------------------------------------------
# PASO 2 — CHECKLIST
# ---------------------------------------------------------------------------
ICONO = {"SI": "✅ SI", "NO": "❌ NO", "PARCIAL": "🟡 PARCIAL", "NO EVIDENCIADO": "⚪ NO EVIDENCIADO",
         "NO APLICA": "➖ NO APLICA"}

def _md(v):
    """Texto seguro para una celda de tabla Markdown."""
    return "" if v is None else str(v).replace("|", "\\|").replace("\n", " ")

def _filas_md(filas, columnas):
    """columnas: {encabezado: función(fila) -> valor}"""
    return [{h: _md(fn(f)) for h, fn in columnas.items()} for f in filas]

with st.container(border=True):
    encabezado_paso(2, "Checklist del TDR",
                    "Gemini cruza las Bases y el Pliego con la Propuesta y el Contrato; Python aplica "
                    "las reglas fijas y calcula los conteos y el puntaje.",
                    listo="resultado_checklist" in st.session_state)
    falta = []
    if not API_KEY:
        falta.append("ingresa la API key de Gemini en la barra lateral")
    if not st.session_state.get("textos"):
        falta.append("lee los documentos en el paso 1")
    if falta:
        st.caption("Para habilitar el botón: " + " y ".join(falta) + ".")

    if st.button("🚀 Procesar y generar checklist", type="primary", disabled=bool(falta)):
        barra = st.progress(0.0, text="Iniciando…")
        modelos = [st.session_state["modelo"]]
        if st.session_state["respaldo"]:
            modelos += gemini_client.modelos_de_respaldo(modelos[0], _disponibles)
        # PDFs subidos en el paso 1: para leer las tablas desde la imagen de la página
        pdfs = {etq: st.session_state[f"pdf_{etq}"].getvalue()
                for etq, _, _ in DOCUMENTOS
                if st.session_state["tablas_imagen"] and st.session_state.get(f"pdf_{etq}")}
        try:
            res = checklist.generar_checklist(
                API_KEY, st.session_state["textos"],
                lambda texto, frac: barra.progress(min(frac, 1.0), text=texto), modelos, pdfs)
            barra.progress(1.0, text="Armando el Excel…")
            st.session_state["excel_checklist"] = excel_checklist.construir_excel(
                res, st.session_state["textos"])
            st.session_state["resultado_checklist"] = res
            st.session_state["checklist"] = res["filas"]     # para la EDT (hoja B)
            st.session_state["factores"] = res["factores"]   # hoja B2
            barra.progress(1.0, text=f"Listo: {res['conteos']['total']} puntos analizados.")
        except (ErrorGemini, ValueError) as e:
            barra.empty(); st.error(str(e))
        except Exception as e:
            barra.empty(); st.error(f"Error generando el checklist ({type(e).__name__}): {e}")

    if "tiempos" not in st.session_state.get("resultado_checklist", {"tiempos": 1}):
        # resultado de una versión anterior de la app: se descarta y se pide regenerar
        del st.session_state["resultado_checklist"]
        st.info("El checklist cambió de formato. Vuelve a presionar «Procesar y generar checklist».")

    if "resultado_checklist" in st.session_state:
        res = st.session_state["resultado_checklist"]
        c = res["conteos"]
        puntaje = f"{c['puntaje_obt']:g}/{c['puntaje_max']:g}" if c["puntaje_max"] else "—"
        fila1, fila2 = st.columns(4), st.columns(3)
        fila1[0].metric("Puntos del TDR rastreados", c["total"], border=True,
                        help="Requisitos extraídos del Capítulo III (Bases + Pliego), sin duplicados.")
        fila1[1].metric("Acreditados con dato", c["si"], border=True,
                        help="Veredicto SI: la propuesta trae el valor que pide el TDR, igual o mejor.")
        fila1[2].metric("Con hallazgo", c["no"], border=True,
                        help="Veredicto NO: el valor ofertado existe y no alcanza lo exigido.")
        fila1[3].metric("Parciales", c["parcial"], border=True,
                        help="Veredicto PARCIAL: cubre solo una parte.")
        fila2[0].metric("Pendientes de evidencia", c["no_evidenciado"], border=True,
                        help="NO EVIDENCIADO: el expediente no lo acredita en las páginas leídas.")
        fila2[1].metric("Mejoras ofertadas", c["mejoras"], border=True,
                        help="Compromisos por encima del mínimo, sustentados con una cifra.")
        fila2[2].metric("Puntaje de la oferta", puntaje, border=True,
                        help="Suma de los factores de evaluación (Capítulo IV, hoja B2), calculada en Python.")
        st.progress(min(max(c["cobertura"], 0.0), 1.0),
                    text=f"Cobertura documental: {c['cobertura']:.0%} de los {c['aplicables']} puntos "
                         f"que aplican tienen un dato concreto · Cumplimiento: {c['cumplimiento']:.0%}")

        av1, av2 = st.columns(2)
        with av1:
            if c["no"]:
                st.error("Puntos del TDR marcados como incumplidos (NO):\n\n" +
                         "\n".join(f"- {_md(x)}" for x in c["items_no"]))
            else:
                st.success("Ningún punto del TDR quedó marcado como incumplido.")
        with av2:
            if c["puntaje_minimo"] is None:
                st.info("Las Bases leídas no indican un puntaje mínimo; no se puede comparar.")
            elif c["puntaje_obt"] >= c["puntaje_minimo"]:
                st.success(f"Alcanza el puntaje mínimo ({c['puntaje_minimo']:g}).")
            else:
                st.error(f"No alcanza el puntaje mínimo ({c['puntaje_obt']:g} de {c['puntaje_minimo']:g}).")
        for aviso in res.get("avisos_imagen") or []:
            st.warning("📷 " + aviso)
        if res.get("verificar_pdf"):
            st.warning("Cifras que no cuadran entre la mejora, lo exigido y el Contrato (riesgo "
                       "«Verificar en el PDF»; no se cuentan como mejora ni como hallazgo): "
                       + ", ".join(res["verificar_pdf"]))

        st.download_button("⬇️ Descargar Excel (checklist + hojas A–F)",
                           key="descargar_checklist",
                           data=st.session_state["excel_checklist"],
                           file_name=excel_checklist.NOMBRE_ARCHIVO,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        t_res, t_chk, t_fac, t_fic, t_inc, t_tie = st.tabs(
            ["📊 Cómo leer los números", "✅ Checklist (B)", "🏅 Factores (B2)", "🗂️ Ficha (A)",
             "⚠️ Inconsistencias (D)", "⏱️ Tiempos"])
        with t_res:
            nota = [f"{c['no_aplica']} no aplican según las Bases y no cuentan en el porcentaje"] if c["no_aplica"] else []
            if res.get("duplicados_quitados"):
                nota.append(f"se quitaron {len(res['duplicados_quitados'])} filas 3.3.* que repetían el TDR (3.4.*)")
            if nota:
                st.caption("Además, " + " y ".join(nota) + ".")
            if res.get("paginas_imagen"):
                st.caption("📷 Tablas leídas desde la imagen de la página (no del OCR): " + "; ".join(
                    f"{doc} p. {', '.join(map(str, pags))}" for doc, pags in res["paginas_imagen"].items()))
            st.markdown("**Este panel audita la trazabilidad de la documentación; no descalifica al postor.**")
            st.markdown(tabla_md([
                {"Indicador": "Puntos del TDR rastreados",
                 "Qué dice": "Cuántos requisitos exigibles se extrajeron del alcance (Bases + consultas del Pliego), sin duplicados y sin los factores de evaluación (están en la hoja B2).",
                 "Qué NO dice": "Nada sobre la oferta."},
                {"Indicador": "No aplica",
                 "Qué dice": "Las Bases dicen expresamente que el punto no aplica (adelanto directo, reajustes, etc.). Se excluye del porcentaje de cumplimiento y de la cobertura.",
                 "Qué NO dice": "Que falte algo: no hay nada que acreditar."},
                {"Indicador": "Acreditados con dato",
                 "Qué dice": "Se encontró en la propuesta el valor que pide el TDR y es igual o superior.",
                 "Qué NO dice": "Que el servicio ya se esté prestando así: eso se verifica en la ejecución."},
                {"Indicador": "Con hallazgo",
                 "Qué dice": "El valor ofertado existe y no alcanza lo exigido. Solo aquí se afirma un incumplimiento, con la comparación a la vista.",
                 "Qué NO dice": "—"},
                {"Indicador": "Parciales",
                 "Qué dice": "La oferta cubre solo una parte del requisito.",
                 "Qué NO dice": "Que el resto incumpla: puede estar en páginas no leídas."},
                {"Indicador": "Pendientes de evidencia",
                 "Qué dice": "El requisito no aparece acreditado en las páginas que se cargaron.",
                 "Qué NO dice": "**No** significa que el postor incumpla: significa que el expediente no lo acredita en lo leído."},
                {"Indicador": "Mejoras ofertadas",
                 "Qué dice": "Compromisos por encima del mínimo, con cifra; pasan a ser obligación contractual.",
                 "Qué NO dice": "—"},
                {"Indicador": "Puntaje de la oferta",
                 "Qué dice": "Suma de los factores de evaluación según la evidencia encontrada.",
                 "Qué NO dice": "El puntaje oficial del comité de selección."},
            ]))
            st.markdown("**Por qué la cifra de pendientes suele ser alta en una oferta ganadora.** "
                        "La propuesta técnica no repite todos los requisitos del TDR: acredita con detalle "
                        "lo que puntúa y cubre el resto con una declaración de sometimiento. Esa declaración "
                        "obliga al postor, pero no es un dato verificable, así que el sistema la deja "
                        "pendiente. Si quieres bajar esa cifra, amplía el rango de páginas de la propuesta "
                        "y de sus anexos.")
        with t_chk:
            st.markdown(tabla_md(_filas_md(res["filas"], {
                "N°": lambda f: f["n"], "Numeral": lambda f: f["numeral"], "Ítem": lambda f: f["titulo"],
                "¿Cumple?": lambda f: ICONO[f["veredicto"]], "Fuente": lambda f: f.get("fuente_veredicto"),
                "Exigido": lambda f: f.get("valor_exigido"), "Ofertado": lambda f: f.get("valor_ofertado"),
                "Motivo": lambda f: f.get("motivo")})))
        with t_fac:
            st.markdown(tabla_md(_filas_md(res["factores"], {
                "Factor": lambda f: f["numeral"], "Descripción": lambda f: f["titulo"],
                "Exigido": lambda f: f.get("valor_exigido"), "Ofertado": lambda f: f.get("valor_ofertado"),
                "Puntaje": lambda f: f"{f['puntaje_obt']:g} / {f['puntaje_max']:g}",
                "Motivo": lambda f: f.get("motivo")})))
        with t_fic:
            st.markdown(tabla_md(_filas_md(res["ficha"], {
                "Campo": lambda x: x["campo"], "Valor": lambda x: x["valor"],
                "Fuente": lambda x: x["fuente"]})))
        with t_inc:
            if res["inconsistencias"]:
                st.markdown(tabla_md(_filas_md(res["inconsistencias"], {
                    "Prioridad": lambda x: x["prioridad"], "Descripción": lambda x: x["descripcion"],
                    "Documentos en conflicto": lambda x: x["documentos"]})))
            else:
                st.write("No se registraron inconsistencias.")
        with t_tie:
            total_s = sum(t["Tiempo (s)"] for t in res["tiempos"])
            cache_n = sum(t["Desde la caché"] for t in res["tiempos"])
            api_n = sum(t["Llamadas a Gemini"] for t in res["tiempos"])
            st.markdown(tabla_md(res["tiempos"]))
            st.caption(f"Total: {total_s:.0f} s · {api_n} llamadas a Gemini · "
                       f"{cache_n} respuestas salieron de la caché.")

# ---------------------------------------------------------------------------
# PASO 3 — EDT
# ---------------------------------------------------------------------------
try:   # edt_builder usa pandas: si Windows lo bloquea, los pasos 1 y 2 siguen funcionando
    import edt_builder
    _error_edt = None
except Exception as e:
    edt_builder, _error_edt = None, e

def _checklist_a_dataframe(filas):
    """Checklist del paso 2 -> DataFrame con las mismas columnas que la hoja B del Excel."""
    datos = [[f["n"], f["numeral"], f["titulo"], f.get("clase", ""), f.get("alcance_bases", ""),
              f.get("fuente_bases", ""), f.get("consulta_pliego", ""), f.get("valor_exigido", ""),
              f.get("alcance_final", ""), f.get("valor_ofertado", ""), f.get("fuente_propuesta", ""),
              f["veredicto"], f.get("fuente_veredicto", ""), f.get("motivo", ""),
              f.get("reflejo_contrato", ""), f.get("mejora", ""), f.get("riesgo", "")] for f in filas]
    return edt_builder.pd.DataFrame(datos, columns=excel_checklist.COLS_B)

df_alcance, res_edt, archivos, firma_edt, edt_lista = None, None, None, None, False
with st.container(border=True):
    encabezado_paso(3, "EDT — Estructura de desglose del trabajo",
                    "Organiza el alcance en entregables y paquetes de trabajo (PMBOK) y verifica la "
                    "regla del 100%.", listo=bool(st.session_state.get("edt_archivos")))
    if _error_edt:
        st.error(f"No se pudo cargar el módulo de la EDT ({type(_error_edt).__name__}: {_error_edt}). "
                 f"Si menciona una «directiva de Control de aplicaciones», Windows bloqueó pandas: "
                 f"reinicia la app más tarde.")
    if edt_builder:
        fuente_alcance = st.radio("¿De dónde tomo el alcance?",
                                  ["Del checklist que acabo de generar", "De un Excel de checklist que ya tengo"],
                                  horizontal=True, key="fuente_alcance")
        try:
            if fuente_alcance == "Del checklist que acabo de generar":
                if st.session_state.get("checklist"):
                    df_alcance = edt_builder.normalizar_checklist(
                        _checklist_a_dataframe(st.session_state["checklist"]))
                else:
                    st.warning("Todavía no hay un checklist generado: genéralo en el paso 2 "
                               "o elige «De un Excel de checklist que ya tengo».")
            else:
                excel_alcance = st.file_uploader("📄 Excel del checklist", type=["xlsx"], key="excel_alcance")
                if excel_alcance:
                    df_alcance = edt_builder.leer_checklist_excel(excel_alcance)
        except Exception as e:
            st.error(f"No se pudo leer el checklist: {e}")

    if df_alcance is not None:
        res_edt = edt_builder.construir_edt(df_alcance)
        met = res_edt["metricas"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ítems del alcance", met["Ítems del alcance"], border=True,
                  help="Filas del checklist (hoja B): cada requisito del TDR.")
        c2.metric("Se vuelven entregables", met["Se vuelven entregables"], border=True,
                  help="Ítems asignados a un paquete de trabajo de la EDT.")
        c3.metric("Fuera, con motivo escrito", met["Fuera, con motivo escrito"], border=True,
                  help="Ítems que no son trabajo a entregar; su motivo queda en la hoja «Fuera de la EDT».")
        c4.metric("Paquetes de trabajo", met["Paquetes de trabajo"], border=True,
                  help="Último nivel de la EDT (cuenta de control).")

        n_dentro, n_fuera = len(res_edt["asignados"]), len(res_edt["fuera"])
        if res_edt["cumple_100"]:
            st.success(f"Regla del 100%: {n_dentro} + {n_fuera} = {res_edt['items']}. Cada ítem del alcance "
                       f"está dentro de la EDT o tiene justificación escrita en la hoja «Fuera de la EDT». "
                       f"Ninguno se perdió.")
        if res_edt["sin"]:
            st.error("Ítems sin clasificar (ninguna regla los ubicó; quedan en la hoja «Sin clasificar»):\n\n"
                     + "\n".join(f"- {_md(it['numeral'])} {_md(it['titulo'])}" for it in res_edt["sin"]))
        if res_edt["vacios"]:
            st.info("Paquetes sin ítem directo (justificados en el diccionario): "
                    + ", ".join(res_edt["vacios"]))

        with st.expander("📖 Qué son estos entregables y por qué quedan 'fuera' algunos ítems"):
            st.markdown(
                "- Según el **PMBOK**, la EDT solo contiene **entregables** (sustantivos): lo que se "
                "entrega o se produce. Las actividades (verbos) van en el cronograma.\n"
                "- **Regla del 100%:** la EDT incluye todo el trabajo del alcance y solo ese trabajo. "
                "Por eso cada ítem del checklist termina en un paquete de trabajo o en la hoja "
                "«Fuera de la EDT», nunca se pierde.\n"
                "- **Lo que queda fuera** no es trabajo a entregar: condiciones contractuales "
                "(penalidades, reajustes, arbitraje), plazos y lugar (alimentan el cronograma), "
                "contexto o justificación del requerimiento, o factores de evaluación de la oferta. "
                "Cada uno tiene su motivo escrito y dónde se registra.\n"
                "- Algunos paquetes no tienen un numeral propio del TDR: son **derivados** (por ejemplo "
                "el plan de trabajo) y su justificación está en el diccionario de la EDT.")

        firma_edt = (res_edt["items"], tuple(str(x) for x in df_alcance["numeral"]))
        if st.button("🧩 Generar EDT", type="primary", key="generar_edt"):
            try:
                st.session_state["edt_archivos"] = {
                    "firma": firma_edt, "excel": edt_builder.excel_edt(res_edt),
                    "pptx": edt_builder.pptx_edt(res_edt)}
                st.session_state["edt"] = res_edt          # para el cronograma
                st.session_state.pop("cronograma", None)   # el cronograma anterior era de otra EDT
            except Exception as e:
                st.error(f"No se pudo generar la EDT ({type(e).__name__}): {e}")
        archivos = st.session_state.get("edt_archivos")
        # la EDT cuenta como generada si se presionó «Generar EDT» con el checklist actual
        edt_lista = bool(archivos and archivos["firma"] == firma_edt and st.session_state.get("edt"))
        if edt_lista:
            d1, d2, _ = st.columns([1, 1, 2])
            d1.download_button("⬇️ EDT + diccionario (Excel)", data=archivos["excel"],
                               file_name="EDT_SUNAFIL.xlsx", key="descargar_edt_xlsx", width="stretch",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            d2.download_button("⬇️ Diagrama del árbol (PowerPoint)", data=archivos["pptx"],
                               file_name="EDT_diagrama.pptx", key="descargar_edt_pptx", width="stretch",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            with st.expander("🌳 Ver la EDT en pantalla", expanded=True):
                tabla = edt_builder.tabla_edt(res_edt)
                try:
                    st.dataframe(tabla, hide_index=True, width="stretch")
                except Exception:   # pyarrow bloqueado por Windows: misma tabla en Markdown
                    filas_md = tabla.to_dict("records")
                    for f in filas_md:
                        f["Entregable"] = f["Entregable"].replace("   ", "\u00a0\u00a0\u00a0")
                    st.markdown(tabla_md(filas_md))

# ---------------------------------------------------------------------------
# PASO 4 — CRONOGRAMA (debajo de la EDT)
# ---------------------------------------------------------------------------
with st.container(border=True):
    encabezado_paso(4, "Cronograma del proyecto",
                    "Convierte cada paquete de trabajo de la EDT en actividades (verbos) con duración, "
                    "fechas, hitos y costos, como en MS Project.",
                    listo=bool(st.session_state.get("cronograma") and edt_lista))
    if df_alcance is None:
        st.caption("Carga el alcance y genera la EDT en el paso 3 para habilitar el cronograma.")
    else:
        p1, p2, p3, p4 = st.columns(4)
        p1.date_input("Fecha de firma del contrato", value=datetime.date(2026, 5, 4),
                      format="DD/MM/YYYY", key="crono_firma")
        p2.number_input("Plazo de activación (días)", min_value=1, value=14, step=1, key="crono_plazo")
        p3.number_input("Plazo del servicio (días)", min_value=1, value=731, step=1, key="crono_servicio")
        p4.number_input("Monto del contrato (S/)", min_value=0.0, value=1614806.00, step=1000.0,
                        format="%.2f", key="crono_monto")
        if st.button("📅 Generar cronograma", type="primary", key="generar_cronograma"):
            if not edt_lista:
                st.warning("Primero genera la EDT (botón «Generar EDT» del paso 3).")
            else:
                try:
                    import cronograma_builder
                    cr = cronograma_builder.construir_cronograma(
                        st.session_state["crono_firma"], int(st.session_state["crono_plazo"]),
                        float(st.session_state["crono_monto"]), int(st.session_state["crono_servicio"]),
                        res_edt=st.session_state["edt"])
                    st.session_state["cronograma"] = cr
                    st.session_state["cronograma_excel"] = cronograma_builder.excel_cronograma(cr)
                except Exception as e:
                    st.error(f"No se pudo generar el cronograma ({type(e).__name__}): {e}")
        st.caption("Las actividades (verbos) abren cada paquete de la EDT como en MS Project. Las fechas "
                   "del Excel son fórmulas: si cambias la fecha de firma (celda amarilla de la hoja "
                   "Resumen), todo se recalcula.")

        cr = st.session_state.get("cronograma")
        if cr and edt_lista:
            mc = cr["metricas"]
            for fila_m in (["Acta de implementación", "Días de activación usados", "Inicio del servicio",
                            "Fin del servicio"],
                           ["Cierre del contrato", "Actividades", "Hitos", "Costo total (S/)"]):
                for col, nombre in zip(st.columns(4), fila_m):
                    valor = mc[nombre]
                    if nombre == "Costo total (S/)":
                        valor = f"S/ {valor:,.2f}"
                    col.metric(nombre, valor, border=True)
            dias_usados = int(str(mc["Días de activación usados"]).split()[0])
            if cr["cumple_activacion"]:
                st.success(f"Cumple el plazo de activación: el acta de implementación queda dentro de los "
                           f"{cr['plazo_activacion']} días.")
            else:
                exceso = dias_usados - cr["plazo_activacion"]
                st.error(f"No cumple el plazo de activación: el acta de implementación se pasa {exceso} "
                         f"día{'s' if exceso != 1 else ''} de los {cr['plazo_activacion']} días "
                         f"({dias_usados} días usados).")
            st.info("🛣️ Ruta crítica: " + cr["ruta_critica"])
            st.download_button("⬇️ Cronograma (Excel)", data=st.session_state["cronograma_excel"],
                               file_name="CRONOGRAMA_SUNAFIL.xlsx", key="descargar_cronograma",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with st.expander("🗓️ Ver el cronograma en pantalla"):
                try:
                    st.dataframe(cr["tabla"], hide_index=True, height=450, width="stretch")
                except Exception:   # pyarrow bloqueado por Windows: misma tabla en Markdown
                    filas_md = cr["tabla"].to_dict("records")
                    for f in filas_md:
                        f["Tarea"] = f["Tarea"].replace("   ", "\u00a0\u00a0\u00a0")
                    st.markdown(tabla_md(filas_md))

# ---------------------------------------------------------------------------
# Banner y progreso (al final, con el estado ya actualizado en esta ejecución)
# ---------------------------------------------------------------------------
pasos = [("Documentos", "textos" in st.session_state),
         ("Checklist", "resultado_checklist" in st.session_state),
         ("EDT", edt_lista),
         ("Cronograma", bool(st.session_state.get("cronograma") and edt_lista))]
chips = "".join(f'<span class="chip{" ok" if hecho else ""}">{"✓" if hecho else k}&nbsp; {nombre}</span>'
                for k, (nombre, hecho) in enumerate(pasos, 1))
banner.markdown(f"""
<div class="hero">
  <div class="eyebrow">SUNAFIL · Servicio de telefonía móvil · Administración de Proyectos Informáticos</div>
  <h1>Analizador de TDR y plan del proyecto</h1>
  <p>Lee los documentos del procedimiento (con OCR donde haga falta), arma el checklist del TDR con
     Gemini y, a partir de ese alcance, genera la EDT y el cronograma.</p>
  <div class="chips">{chips}</div>
</div>""", unsafe_allow_html=True)
with progreso_lateral.container():
    st.markdown('<div class="side-titulo">Progreso</div>', unsafe_allow_html=True)
    hechos = sum(h for _, h in pasos)
    st.progress(hechos / len(pasos), text=f"{hechos} de {len(pasos)} pasos completados")

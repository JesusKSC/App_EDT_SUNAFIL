# -*- coding: utf-8 -*-
"""
==============================================================================
GENERADOR AUTOMÁTICO DE EDT / WBS  (Parte 2)
Caso 1 — Administración de Proyectos Informáticos (URP)
------------------------------------------------------------------------------
Toma el Excel del ALCANCE (checklist maestro de la Parte 1) y, usando la API
de Gemini con un prompt maestro basado en el PMBOK, genera la Estructura de
Desglose de Trabajo (EDT / WBS) en un Excel con roll-up automático.

Flujo (topología):
    Excel del alcance  ->  lectura de ítems  ->  Gemini (prompt PMI/PMBOK)
    ->  EDT estructurada (JSON)  ->  Excel EDT + Diccionario
==============================================================================
"""
import io, json, time
import urllib.request, urllib.error
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Modelo de Gemini a usar. Si tu API key da error "model not found",
# cámbialo por uno disponible (p. ej. "gemini-2.0-flash" o "gemini-1.5-flash").
MODELO_GEMINI = "gemini-3.6-flash"

# ---------------------------------------------------------------------------
# 1. PROMPT MAESTRO (el mismo del archivo Prompt_EDT_SUNAFIL.txt)
# ---------------------------------------------------------------------------
PROMPT_MAESTRO = """
Actúa como un Director de Proyectos (Project Manager) certificado PMP por el PMI.
Transforma el ALCANCE ya definido de un proyecto (tabla que te entrego) en una
ESTRUCTURA DE DESGLOSE DE TRABAJO (EDT / WBS), siguiendo el PMBOK.

REGLAS PMI OBLIGATORIAS:
1. Solo ENTREGABLES (sustantivos). NUNCA actividades ni características del producto.
2. Dos ramas de primer nivel: 1.1 GESTIÓN DEL PROYECTO y 1.2 PRODUCTO / SERVICIO.
3. PAQUETE DE TRABAJO = última descomposición de un entregable; nivel más bajo,
   un solo responsable, donde se estiman tiempo y costo. Clasifícalo explícitamente.
4. REGLA DEL 100%: los hijos suman exactamente el padre; la EDT contiene TODO el
   trabajo del alcance y solo ese trabajo.
5. NUNCA 1 a 1: si un entregable se descompone, es en 2 o más hijos.
6. ALTO NIVEL, presentable a gerencia; agrupa lo similar (Gama A/B/C, etc.).
7. Conserva la trazabilidad al numeral del TDR.

DEVUELVE EXCLUSIVAMENTE un JSON válido (sin texto adicional, sin ```):
{
  "proyecto": "",
  "items": [
    {"cuenta_control":"1.1.1.1","nivel":4,"nombre":"","clasificacion":"Paquete de trabajo",
     "rama":"Gestión|Producto","tdr":"","responsable":"","duracion_dias":0,"costo_estim":0,"descripcion":""}
  ]
}
Solo los PAQUETES DE TRABAJO llevan duracion_dias, costo_estim y descripcion con
valor; los entregables intermedios y ramas los dejan en 0/"" (se calculan por fórmula).
Ordena los ítems en orden de árbol.
"""

# ---------------------------------------------------------------------------
# 2. LECTURA DEL ALCANCE (Excel de la Parte 1)
# ---------------------------------------------------------------------------
def leer_alcance(file_bytes):
    """Lee todas las hojas del Excel del alcance y las vuelve texto para el modelo."""
    xls = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)
    partes = []
    for hoja, df in xls.items():
        partes.append(f"===== HOJA: {hoja} =====")
        for _, fila in df.iterrows():
            celdas = [str(c) for c in fila.tolist() if str(c) != "nan"]
            if celdas:
                partes.append(" | ".join(celdas))
    return "\n".join(partes)

# ---------------------------------------------------------------------------
# 3. GEMINI — genera la EDT en JSON a partir del alcance
# ---------------------------------------------------------------------------
def generar_edt(api_key, texto_alcance):
    # Llamada directa a la API REST de Gemini (sin la librería google-generativeai,
    # que arrastra grpc y es bloqueada por Smart App Control de Windows).
    entrada = PROMPT_MAESTRO + "\n\n===== ALCANCE DEL PROYECTO =====\n" + texto_alcance[:180000]
    url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
           f"{MODELO_GEMINI}:generateContent?key={api_key}")
    cuerpo = json.dumps({
        "contents": [{"parts": [{"text": entrada}]}],
        "generationConfig": {"temperature": 0},
    }).encode("utf-8")
    req = urllib.request.Request(url, data=cuerpo,
                                 headers={"Content-Type": "application/json"})
    # Reintentos: el 503 ("high demand"/UNAVAILABLE) y el 429 (rate limit) son
    # temporales del servidor de Gemini; esperamos y volvemos a intentar.
    intentos = 5
    for i in range(intentos):
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                respuesta = json.loads(r.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as e:
            detalle = e.read().decode("utf-8", "ignore")
            if e.code in (429, 500, 503) and i < intentos - 1:
                espera = 3 * (i + 1)  # 3, 6, 9, 12 segundos
                st.write(f"⏳ Gemini saturado ({e.code}). Reintentando en {espera}s… "
                         f"(intento {i + 2} de {intentos})")
                time.sleep(espera)
                continue
            raise RuntimeError(f"La API de Gemini respondió {e.code}: {detalle}")
    raw = respuesta["candidates"][0]["content"]["parts"][0]["text"].strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw)

# ---------------------------------------------------------------------------
# 4. EXCEL — arma la EDT con roll-up (regla del 100%) + Diccionario
# ---------------------------------------------------------------------------
def construir_excel(data):
    NAVY,BLUE,NARANJA,L3A,L3B,YELLOW,WHITE = "1F3864","2E5496","C55A11","D6E4F0","FCE4D6","FFF2CC","FFFFFF"
    A=lambda **k: Font(name="Arial", **k)
    thin=Side(style="thin",color="BFBFBF"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    ctr=Alignment(wrap_text=True,vertical="center",horizontal="center")
    lft=Alignment(wrap_text=True,vertical="center",horizontal="left")

    wb=Workbook(); ws=wb.active; ws.title="EDT - WBS"; ws.sheet_view.showGridLines=False
    cols=["Cuenta de control","Nivel","Entregable / Paquete de trabajo","Clasificación PMI",
          "N° TDR","Responsable","Duración (días)","Costo estim. (S/)"]
    NC=len(cols)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=NC)
    t=ws.cell(1,1,"ESTRUCTURA DE DESGLOSE DE TRABAJO (EDT / WBS) — según PMBOK")
    t.font=A(bold=True,color="FFFFFF",size=14); t.fill=PatternFill("solid",fgColor=NAVY); t.alignment=ctr
    ws.row_dimensions[1].height=28
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=NC)
    s=ws.cell(2,1,f"{data.get('proyecto','Proyecto')} · Descomposición orientada a entregables · Regla del 100%")
    s.font=A(italic=True,size=9,color="404040"); s.alignment=ctr
    for j,h in enumerate(cols,1):
        c=ws.cell(3,j,h); c.font=A(bold=True,color="FFFFFF",size=10); c.fill=PatternFill("solid",fgColor=BLUE); c.alignment=ctr; c.border=bd
    ws.row_dimensions[3].height=30

    items=data.get("items",[])
    r=4; filas={}
    for it in items:
        niv=int(it.get("nivel",4)); ind="   "*(niv-1)
        clasif=it.get("clasificacion","")
        rama=it.get("rama","")
        es_pkg = "Paquete" in clasif
        vals=[it.get("cuenta_control",""),niv,ind+it.get("nombre",""),clasif,
              it.get("tdr","—"),it.get("responsable",""),
              it.get("duracion_dias","") if es_pkg else "",
              it.get("costo_estim","") if es_pkg else ""]
        # color de fila
        if niv==1: fill=NAVY; fcolor=WHITE; bold=True
        elif niv==2: fill=(BLUE if rama=="Gestión" else NARANJA); fcolor=WHITE; bold=True
        elif niv==3: fill=(L3A if rama=="Gestión" else L3B); fcolor="000000"; bold=True
        else: fill=None; fcolor="000000"; bold=False
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.border=bd; c.font=A(bold=bold,size=10 if j==3 else 9,color=fcolor)
            c.alignment = lft if j in (3,6) else ctr
            if fill: c.fill=PatternFill("solid",fgColor=fill)
            if j==8: c.number_format='#,##0'
            if j in (7,8) and es_pkg: c.fill=PatternFill("solid",fgColor=YELLOW)
        filas[it.get("cuenta_control","")]=(r,niv)
        r+=1
    last=r-1

    # roll-up regla del 100%
    def hijos(padre):
        return [rr for c,(rr,n) in filas.items() if c and c.startswith(padre+".") and c.count(".")==padre.count(".")+1]
    for c,(rr,n) in filas.items():
        if n in (2,3):
            hs=hijos(c)
            if hs:
                ws.cell(rr,7,f"=SUM(G{min(hs)}:G{max(hs)})")
                ws.cell(rr,8,f"=SUM(H{min(hs)}:H{max(hs)})")
    if "1" in filas:
        r1=filas["1"][0]
        ramas=[rr for c,(rr,n) in filas.items() if n==2]
        if ramas:
            ws.cell(r1,8,f"=SUM(H{min(ramas)}:H{max(ramas)})")

    widths=[13,7,44,22,13,22,13,14]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A4"

    # Hoja Diccionario
    ws2=wb.create_sheet("Diccionario EDT"); ws2.sheet_view.showGridLines=False
    c2=["Cuenta de control","Paquete de trabajo","Rama","N° TDR","Responsable","Descripción del alcance"]
    ws2.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(c2))
    h=ws2.cell(1,1,"DICCIONARIO DE LA EDT (documento complementario)")
    h.font=A(bold=True,color="FFFFFF",size=13); h.fill=PatternFill("solid",fgColor=NAVY); h.alignment=ctr
    for j,x in enumerate(c2,1):
        c=ws2.cell(2,j,x); c.font=A(bold=True,color="FFFFFF",size=10); c.fill=PatternFill("solid",fgColor=BLUE); c.alignment=ctr; c.border=bd
    rr=3
    for it in items:
        if "Paquete" in it.get("clasificacion",""):
            vals=[it.get("cuenta_control",""),it.get("nombre",""),it.get("rama",""),
                  it.get("tdr","—"),it.get("responsable",""),it.get("descripcion","")]
            for j,v in enumerate(vals,1):
                c=ws2.cell(rr,j,v); c.border=bd; c.font=A(size=9); c.alignment=lft if j in (2,6) else ctr
                if j==3: c.fill=PatternFill("solid",fgColor=(L3A if it.get("rama")=="Gestión" else L3B))
            rr+=1
    for i,w in enumerate([13,40,11,14,22,58],1): ws2.column_dimensions[get_column_letter(i)].width=w
    ws2.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# Generador de EDT / WBS — SUNAFIL (Parte 2)

Convierte el **alcance** (checklist maestro de la Parte 1) en una **Estructura
de Desglose de Trabajo (EDT / WBS)** aplicando el PMBOK, usando la API de Gemini.

## Requisitos
1. **Python 3.11+** (marca "Add Python to PATH" al instalar).
2. **API key de Gemini** (gratis): https://aistudio.google.com/apikey

## Instalación
```
python -m pip install -r requirements.txt
```

## Ejecución
```
streamlit run app.py
```
Se abre el navegador. Pega tu API key, sube el **Excel del alcance** (el
checklist), presiona **Generar EDT** y descarga el `EDT_WBS.xlsx`.

## Flujo
```
Excel del alcance -> lectura -> Gemini (prompt PMBOK) -> EDT (JSON) -> Excel EDT + Diccionario
```

## Qué aplica (PMBOK)
- Solo entregables (no actividades ni características).
- Dos ramas: 1.1 Gestión del Proyecto · 1.2 Producto/Servicio.
- Distingue Entregable vs Paquete de trabajo (última descomposición).
- Regla del 100%: los hijos suman el padre (roll-up automático por fórmula).
- Cuenta de control jerárquica (1, 1.1, 1.1.1) y trazabilidad al TDR.

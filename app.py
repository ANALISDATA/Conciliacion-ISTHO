"""Pantalla inicial: elegir entre los procesos de la Financiera.

Este archivo YA NO contiene la lógica de conciliación bancaria — se movió tal cual,
sin cambiar una sola línea, a `pages/1_🏦_Cruce_Bancario.py`. Este es solo el punto de
entrada con el menú; cada proceso vive en su propia página, con su propio motor y su
propia tabla en la base de datos (ver `conciliacion.py`/`db.py` para BANCARIO y
`conciliacion_dian.py`/`db_dian.py` para DIAN — no se tocan entre sí)."""
import os

import streamlit as st

from config import CLAVE_ACCESO
from excel_export import EMPRESA, LOGO_PATH
from ui import acceso_permitido, inject_css, menu_css, menu_encabezado, menu_pie, menu_tarjeta

st.set_page_config(page_title="Financiera ISTHO", layout="wide",
                   page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🔷")

inject_css()

if not acceso_permitido(CLAVE_ACCESO, EMPRESA, "Financiera ISTHO",
                        "Elige el proceso que necesitas"):
    st.stop()

menu_css()
menu_encabezado("Área financiera · ISTHO S.A.S.", "Panel de Conciliaciones",
                "Elige el proceso que vas a trabajar hoy")

col1, col2 = st.columns(2, gap="large")

with col1:
    menu_tarjeta(
        "bancario", "🏦", "Cruce Bancario",
        "Concilia el extracto del banco contra el libro auxiliar contable — "
        "por fecha, valor, nombre y número de manifiesto.",
        "Ingresar aquí →",
        on_click=lambda: st.switch_page("pages/1_🏦_Cruce_Bancario.py"),
    )

with col2:
    menu_tarjeta(
        "dian", "📄", "Cruce DIAN",
        "Cruza los documentos electrónicos de la DIAN contra las causaciones de "
        "Avansant — muestra qué facturas todavía no se han digitado.",
        "Ingresar aquí →",
        on_click=lambda: st.switch_page("pages/2_📄_Cruce_DIAN.py"),
    )

menu_pie(f"{EMPRESA} · Módulo de cruces financieros")

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
from ui import acceso_permitido, hero, inject_css, sidebar_brand

st.set_page_config(page_title="Financiera ISTHO", layout="wide",
                   page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🔷")

inject_css()

if not acceso_permitido(CLAVE_ACCESO, EMPRESA, "Financiera ISTHO",
                        "Elige el proceso que necesitas"):
    st.stop()

sidebar_brand("ISTHO S.A.S.", "Financiera")

hero("Financiera ISTHO", "Elige el proceso que quieres trabajar", empresa=EMPRESA)

st.write("")
col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("### 🏦 Cruce Bancario")
        st.caption("Cruza el extracto bancario contra el libro auxiliar contable — "
                   "por fecha, valor, nombre y número de manifiesto.")
        st.page_link("pages/1_🏦_Cruce_Bancario.py", label="Entrar a Cruce Bancario",
                     icon="🏦", use_container_width=True)

with col2:
    with st.container(border=True):
        st.markdown("### 📄 Cruce DIAN")
        st.caption("Cruza los documentos electrónicos de la DIAN contra las causaciones "
                   "de Avansant — muestra qué facturas todavía no se han digitado.")
        st.page_link("pages/2_📄_Cruce_DIAN.py", label="Entrar a Cruce DIAN",
                     icon="📄", use_container_width=True)

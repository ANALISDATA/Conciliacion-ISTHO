"""Datos de la empresa y clave de acceso.

Se leen de la **configuración privada** (`st.secrets`), no del código, para que el
repositorio pueda ser público sin exponer el NIT ni el número de cuenta.

- En el computador: se toman de `.streamlit/secrets.toml`, que está excluido de Git.
- En Streamlit Cloud: se pegan en *Settings → Secrets*, donde nadie más los ve.

Si no hay configuración, la app sigue funcionando con valores genéricos: así quien
descargue el código puede probarla sin datos de ninguna empresa real.
"""
import streamlit as st


def _leer(clave, por_defecto):
    """Lee un valor de la configuración privada; si no existe, devuelve el genérico.
    Se captura cualquier error porque Streamlit lanza excepción cuando no hay ningún
    archivo de secretos, que es justamente el caso de una copia recién descargada."""
    try:
        valor = st.secrets.get(clave)
        return valor if valor not in (None, "") else por_defecto
    except Exception:
        return por_defecto


EMPRESA = _leer("empresa", "Empresa")
NIT = _leer("nit", "—")
CUENTA_DEFECTO = _leer("cuenta", "Cuenta bancaria")

# Clave para entrar a la app. Si queda vacía no se pide contraseña (uso local).
CLAVE_ACCESO = _leer("password", "")

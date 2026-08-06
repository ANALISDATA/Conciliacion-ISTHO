"""Persistencia del estado de conciliación en Supabase.

Todo el trabajo (archivos cargados y la lista de cruces, incluida la conciliación
manual) vivía solo en `st.session_state`: si el servidor se reiniciaba o la sesión
del navegador se cerraba, se perdía y había que empezar de cero. Este módulo guarda
ese mismo estado en una tabla de Supabase cada vez que cambia, para poder retomarlo.

Si no hay credenciales de Supabase configuradas, todas las funciones no hacen nada
(o devuelven vacío): la app sigue funcionando solo con la sesión en memoria, como
antes, para que seguir usándola en el computador sin configurar nada no se rompa.
"""
import io

import pandas as pd
import streamlit as st

TABLA = "conciliaciones"


def _config(clave):
    """Lee de `st.secrets`; nunca lanza, porque Streamlit lanza excepción cuando
    no hay ningún archivo de secretos (el caso normal al correr en el computador)."""
    try:
        return st.secrets.get(clave, "")
    except Exception:
        return ""


@st.cache_resource
def _cliente():
    url, key = _config("supabase_url"), _config("supabase_key")
    if not url or not key:
        return None
    from supabase import create_client
    return create_client(url, key)


def disponible():
    return _cliente() is not None


def guardar_estado(periodo, estado):
    """Guarda (o actualiza) el estado completo de un período."""
    cliente = _cliente()
    if cliente is None or not periodo:
        return
    fila = {
        "periodo": periodo,
        "banco": estado["banco"].to_json(orient="split", date_format="iso"),
        "libro": estado["libro"].to_json(orient="split", date_format="iso"),
        "cruces": estado["cruces"],
        "posibles": estado["posibles"],
        "margen_valor": estado["margen_valor"],
        "tolerancia": estado["tolerancia"],
    }
    cliente.table(TABLA).upsert(fila, on_conflict="periodo").execute()


def listar_periodos():
    """Períodos guardados, del más reciente al más antiguo."""
    cliente = _cliente()
    if cliente is None:
        return []
    res = (cliente.table(TABLA).select("periodo, actualizado_en")
           .order("actualizado_en", desc=True).execute())
    return res.data


def cargar_estado(periodo):
    """Reconstruye el estado de un período guardado, listo para `st.session_state`."""
    cliente = _cliente()
    if cliente is None:
        return None
    res = cliente.table(TABLA).select("*").eq("periodo", periodo).execute()
    if not res.data:
        return None
    fila = res.data[0]
    return {
        "banco": pd.read_json(io.StringIO(fila["banco"]), orient="split"),
        "libro": pd.read_json(io.StringIO(fila["libro"]), orient="split"),
        "cruces": fila["cruces"],
        "posibles": fila["posibles"],
        "margen_valor": fila["margen_valor"],
        "tolerancia": fila["tolerancia"],
        "periodo": periodo,
    }

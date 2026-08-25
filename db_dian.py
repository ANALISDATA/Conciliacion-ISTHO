"""Persistencia del estado del cruce DIAN en Supabase.

Espejo de `db.py` (BANCARIO) en su forma de trabajar, pero en su **propia tabla**
(`conciliaciones_dian`) — cero columnas, cero filas y cero puntos de contacto en común
con `conciliaciones` (la tabla de BANCARIO). Así los dos procesos pueden guardar y
retomar su historial sin ningún riesgo de mezclarse.

Si no hay credenciales de Supabase configuradas, todas las funciones no hacen nada (o
devuelven vacío): el cruce DIAN sigue funcionando solo con la sesión en memoria.
"""
import io
from datetime import datetime

import pandas as pd
import streamlit as st

TABLA = "conciliaciones_dian"


def _config(clave):
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


def _fila_liviana(estado):
    """Lo que puede cambiar después del guardado inicial: la lista de cruces (automáticos
    y manuales) y los ambiguos que se van resolviendo a mano."""
    return {
        "cruces": estado["cruces"],
        "ambiguos": estado["ambiguos"],
        "duplicados": estado["duplicados"],
        "actualizado_en": datetime.now().isoformat(),
    }


def guardar_estado(periodo, estado):
    """Guarda el estado COMPLETO de un período, DIAN y Avansant incluidos. Se usa una sola
    vez, justo después de cruzar: esos dos archivos ya no cambian en lo que resta de la
    revisión, así que cada edición posterior (cruce manual) usa `guardar_cambios`.

    Nunca lanza: si la tabla `conciliaciones_dian` todavía no existe en Supabase (a
    diferencia de `conciliaciones`, de BANCARIO, esta es nueva y hay que crearla a mano),
    el cruce recién hecho no debe perderse solo porque no se pudo guardar el historial."""
    cliente = _cliente()
    if cliente is None or not periodo:
        return
    fila = _fila_liviana(estado)
    fila["periodo"] = periodo
    fila["dian"] = estado["dian"].to_json(orient="split", date_format="iso")
    fila["avansat"] = estado["avansat"].to_json(orient="split", date_format="iso")
    try:
        cliente.table(TABLA).upsert(fila, on_conflict="periodo").execute()
    except Exception:
        st.session_state["_dian_guardado_fallo"] = True


def guardar_cambios(periodo, estado):
    """Actualiza solo lo que puede cambiar tras el guardado inicial, SIN volver a mandar
    los archivos completos — ver el porqué en `db.guardar_cambios`, es la misma razón.
    Tampoco lanza, por la misma razón que `guardar_estado`."""
    cliente = _cliente()
    if cliente is None or not periodo:
        return
    fila = _fila_liviana(estado)
    try:
        cliente.table(TABLA).update(fila).eq("periodo", periodo).execute()
    except Exception:
        st.session_state["_dian_guardado_fallo"] = True


def listar_periodos():
    """Períodos guardados, del más reciente al más antiguo. Nunca lanza — ver
    `guardar_estado` para el porqué (la tabla puede no existir todavía)."""
    cliente = _cliente()
    if cliente is None:
        return []
    try:
        res = (cliente.table(TABLA).select("periodo, actualizado_en")
               .order("actualizado_en", desc=True).execute())
        return res.data
    except Exception:
        return []


def cargar_estado(periodo):
    """Reconstruye el estado de un período guardado, listo para `st.session_state`."""
    cliente = _cliente()
    if cliente is None:
        return None
    res = cliente.table(TABLA).select("*").eq("periodo", periodo).execute()
    if not res.data:
        return None
    fila = res.data[0]
    dian = pd.read_json(io.StringIO(fila["dian"]), orient="split")
    avansat = pd.read_json(io.StringIO(fila["avansat"]), orient="split")
    # Igual que en db.py: read_json no reconoce "fecha_emision"/"fecha_contable" como
    # fechas porque el nombre no matchea sus heurísticas en inglés, y load_dian()/
    # load_avansant() dejan esas columnas como `date` de Python (no Timestamp) — sin
    # repetir aquí ese paso, una conciliación retomada queda con un tipo de dato distinto
    # al de una recién cargada y las comparaciones de fecha truenan más adelante.
    dian["fecha_emision"] = pd.to_datetime(dian["fecha_emision"]).dt.date
    avansat["fecha_contable"] = pd.to_datetime(avansat["fecha_contable"]).dt.date
    return {
        "dian": dian,
        "avansat": avansat,
        "cruces": fila["cruces"],
        "ambiguos": fila["ambiguos"],
        "duplicados": fila["duplicados"],
        "periodo": periodo,
    }

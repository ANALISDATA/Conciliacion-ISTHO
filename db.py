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
from datetime import datetime

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


def _fila_liviana(estado):
    """Los campos que sí pueden cambiar después del guardado inicial. `actualizado_en` se
    fija a propósito en cada llamada: el upsert solo toca las columnas que le mandamos, así
    que sin esto la fecha se quedaba pegada en el primer guardado y nunca reflejaba los
    siguientes — inútil para saber cuál de dos guardados (de dos personas, por ejemplo) es
    el más reciente."""
    return {
        "cruces": estado["cruces"],
        "posibles": estado["posibles"],
        "margen_valor": estado["margen_valor"],
        "tolerancia": estado["tolerancia"],
        "saldo_inicial_banco": estado.get("saldo_inicial_banco"),
        "saldo_inicial_libro": estado.get("saldo_inicial_libro"),
        "cerrado": estado.get("cerrado", False),
        "saldo_final_banco": estado.get("saldo_final_banco"),
        "saldo_final_libro": estado.get("saldo_final_libro"),
        "cerrado_en": estado.get("cerrado_en"),
        "actualizado_en": datetime.now().isoformat(),
    }


def guardar_estado(periodo, estado):
    """Guarda el estado COMPLETO de un período, banco y libro incluidos. Se usa una sola
    vez, justo después de Conciliar: esos dos ya no cambian en lo que resta del mes, así que
    cada edición posterior (desconciliar, cruzar a mano, cerrar) usa `guardar_cambios` en vez
    de esta, para no volver a subir cientos de movimientos por cada clic."""
    cliente = _cliente()
    if cliente is None or not periodo:
        return
    fila = _fila_liviana(estado)
    fila["periodo"] = periodo
    fila["banco"] = estado["banco"].to_json(orient="split", date_format="iso")
    fila["libro"] = estado["libro"].to_json(orient="split", date_format="iso")
    cliente.table(TABLA).upsert(fila, on_conflict="periodo").execute()


def guardar_cambios(periodo, estado):
    """Actualiza solo lo que puede cambiar tras el guardado inicial (cruces, posibles,
    saldos, cierre), SIN volver a mandar banco/libro — eso ya quedó en Supabase desde el
    primer guardado. A propósito usa `update()` y no `upsert()`: el upsert arma un INSERT
    con TODAS las columnas de la tabla, y a la que no le mandamos valor la manda en NULL —
    o sea que hubiera borrado banco/libro en cada guardado liviano. `update()` sí modifica
    solo las columnas que se le pasan, dejando las demás intactas."""
    cliente = _cliente()
    if cliente is None or not periodo:
        return
    fila = _fila_liviana(estado)
    cliente.table(TABLA).update(fila).eq("periodo", periodo).execute()


def ultimo_cierre():
    """El cierre más reciente (una conciliación marcada como terminada a propósito), para
    sugerir su saldo final del banco como saldo inicial del mes siguiente. A diferencia del
    resto del estado, esto NO se guarda solo: el usuario decide cuándo un mes queda cerrado."""
    cliente = _cliente()
    if cliente is None:
        return None
    res = (cliente.table(TABLA).select("periodo, saldo_final_banco, saldo_final_libro, cerrado_en")
           .eq("cerrado", True).order("cerrado_en", desc=True).limit(1).execute())
    return res.data[0] if res.data else None


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
    # read_json (orient="split") no adivina que "fecha" es una fecha porque el nombre no
    # matchea sus heurísticas en inglés (date/time/_at); sin esto vuelve como texto y todo lo
    # que espera un Timestamp (el cálculo del período, los filtros de fecha) truena.
    #
    # El `.dt.date` final no es cosmético: `load_extracto`/`load_libro_auxiliar` (la carga
    # normal de archivos) dejan la columna "fecha" como objetos `date` de Python, no como
    # Timestamp. Sin repetir aquí ese mismo paso, una conciliación retomada queda con un tipo
    # de dato distinto al de una recién cargada, y comparar esa columna contra el rango de
    # fechas del filtro (que sí son `date`) truena con TypeError — solo al retomar, nunca en
    # una conciliación nueva, que es justo lo que lo hizo tan difícil de notar.
    banco = pd.read_json(io.StringIO(fila["banco"]), orient="split")
    libro = pd.read_json(io.StringIO(fila["libro"]), orient="split")
    banco["fecha"] = pd.to_datetime(banco["fecha"]).dt.date
    libro["fecha"] = pd.to_datetime(libro["fecha"]).dt.date
    return {
        "banco": banco,
        "libro": libro,
        "cruces": fila["cruces"],
        "posibles": fila["posibles"],
        "margen_valor": fila["margen_valor"],
        "tolerancia": fila["tolerancia"],
        "periodo": periodo,
        "saldo_inicial_banco": fila.get("saldo_inicial_banco"),
        "saldo_inicial_libro": fila.get("saldo_inicial_libro"),
        "cerrado": fila.get("cerrado", False),
        "saldo_final_banco": fila.get("saldo_final_banco"),
        "saldo_final_libro": fila.get("saldo_final_libro"),
        "cerrado_en": fila.get("cerrado_en"),
    }

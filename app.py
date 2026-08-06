import hmac
import os
from datetime import datetime

import streamlit as st

import db
from conciliacion import (construir_vistas, crear_cruce_manual, eliminar_cruces, load_extracto,
                           load_libro_auxiliar, reconciliar, resumen_cruces)
from config import CLAVE_ACCESO
from excel_export import (CUENTA_DEFECTO, EMPRESA, LOGO_PATH, NIT, build_tabla_workbook,
                           periodo_desde_fechas)
from ui import (GRIS, NARANJA, ROJO, VERDE, VERDE_OSC, badge, hero, inject_css, loader,
                 login_css, login_encabezado, login_pie, section, sidebar_brand, sidebar_step,
                 stat_cards, tabla)

# El ícono de la pestaña del navegador es el propio logo; si el archivo no estuviera,
# se cae a un rombo para que la app no falle al arrancar.
st.set_page_config(page_title="Conciliación Bancaria", layout="wide",
                   page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🔷")

inject_css()


def _fmt_dt(iso):
    """Convierte un timestamp ISO (como los que guarda Supabase) a dd/mm/aaaa hh:mm.
    Devuelve vacío si no hay valor o no se puede interpretar, en vez de lanzar."""
    if not iso:
        return ""
    try:
        return datetime.fromisoformat(str(iso)).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return ""


# ------------------------------------------------------------------ Acceso --
def acceso_permitido():
    """Puerta de entrada. La app queda publicada en una dirección pública, así que sin la
    clave no se muestra ningún dato. Si no hay clave configurada (uso en el computador),
    no se pide nada."""
    if not CLAVE_ACCESO:
        return True
    if st.session_state.get("acceso_ok"):
        return True

    login_css()
    _, centro, _ = st.columns([1, 1.45, 1])
    with centro:
        with st.container(key="login_card"):
            login_encabezado("Conciliación Bancaria",
                             "Cruce del extracto bancario contra el libro auxiliar")
            with st.form("acceso", border=False):
                clave = st.text_input("Clave de acceso", type="password",
                                      placeholder="Ingresa tu clave")
                if st.form_submit_button("Ingresar", use_container_width=True):
                    # compare_digest evita filtrar la clave por el tiempo de comparación
                    if hmac.compare_digest(clave, CLAVE_ACCESO):
                        st.session_state["acceso_ok"] = True
                        st.rerun()
                    else:
                        st.error("Clave incorrecta. Vuelve a intentarlo.")
        login_pie(EMPRESA)
    return False


if not acceso_permitido():
    st.stop()

# Si algún mes anterior quedó marcado como "terminado", sus dos saldos finales (banco y
# contabilidad, que no siempre son iguales si el mes no cerró perfecto) se sugieren como
# saldos iniciales del mes nuevo. Se hace una sola vez por sesión: en cuanto los widgets de
# abajo se dibujan, session_state ya tiene su propio valor y esto deja de aplicar.
if "saldo_inicial_banco" not in st.session_state and db.disponible():
    ultimo = db.ultimo_cierre()
    if ultimo:
        if ultimo.get("saldo_final_banco") is not None:
            st.session_state["saldo_inicial_banco"] = ultimo["saldo_final_banco"]
        if ultimo.get("saldo_final_libro") is not None:
            st.session_state["saldo_inicial_libro"] = ultimo["saldo_final_libro"]

# Streamlit no deja cambiar el valor guardado de un widget (ej. "saldo_inicial_banco")
# después de que ese widget ya se dibujó en la misma corrida del script. El botón
# "↻ Retomar" vive más abajo, después de que el campo de saldo inicial ya se dibujó, así
# que no puede tocar session_state directamente: deja la conciliación cargada "pendiente"
# aquí, ANTES de que se dibuje ningún widget, que es el único momento en que se puede.
if "_retomar_pendiente" in st.session_state:
    cargado = st.session_state.pop("_retomar_pendiente")
    st.session_state["estado"] = cargado
    if cargado:
        if cargado.get("saldo_inicial_banco") is not None:
            st.session_state["saldo_inicial_banco"] = cargado["saldo_inicial_banco"]
        if cargado.get("saldo_inicial_libro") is not None:
            st.session_state["saldo_inicial_libro"] = cargado["saldo_inicial_libro"]

# ---------------------------------------------------------------- Sidebar --
sidebar_brand()

sidebar_step(1, "Archivos")
st.sidebar.caption("Acepta el .xlsx o el .csv tal como se descarga del banco y del sistema contable.")
archivo_banco = st.sidebar.file_uploader("Extracto bancario (Bancolombia)", type=["xlsx", "xls", "csv"])
archivo_libro = st.sidebar.file_uploader("Libro auxiliar contable", type=["xlsx", "xls", "csv"])

sidebar_step(2, "Parámetros de cruce")
tolerancia = st.sidebar.slider(
    "Tolerancia de fecha (días)", min_value=0, max_value=10, value=3,
    help="Además del cruce exacto (misma fecha y mismo valor), se buscan coincidencias con "
         "el mismo valor dentro de este margen de días, quedándose con la fecha más cercana."
)
tolerancia_nombre = st.sidebar.slider(
    "Tolerancia de fecha cuando el nombre coincide (días)", min_value=tolerancia, max_value=30,
    value=max(15, tolerancia),
    help="Cuando el valor es exacto Y el nombre del beneficiario coincide, se acepta una diferencia "
         "de fecha más amplia que la tolerancia normal (ej: el banco gira el pago varios días después "
         "de la liquidación contable, pero es evidentemente el mismo movimiento)."
)

agrupar = st.sidebar.checkbox(
    "Cruzar movimientos agrupados (suma)", value=True,
    help="Cruza los casos en que un lado detalla y el otro consolida. Busca en tres niveles:\n\n"
         "• **Por nombre**: un anticipo y su pago sobre anticipo del mismo tercero, que en "
         "contabilidad son 2 líneas pero en el banco es 1 solo giro.\n\n"
         "• **Por concepto**: varios pagos de nómina, IVA, 4x1000 o comisiones del banco que "
         "suman un único asiento contable.\n\n"
         "• **Por lote de días**: la nómina completa que el banco pagó en uno o varios días "
         "contra el asiento global de la contabilidad."
)
max_grupo = st.sidebar.slider(
    "Máximo de movimientos por grupo", min_value=2, max_value=10, value=6, disabled=not agrupar,
    help="Cuántos movimientos como máximo se combinan para buscar que la suma coincida."
)

buscar_posibles = st.sidebar.checkbox(
    "Detectar posibles cruces con diferencia de valor", value=True,
    help="Busca pares que coinciden en fecha y en el nombre del beneficiario, pero cuyo valor no es "
         "idéntico (por un descuento, comisión o error de digitación). No se dan por conciliados: "
         "quedan en una hoja aparte para que los confirmes a mano."
)
margen_valor = st.sidebar.number_input(
    "Margen de valor permitido ($)", min_value=0, max_value=5_000_000, value=100_000, step=10_000,
    disabled=not buscar_posibles,
    help="Diferencia máxima en pesos para considerarlo un 'posible' cruce en vez de descartarlo."
)

sidebar_step(3, "Saldo inicial")
st.sidebar.caption("Saldo final del mes anterior de cada lado. Si el mes pasado quedó "
                   "totalmente conciliado deberían ser iguales; si no, cada uno trae su "
                   "propio arrastre.")
saldo_inicial_banco = st.sidebar.number_input(
    # Sin `value=`: el valor inicial lo pone session_state (arranca en 0 si no hay nada
    # guardado). Pasar `value` a la vez que el bloque de arriba ya fijó session_state hace
    # que Streamlit avise (sin romper nada) de que sobra uno de los dos.
    "Saldo inicial banco ($)", step=1_000_000.0, format="%.2f", key="saldo_inicial_banco",
    help="Se toma del extracto bancario del mes anterior («SALDO ANTERIOR»). Con este valor "
         "la app calcula Saldo final banco = Saldo inicial + Total ingresos − Total salidas."
)
saldo_inicial_libro = st.sidebar.number_input(
    "Saldo inicial contabilidad ($)", step=1_000_000.0, format="%.2f",
    key="saldo_inicial_libro",
    help="Saldo con el que el libro auxiliar traía la cuenta al empezar el mes. Se usa igual "
         "que el de arriba, pero del lado contable, para poder comparar los dos saldos finales."
)

ejecutar = st.sidebar.button("Conciliar", type="primary", use_container_width=True)

# El estado guarda los datos cargados y la LISTA DE CRUCES. Todas las tablas se derivan
# de ahí, así que conciliar o desconciliar a mano nunca altera los datos originales.
if "estado" not in st.session_state:
    st.session_state["estado"] = None
if "gen" not in st.session_state:
    st.session_state["gen"] = 0          # fuerza reiniciar la selección de las tablas

# Si hay una conciliación guardada en Supabase (de una sesión anterior, un reinicio del
# servidor, etc.) se ofrece retomarla en vez de tener que resubir los archivos y perder
# el trabajo ya hecho, incluidas las conciliaciones manuales.
if st.session_state["estado"] is None and db.disponible():
    periodos = db.listar_periodos()
    if periodos:
        st.sidebar.divider()
        st.sidebar.caption("☁️ Hay conciliaciones guardadas")
        # Se muestra la hora del último guardado junto a cada período: si dos personas
        # trabajaron el mismo mes, así se puede distinguir cuál versión es la más reciente
        # antes de decidir cuál retomar (dos guardados del mismo período se pisan entre sí,
        # gana el último — esto solo ayuda a saber cuál fue).
        opciones = {f"{p['periodo']} — actualizado {_fmt_dt(p['actualizado_en'])}": p["periodo"]
                    for p in periodos}
        etiqueta = st.sidebar.selectbox(
            "Retomar conciliación guardada", list(opciones.keys()),
            index=None, placeholder="Elige un período…", key="periodo_a_retomar",
        )
        elegido = opciones.get(etiqueta)
        if elegido and st.sidebar.button("↻ Retomar", use_container_width=True):
            # No se toca session_state["estado"] ni el saldo inicial aquí: el campo de saldo
            # inicial ya se dibujó más arriba en esta misma corrida, y Streamlit no deja
            # cambiar el valor de un widget después de dibujado. Se deja pendiente y el
            # bloque de arriba lo aplica en la siguiente corrida, antes de dibujar nada.
            st.session_state["_retomar_pendiente"] = db.cargar_estado(elegido)
            st.rerun()

if ejecutar:
    if not archivo_banco or not archivo_libro:
        st.sidebar.error("Debes cargar los dos archivos.")
    else:
        pantalla = st.empty()
        try:
            pantalla.markdown(loader("Leyendo el extracto bancario…",
                                     "Interpretando fechas, valores y descripciones"),
                              unsafe_allow_html=True)
            df_banco = load_extracto(archivo_banco)

            pantalla.markdown(loader("Leyendo el libro auxiliar…",
                                     f"{len(df_banco):,} movimientos del banco cargados"),
                              unsafe_allow_html=True)
            df_libro = load_libro_auxiliar(archivo_libro)

            pantalla.markdown(loader("Conciliando movimientos…",
                                     f"Cruzando {len(df_banco):,} movimientos del banco contra "
                                     f"{len(df_libro):,} de contabilidad · por fecha, valor y nombre"),
                              unsafe_allow_html=True)
            cruces, posibles = reconciliar(
                df_banco, df_libro, tolerancia_dias=tolerancia, tolerancia_dias_nombre=tolerancia_nombre,
                agrupar_por_fecha=agrupar, max_grupo=max_grupo,
                buscar_posibles=buscar_posibles, margen_valor=margen_valor,
            )
            periodo = periodo_desde_fechas(list(df_banco["fecha"]) + list(df_libro["fecha"]))
            st.session_state["estado"] = {
                "banco": df_banco, "libro": df_libro, "cruces": cruces, "posibles": posibles,
                "margen_valor": margen_valor, "tolerancia": tolerancia, "periodo": periodo,
                "saldo_inicial_banco": saldo_inicial_banco, "saldo_inicial_libro": saldo_inicial_libro,
                "cerrado": False, "saldo_final_banco": None, "saldo_final_libro": None, "cerrado_en": None,
            }
            db.guardar_estado(periodo, st.session_state["estado"])
            st.session_state["gen"] += 1
        except Exception as e:
            st.sidebar.error(f"Error procesando los archivos: {e}")
            st.session_state["estado"] = None
        finally:
            pantalla.empty()

est = st.session_state["estado"]

# ------------------------------------------------------------------ Hero --
periodo_hero = ""
if est is not None:
    periodo_hero = periodo_desde_fechas(list(est["banco"]["fecha"]) + list(est["libro"]["fecha"]))

if est is None:
    hero("Conciliación Bancaria",
         "Cruce automático entre el extracto bancario y el libro auxiliar contable — "
         "por fecha, valor y nombre.",
         empresa=EMPRESA, periodo="Sin periodo cargado")
    st.info("Carga el extracto bancario y el libro auxiliar en el panel izquierdo, ajusta los parámetros "
            "de cruce si lo necesitas, y presiona **Conciliar**.")
    st.stop()

df_banco, df_libro = est["banco"], est["libro"]
cruces, posibles = est["cruces"], est["posibles"]
margen_valor = est["margen_valor"]

df_conc, df_posibles, df_solo_banco, df_solo_libro = construir_vistas(
    df_banco, df_libro, cruces, posibles)

# --------------------------------------------------------- Barra superior --
HOJAS = ["📊 Resumen", "🔗 Conciliados", "🔍 Por revisar",
         "🏦 Pend. extracto", "📘 Pend. libro auxiliar", "⚖️ Cruce manual"]

# Primero la franja azul y debajo la barra de hojas. Como el encabezado necesita saber qué
# hoja está activa (para usar su versión delgada), se lee la selección ya guardada antes de
# dibujar la barra: Streamlit actualiza el estado del widget antes de reejecutar el script,
# así que al pulsar otra hoja el valor que se lee aquí ya es el nuevo.
vista = st.session_state.get("hoja_activa") or HOJAS[0]

hero("Conciliación Bancaria",
     "Cruce automático entre el extracto bancario y el libro auxiliar contable — por fecha, valor y nombre.",
     empresa=EMPRESA, periodo=periodo_hero, compacto=(vista != HOJAS[0]))

with st.container(key="nav_hojas"):
    seleccion = st.segmented_control("Hojas", HOJAS, default=HOJAS[0],
                                      label_visibility="collapsed", key="hoja_activa")
vista = seleccion or vista

# ------------------------------------------------------------ Indicadores --
total_banco = df_banco["valor"].sum()
total_libro = df_libro["valor"].sum()
total_solo_banco = df_solo_banco["valor"].sum() if not df_solo_banco.empty else 0.0
total_solo_libro = df_solo_libro["valor"].sum() if not df_solo_libro.empty else 0.0
n_conciliados = len(df_conc)
n_cruces = len(cruces)
n_manuales = sum(1 for c in cruces if c["origen"] == "Manual")
n_posibles = len(df_posibles)
diferencia = total_banco - total_libro
pct = (n_conciliados / len(df_banco) * 100) if len(df_banco) else 0

# Ingresos y salidas se determinan solo por el SIGNO del movimiento, sin mirar el concepto.
total_ingresos = float(df_banco.loc[df_banco["valor"] > 0, "valor"].sum())
total_salidas = float(-df_banco.loc[df_banco["valor"] < 0, "valor"].sum())
saldo_final = saldo_inicial_banco + total_ingresos - total_salidas

# Mismo cálculo con el lado contable, pero con su PROPIO saldo inicial: si el mes anterior no
# quedó perfectamente conciliado, el arrastre del banco y el de la contabilidad no tienen por
# qué coincidir, y forzarlos a ser iguales escondería esa diferencia en vez de mostrarla.
total_ingresos_libro = float(df_libro.loc[df_libro["valor"] > 0, "valor"].sum())
total_salidas_libro = float(-df_libro.loc[df_libro["valor"] < 0, "valor"].sum())
saldo_final_libro = saldo_inicial_libro + total_ingresos_libro - total_salidas_libro

# "diferencia" es solo del período (banco - contabilidad de este mes, sin arrastre) y se usa
# para validar que lo pendiente explica todo lo que no cruzó. "diferencia_saldos" es la cifra
# de fondo: cuánto se distancian los dos saldos finales, arrastre incluido.
total_dif_posibles = df_posibles["Diferencia"].sum() if not df_posibles.empty else 0.0
dif_esperada = (total_solo_banco - total_solo_libro) + total_dif_posibles
cuadra = abs(diferencia - dif_esperada) <= 0.01

diferencia_saldos = saldo_final - saldo_final_libro
saldos_cuadran = abs(diferencia_saldos) <= 0.01

# ------------------------------------------------------------------ Meta --
periodo = periodo_desde_fechas(list(df_banco["fecha"]) + list(df_libro["fecha"]))
cuenta = CUENTA_DEFECTO
if "cuenta_nombre" in df_libro.columns and not df_libro["cuenta_nombre"].dropna().empty:
    cuenta = df_libro["cuenta_nombre"].dropna().mode().iloc[0]
meta = {"empresa": EMPRESA, "nit": NIT, "cuenta": cuenta, "periodo": periodo,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "tolerancia": est["tolerancia"], "margen_valor": margen_valor}
sufijo = periodo.replace(" ", "_").replace("/", "-")


# --------------------------------------------------------------- Helpers --
def filtrar(df, key, columna_fecha, columnas_texto, columna_tipo=None, permitir_orden=False,
            columna_orden_alt=None):
    """Barra de filtros (texto + rango de fechas + tipo, y opcionalmente orden) reutilizable.

    Cada widget usa una key fija por hoja (no depende de `gen` ni de cuál hoja está activa),
    así que lo que busques, el rango de fechas y el orden elegido se mantienen aunque cambies
    de pestaña o hagas otra acción — hasta que tú mismo los borres.

    `permitir_orden` agrega un selector Fecha / Descripción (A-Z): ordenar por nombre ayuda a
    emparejar a simple vista los pendientes del banco contra los de contabilidad."""
    if df.empty:
        return df
    orden_col = (columna_orden_alt or (columnas_texto[0] if columnas_texto else None)) if permitir_orden else None
    with st.container(border=True, key=f"toolbar_{key}"):
        cols = st.columns([2.2, 1.6, 1, 1.3]) if orden_col else st.columns([2.5, 1.7, 1])
        col1, col2, col3 = cols[0], cols[1], cols[2]
        texto = col1.text_input("Buscar", placeholder="Nombre, descripción, comprobante...",
                                 key=f"buscar_{key}")
        resultado = df
        fechas = df[columna_fecha].dropna()
        if not fechas.empty and fechas.min() != fechas.max():
            rango = col2.date_input("Rango de fechas", value=(fechas.min(), fechas.max()),
                                     min_value=fechas.min(), max_value=fechas.max(), key=f"fechas_{key}")
            if isinstance(rango, tuple) and len(rango) == 2:
                resultado = resultado[(resultado[columna_fecha] >= rango[0])
                                       & (resultado[columna_fecha] <= rango[1])]
        if columna_tipo and columna_tipo in df.columns:
            opciones = ["Todos"] + sorted(df[columna_tipo].dropna().unique().tolist())
            sel = col3.selectbox("Tipo", opciones, key=f"tipo_{key}")
            if sel != "Todos":
                resultado = resultado[resultado[columna_tipo] == sel]
        orden = "Fecha"
        if orden_col:
            orden = cols[3].selectbox("Ordenar por", ["Fecha", "Descripción (A-Z)"], key=f"orden_{key}")
        if texto:
            mascara = None
            for c in columnas_texto:
                if c in resultado.columns:
                    coincide = resultado[c].astype(str).str.contains(texto, case=False, na=False)
                    mascara = coincide if mascara is None else (mascara | coincide)
            if mascara is not None:
                resultado = resultado[mascara]
    if orden_col and orden == "Descripción (A-Z)" and orden_col in resultado.columns:
        resultado = resultado.sort_values(orden_col, key=lambda s: s.astype(str).str.upper())
    else:
        resultado = resultado.sort_values(columna_fecha)
    return resultado


def barra_resultado(df_filtrado, df_total, titulo_excel, nombre_archivo, key):
    """Conteo a la izquierda y descarga de esa tabla (con el filtro aplicado) a la derecha."""
    col_a, col_b = st.columns([3, 1.15])
    with col_a:
        badge(f"Mostrando {len(df_filtrado):,} de {len(df_total):,} movimiento(s)")
    with col_b:
        st.download_button(
            "⬇  Descargar esta tabla en Excel",
            data=build_tabla_workbook(df_filtrado, meta, titulo_excel, nombre_hoja=key.capitalize()),
            file_name=f"{nombre_archivo} - {sufijo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, disabled=df_filtrado.empty, key=f"btn_dl_{key}",
        )


# ============================================================ RESULTADOS ==
if st.session_state.get("aviso"):
    st.success(st.session_state.pop("aviso"))

# =============================================== HOJA 1 · RESUMEN + KPIs ==
if vista == HOJAS[0]:
    if not os.path.exists(LOGO_PATH):
        st.info("El logo de ISTHO aún no está cargado. Guarda el archivo del logo como "
                "**logo_istho.png** dentro de la carpeta APP_CONCILIACION.")

    section("Saldos según el banco",
            "Saldo final = Saldo inicial + Total ingresos − Total salidas"
            + ("" if saldo_inicial_banco else " · escribe el saldo inicial en el panel izquierdo"))
    stat_cards([
        {"label": "Saldo inicial", "value": f"{saldo_inicial_banco:,.2f}", "icon": "🗓️", "accent": GRIS,
         "sub": "saldo final del mes anterior",
         "tip": "Saldo con el que cerró el mes anterior («SALDO ANTERIOR» en el extracto). "
                "Se escribe en el panel izquierdo y es el punto de partida del cálculo."},
        {"label": "Total ingresos", "value": f"{total_ingresos:,.2f}", "icon": "📈", "accent": VERDE,
         "sub": f"{int((df_banco['valor'] > 0).sum()):,} movimiento(s)",
         "tip": "Suma de todos los movimientos positivos del extracto, sin importar el concepto."},
        {"label": "Total salidas", "value": f"{total_salidas:,.2f}", "icon": "📉", "accent": ROJO,
         "sub": f"{int((df_banco['valor'] < 0).sum()):,} movimiento(s)",
         "tip": "Suma de todos los movimientos negativos del extracto: nómina, manifiestos, "
                "transferencias, proveedores, comisiones y cualquier otro cargo."},
        {"label": "Saldo final", "value": f"{saldo_final:,.2f}", "icon": "💠", "accent": VERDE_OSC,
         "sub": "calculado automáticamente",
         "tip": "Saldo inicial + Total ingresos − Total salidas. Debe coincidir con el "
                "«SALDO ACTUAL» que reporta el extracto bancario."},
    ])

    section("Saldos según el libro auxiliar",
            "Su propio saldo inicial (no necesariamente igual al del banco) con los ingresos "
            "y salidas de la contabilidad")
    stat_cards([
        {"label": "Saldo inicial", "value": f"{saldo_inicial_libro:,.2f}", "icon": "🗓️", "accent": GRIS,
         "sub": "arrastre de la contabilidad",
         "tip": "Saldo con el que el libro auxiliar traía la cuenta al empezar el mes. Se escribe "
                "en el panel izquierdo, por separado del saldo inicial del banco."},
        {"label": "Total ingresos", "value": f"{total_ingresos_libro:,.2f}", "icon": "📈", "accent": VERDE,
         "sub": f"{int((df_libro['valor'] > 0).sum()):,} registro(s)",
         "tip": "Suma de los débitos del libro auxiliar (entran al banco)."},
        {"label": "Total salidas", "value": f"{total_salidas_libro:,.2f}", "icon": "📉", "accent": ROJO,
         "sub": f"{int((df_libro['valor'] < 0).sum()):,} registro(s)",
         "tip": "Suma de los créditos del libro auxiliar (salen del banco)."},
        {"label": "Saldo final", "value": f"{saldo_final_libro:,.2f}", "icon": "💠", "accent": VERDE_OSC,
         "sub": "según la contabilidad",
         "tip": "Saldo inicial + Total ingresos − Total salidas, esta vez con los valores del "
                "libro auxiliar en vez de los del extracto."},
    ])

    section("Comparación banco vs. contabilidad",
            "Incluye el arrastre de saldos iniciales, no solo los movimientos de este mes")
    stat_cards([
        {"label": "Saldo final banco", "value": f"{saldo_final:,.2f}", "icon": "🏦", "accent": VERDE_OSC},
        {"label": "Saldo final contabilidad", "value": f"{saldo_final_libro:,.2f}", "icon": "📘",
         "accent": VERDE_OSC},
        {"label": "Diferencia", "value": f"{diferencia_saldos:,.2f}",
         "icon": ("✔️" if saldos_cuadran else "⚠️"), "accent": (VERDE if saldos_cuadran else ROJO),
         "sub": ("Los dos saldos finales coinciden" if saldos_cuadran
                 else "No coinciden — revisa los saldos iniciales o lo pendiente")},
    ])

    if est.get("cerrado"):
        cerrado_en_fmt = _fmt_dt(est.get("cerrado_en"))
        st.success(
            f"🔒 Esta conciliación quedó marcada como **terminada**"
            f"{f' el {cerrado_en_fmt}' if cerrado_en_fmt else ''} "
            f"— el saldo final del banco (${(est.get('saldo_final_banco') or 0):,.2f}) va a sugerirse como saldo "
            "inicial del próximo mes."
        )
        if st.button("↺ Reabrir (seguir editando este mes)"):
            st.session_state["estado"]["cerrado"] = False
            db.guardar_cambios(est["periodo"], st.session_state["estado"])
            st.rerun()
    else:
        st.info("Cuando termines de revisar todo lo pendiente de este mes, márcalo como terminado: "
                 "el saldo final queda guardado para sugerir el saldo inicial del mes siguiente, sin "
                 "que tengas que volver a escribirlo.")
        if st.button("🔒 Marcar esta conciliación como terminada", type="primary"):
            st.session_state["estado"]["cerrado"] = True
            st.session_state["estado"]["saldo_final_banco"] = saldo_final
            st.session_state["estado"]["saldo_final_libro"] = saldo_final_libro
            st.session_state["estado"]["cerrado_en"] = datetime.now().isoformat()
            db.guardar_cambios(est["periodo"], st.session_state["estado"])
            st.session_state["aviso"] = "Conciliación marcada como terminada."
            st.rerun()

    section("Resumen del cruce", f"{pct:.0f}% de los movimientos del banco quedaron conciliados"
                                 + (f" · {n_manuales} conciliación(es) manual(es)" if n_manuales else ""))
    stat_cards([
        {"label": "Movimientos banco", "value": f"{len(df_banco):,}", "icon": "🏦", "accent": GRIS,
         "tip": "Total de movimientos del extracto bancario que cargaste."},
        {"label": "Movimientos contabilidad", "value": f"{len(df_libro):,}", "icon": "📘", "accent": GRIS,
         "tip": "Total de registros del libro auxiliar contable que cargaste."},
        {"label": "Conciliados", "value": f"{n_conciliados:,}", "icon": "🔗", "accent": VERDE,
         "sub": f"en {n_cruces} conciliación(es)",
         "tip": "Movimientos que cruzaron con la contabilidad por valor exacto, ya sea uno a uno "
                "o sumando varios entre sí."},
        {"label": "Posibles (dif. valor)", "value": f"{n_posibles:,}", "icon": "🔍", "accent": NARANJA,
         "sub": "fecha y nombre coinciden",
         "tip": f"Coinciden en fecha y nombre, pero el valor difiere hasta ${margen_valor:,.0f}. "
                "No están conciliados: hay que revisarlos a mano."},
        {"label": "Solo en banco", "value": f"{len(df_solo_banco):,}", "icon": "📌", "accent": ROJO,
         "sub": "falta contabilizar",
         "tip": "Están en el extracto pero no tienen registro en la contabilidad. Posible partida "
                "pendiente de contabilizar."},
        {"label": "Solo en contabilidad", "value": f"{len(df_solo_libro):,}", "icon": "📌", "accent": ROJO,
         "sub": "falta en el banco",
         "tip": "Están en la contabilidad pero no aparecen en el extracto. Partida pendiente en el "
                "banco o error de digitación."},
    ])

    st.caption(
        f"**Control:** total banco {total_banco:,.2f} − total contabilidad {total_libro:,.2f} = "
        f"{diferencia:,.2f}, que "
        + ("✔ cuadra con las partidas sin cruzar." if cuadra
           else f"⚠ no cuadra (esperado {dif_esperada:,.2f}).")
    )
    if not cuadra:
        st.warning("La diferencia total no cuadra con las partidas sin cruzar. Revisa si hay valores "
                   "duplicados o registros parciales.")

# ================================================== HOJA 2 · CONCILIADOS ==
elif vista == HOJAS[1]:
    if df_conc.empty:
        st.info("No hay conciliaciones. Puedes crearlas a mano en la hoja **Conciliación manual**.")
    else:
        # Se usa un interruptor y no un expander: el expander se vuelve a cerrar en cada
        # recarga y la recarga la dispara la propia selección de filas, dejando el panel
        # invisible justo cuando el usuario acaba de elegir qué deshacer.
        if st.toggle("↩️  Deshacer conciliaciones (desconciliar)", key="ver_desconciliar"):
            st.caption("Selecciona una o varias conciliaciones y deshazlas. Los movimientos "
                       "vuelven a sus listas de pendientes **sin perder ni modificar ningún dato**.")
            df_res = resumen_cruces(df_banco, df_libro, cruces)
            ev = st.dataframe(
                df_res, use_container_width=True, hide_index=True, height=330, row_height=34,
                on_select="rerun", selection_mode="multi-row",
                key=f"sel_desconciliar_{st.session_state['gen']}",
                column_config={
                    "Descripción Banco": st.column_config.TextColumn(width=220),
                    "Descripción Contabilidad": st.column_config.TextColumn(width=420),
                    "Valor banco": st.column_config.NumberColumn(format="accounting"),
                    "Valor contabilidad": st.column_config.NumberColumn(format="accounting"),
                },
            )
            elegidos = [df_res.iloc[i]["ID"] for i in ev.selection.rows]
            if elegidos:
                n_mov = sum(len(c["banco_ids"]) + len(c["libro_ids"])
                            for c in cruces if c["id"] in elegidos)
                st.warning(f"Vas a deshacer **{len(elegidos)}** conciliación(es) "
                           f"({', '.join(elegidos[:6])}{'…' if len(elegidos) > 6 else ''}), "
                           f"lo que devolverá **{n_mov} movimiento(s)** a pendientes.")
                confirmar = st.checkbox("Confirmo que deseo deshacer estas conciliaciones",
                                         key=f"conf_desc_{st.session_state['gen']}")
                if st.button("↩️  Desconciliar", type="primary", disabled=not confirmar,
                             key="btn_desconciliar"):
                    st.session_state["estado"]["cruces"] = eliminar_cruces(cruces, elegidos)
                    db.guardar_cambios(st.session_state["estado"]["periodo"], st.session_state["estado"])
                    st.session_state["gen"] += 1
                    st.session_state["aviso"] = f"Se deshicieron {len(elegidos)} conciliación(es)."
                    st.rerun()
            else:
                st.caption("Marca las filas de la tabla para habilitar la acción.")

        filtrado = filtrar(df_conc, "conciliados", "Fecha Banco",
                            ["Descripción Banco", "Descripción Contabilidad", "Comprobante",
                             "Documento", "ID", "Origen", "Motivo"])
        filtrado = filtrado.sort_values(["ID", "Fecha Banco"])
        barra_resultado(filtrado, df_conc, "CONCILIACIÓN BANCARIA — CONCILIADOS",
                        "CONCILIACION BANCARIA - CONCILIADOS", "conciliados")
        tabla(filtrado, height=900)

# ===================================================== HOJA 3 · POSIBLES ==
elif vista == HOJAS[2]:
    if df_posibles.empty:
        st.success("No se encontraron posibles diferencias de valor.")
    else:
        st.caption(f"Coinciden en fecha y nombre, pero el valor difiere hasta ${margen_valor:,.0f}. "
                   "**No se dan por conciliados** — revísalos a mano.")
        filtrado = filtrar(df_posibles, "posibles", "Fecha Banco",
                            ["Descripción Banco", "Descripción Contabilidad", "Comprobante",
                             "Documento"]).sort_values("Fecha Banco")
        barra_resultado(filtrado, df_posibles, "CONCILIACIÓN BANCARIA — POSIBLES",
                        "CONCILIACION BANCARIA - POSIBLES", "posibles")
        tabla(filtrado, height=900)

# ==================================== HOJA 4 · PENDIENTES DEL EXTRACTO ==
elif vista == HOJAS[3]:
    if df_solo_banco.empty:
        st.success("Todos los movimientos del extracto bancario fueron cruzados.")
    else:
        st.caption("Movimientos del **extracto bancario** que no tienen registro en la "
                   "contabilidad — posible partida pendiente de contabilizar.")
        mostrar = df_solo_banco[["fecha", "valor", "tipo", "descripcion"]].rename(columns={
            "fecha": "Fecha", "valor": "Valor", "tipo": "Tipo", "descripcion": "Descripción"})
        filtrado = filtrar(mostrar, "solo_banco", "Fecha", ["Descripción"],
                            columna_tipo="Tipo", permitir_orden=True)
        barra_resultado(filtrado, mostrar, "CONCILIACIÓN BANCARIA — PENDIENTES DEL EXTRACTO",
                        "CONCILIACION BANCARIA - PENDIENTES EXTRACTO", "solo_banco")
        tabla(filtrado, height=900)
        st.caption(f"Suma total (sin filtrar): {total_solo_banco:,.2f}")

# =============================== HOJA 5 · PENDIENTES DEL LIBRO AUXILIAR ==
elif vista == HOJAS[4]:
    if df_solo_libro.empty:
        st.success("Todos los registros del libro auxiliar fueron cruzados.")
    else:
        st.caption("Registros del **libro auxiliar** que no aparecen en el extracto — partida "
                   "pendiente en el banco o error de digitación.")
        mostrar = df_solo_libro[["fecha", "valor", "tipo", "descripcion", "comprobante",
                                  "documento"]].rename(
            columns={"fecha": "Fecha", "valor": "Valor", "tipo": "Tipo", "descripcion": "Descripción",
                     "comprobante": "Comprobante", "documento": "Documento"})
        filtrado = filtrar(mostrar, "solo_libro", "Fecha", ["Descripción", "Comprobante", "Documento"],
                            columna_tipo="Tipo", permitir_orden=True)
        barra_resultado(filtrado, mostrar, "CONCILIACIÓN BANCARIA — PENDIENTES DEL LIBRO AUXILIAR",
                        "CONCILIACION BANCARIA - PENDIENTES LIBRO AUXILIAR", "solo_libro")
        tabla(filtrado, height=900)
        st.caption(f"Suma total (sin filtrar): {total_solo_libro:,.2f}")

# =========================================== HOJA 6 · CONCILIACIÓN MANUAL ==
else:
    section("Conciliación manual",
            "Selecciona movimientos de cada lado; el botón Cruzar se habilita cuando los totales coinciden")

    if st.session_state.get("aviso"):
        st.success(st.session_state.pop("aviso"))

    gen = st.session_state["gen"]


    def panel(df, titulo, key, columnas, columnas_texto):
        """Panel de pendientes con búsqueda, filtro por tipo y selección múltiple."""
        st.markdown(f"##### {titulo}")
        if df.empty:
            st.success("No quedan movimientos pendientes de este lado.")
            return [], 0.0
        filtrado = filtrar(df, f"man_{key}", "fecha", columnas_texto, columna_tipo="tipo",
                            permitir_orden=True, columna_orden_alt="descripcion")
        if filtrado.empty:
            st.info("Ningún movimiento coincide con el filtro.")
            return [], 0.0
        filtrado = filtrado.reset_index(drop=True)
        mostrar = filtrado[columnas].rename(columns={
            "fecha": "Fecha", "valor": "Valor", "tipo": "Tipo", "descripcion": "Descripción",
            "comprobante": "Comprobante", "documento": "Documento"})
        ev = st.dataframe(
            mostrar, use_container_width=True, hide_index=True, height=430, row_height=34,
            on_select="rerun", selection_mode="multi-row", key=f"sel_{key}_{gen}",
            column_config={
                "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY", width="small"),
                "Valor": st.column_config.NumberColumn(format="accounting", width="medium"),
                "Tipo": st.column_config.TextColumn(width="small"),
                "Descripción": st.column_config.TextColumn(width="large"),
            },
        )
        elegidos = [int(filtrado.iloc[i]["id"]) for i in ev.selection.rows]
        total = float(sum(filtrado.iloc[i]["valor"] for i in ev.selection.rows))
        st.caption(f"Mostrando {len(filtrado):,} pendiente(s). Marca las casillas para seleccionar.")
        return elegidos, total


    col_izq, col_der = st.columns(2)
    with col_izq:
        ids_banco, total_sel_banco = panel(
            df_solo_banco, "🏦 Pendientes en el banco", "banco",
            ["fecha", "valor", "tipo", "descripcion"], ["descripcion"])
    with col_der:
        ids_libro, total_sel_libro = panel(
            df_solo_libro, "📘 Pendientes en contabilidad", "libro",
            ["fecha", "valor", "tipo", "descripcion", "comprobante"],
            ["descripcion", "comprobante", "documento"])

    st.divider()
    dif_sel = round(total_sel_banco - total_sel_libro, 2)
    hay_seleccion = bool(ids_banco) and bool(ids_libro)
    coinciden = hay_seleccion and abs(dif_sel) < 0.005

    stat_cards([
        {"label": "Seleccionado en banco", "value": f"{total_sel_banco:,.2f}", "icon": "🏦",
         "accent": VERDE_OSC, "sub": f"{len(ids_banco)} registro(s)"},
        {"label": "Seleccionado en contabilidad", "value": f"{total_sel_libro:,.2f}", "icon": "📘",
         "accent": VERDE_OSC, "sub": f"{len(ids_libro)} registro(s)"},
        {"label": "Validación", "value": ("Coinciden" if coinciden else f"{abs(dif_sel):,.2f}"),
         "icon": ("✔️" if coinciden else "⛔"), "accent": (VERDE if coinciden else ROJO),
         "sub": ("Los valores coinciden" if coinciden
                 else ("Selecciona en ambos lados" if not hay_seleccion
                       else f"Diferencia de ${abs(dif_sel):,.2f}"))},
    ])

    if coinciden:
        st.success(f"**Los valores coinciden** — {len(ids_banco)} movimiento(s) del banco por "
                   f"${abs(total_sel_banco):,.2f} contra {len(ids_libro)} de contabilidad. "
                   "Puedes cruzarlos.")
    elif hay_seleccion:
        st.error(f"**Diferencia de ${abs(dif_sel):,.2f}** — los totales deben ser exactamente "
                 "iguales para poder cruzar.")
    else:
        st.info("Selecciona al menos un movimiento en cada panel para poder cruzarlos.")

    col_b1, _ = st.columns([1, 3])
    with col_b1:
        if st.button("⚖️  Cruzar", type="primary", disabled=not coinciden,
                     use_container_width=True, key="btn_cruzar"):
            nuevos, nuevo_id = crear_cruce_manual(cruces, ids_banco, ids_libro)
            st.session_state["estado"]["cruces"] = nuevos
            db.guardar_cambios(st.session_state["estado"]["periodo"], st.session_state["estado"])
            st.session_state["gen"] += 1
            st.session_state["aviso"] = (
                f"Conciliación **{nuevo_id}** creada: {len(ids_banco)} movimiento(s) del banco "
                f"con {len(ids_libro)} de contabilidad por ${abs(total_sel_banco):,.2f}.")
            st.rerun()

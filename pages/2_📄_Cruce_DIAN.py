"""Interfaz del Cruce DIAN. Motor: `conciliacion_dian.py`. Base de datos: `db_dian.py`.
Completamente aparte de Cruce Bancario: no comparte estado, datos, ni tabla de Supabase."""
import os
from datetime import datetime

import streamlit as st

import db_dian as db
from conciliacion_dian import (buscar_causacion, construir_vistas_dian, crear_cruce_manual_dian,
                                cruzar_en_lote, eliminar_cruces_dian, load_avansant, load_dian,
                                reconciliar_dian, vista_duplicados)
from config import CLAVE_ACCESO
from excel_export import EMPRESA, LOGO_PATH, NIT, build_tabla_workbook, periodo_desde_fechas
from ui import (AZUL, GRIS, NARANJA, ROJO, VERDE, acceso_permitido, badge, hero, icono,
                 inject_css, loader, panel_toggle, section, sidebar_brand, sidebar_step,
                 stat_cards, tabla)

st.set_page_config(page_title="Cruce DIAN", layout="wide",
                   page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "📄")

inject_css()


def _fmt_dt(iso):
    if not iso:
        return ""
    try:
        return datetime.fromisoformat(str(iso)).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return ""


# Mismo título genérico que en Cruce Bancario y en el menú: esta clave abre las dos apps
# (queda en session_state, compartido entre páginas), así que si alguien entra por acá
# primero, el título no debe sonar como si fuera solo para el cruce DIAN.
if not acceso_permitido(CLAVE_ACCESO, EMPRESA, "Sistema de Conciliaciones ISTHO",
                        "Elige el proceso que necesitas"):
    st.stop()

# ---------------------------------------------------------------- Sidebar --
panel_toggle()
sidebar_brand("ISTHO S.A.S.", "Cruce DIAN")
st.sidebar.page_link("app.py", label="Volver al menú", icon=":material/home:")
st.sidebar.divider()

sidebar_step(1, "Archivos")
st.sidebar.caption("El informe de documentos de la DIAN y el de comprobantes de Avansant, "
                   "tal como se descargan.")
archivo_dian = st.sidebar.file_uploader("Informe DIAN", type=["xlsx"])
archivo_avansant = st.sidebar.file_uploader("Informe Avansant (comprobantes)", type=["xlsx"])

ejecutar = st.sidebar.button("Cruzar", type="primary", use_container_width=True)

if "estado_dian" not in st.session_state:
    st.session_state["estado_dian"] = None
if "gen_dian" not in st.session_state:
    st.session_state["gen_dian"] = 0

if st.session_state["estado_dian"] is None and db.disponible():
    periodos = db.listar_periodos()
    if periodos:
        st.sidebar.divider()
        st.sidebar.caption(":material/cloud: Hay cruces guardados")
        opciones = {f"{p['periodo']} — actualizado {_fmt_dt(p['actualizado_en'])}": p["periodo"]
                    for p in periodos}
        etiqueta = st.sidebar.selectbox(
            "Retomar cruce guardado", list(opciones.keys()),
            index=None, placeholder="Elige un período…", key="periodo_dian_a_retomar")
        elegido = opciones.get(etiqueta)
        if elegido and st.sidebar.button(":material/restore: Retomar", use_container_width=True, key="btn_retomar_dian"):
            st.session_state["estado_dian"] = db.cargar_estado(elegido)
            st.session_state["gen_dian"] += 1
            st.rerun()

if ejecutar:
    if not archivo_dian or not archivo_avansant:
        st.sidebar.error("Debes cargar los dos archivos.")
    else:
        pantalla = st.empty()
        try:
            pantalla.markdown(loader("Leyendo el informe DIAN…",
                                     "Descartando acuses de recibo, armando el comprobante"),
                              unsafe_allow_html=True)
            df_dian = load_dian(archivo_dian)
            pantalla.markdown(loader("Leyendo el informe de Avansant…",
                                     f"{len(df_dian):,} documentos DIAN cargados"),
                              unsafe_allow_html=True)
            df_avansat = load_avansant(archivo_avansant)
            pantalla.markdown(loader("Cruzando documentos…",
                                     f"Cruzando {len(df_dian):,} documentos DIAN contra "
                                     f"{len(df_avansat):,} causaciones · por comprobante, NIT y valor"),
                              unsafe_allow_html=True)
            cruces, ambiguos, duplicados = reconciliar_dian(df_dian, df_avansat)
            periodo = periodo_desde_fechas(list(df_dian["fecha_emision"]))
            st.session_state["estado_dian"] = {
                "dian": df_dian, "avansat": df_avansat, "cruces": cruces,
                "ambiguos": ambiguos, "duplicados": duplicados, "periodo": periodo,
            }
            db.guardar_estado(periodo, st.session_state["estado_dian"])
            st.session_state["gen_dian"] += 1
        except Exception as e:
            st.sidebar.error(f"Error procesando los archivos: {e}")
            st.session_state["estado_dian"] = None
        finally:
            pantalla.empty()

est = st.session_state["estado_dian"]

if est is None:
    hero("Cruce DIAN",
         "Cruce automático entre los documentos de la DIAN y las causaciones de Avansant — "
         "muestra qué facturas todavía no se han digitado.",
         empresa=EMPRESA, periodo="Sin período cargado")
    st.info("Carga el informe DIAN y el de Avansant en el panel izquierdo y presiona **Cruzar**.")
    st.stop()

df_dian, df_avansat = est["dian"], est["avansat"]
cruces, ambiguos, duplicados = est["cruces"], est["ambiguos"], est["duplicados"]

df_cruz, df_amb, df_nc, df_sin, df_av_pend = construir_vistas_dian(df_dian, df_avansat, cruces, ambiguos)
df_dup = vista_duplicados(df_avansat, duplicados)

periodo_hero = periodo_desde_fechas(list(df_dian["fecha_emision"]))
periodo = est.get("periodo") or periodo_hero
meta = {"empresa": EMPRESA, "nit": NIT, "cuenta": "—", "periodo": periodo,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M"), "tolerancia": 0}
sufijo = periodo.replace(" ", "_").replace("/", "-")

# --------------------------------------------------------------- Helpers --
def filtrar(df, key, columna_fecha, columnas_texto, permitir_orden=False, columna_orden_alt=None):
    """Igual que la de Cruce Bancario (ver pages/1): filtros que se mantienen aunque
    cambies de pestaña, con key propia por hoja."""
    if df.empty:
        return df
    orden_col = (columna_orden_alt or (columnas_texto[0] if columnas_texto else None)) if permitir_orden else None
    with st.container(border=True, key=f"toolbar_dian_{key}"):
        cols = st.columns([2.6, 1.8, 1.3]) if orden_col else st.columns([3, 2])
        col1, col2 = cols[0], cols[1]
        texto = col1.text_input("Buscar", placeholder="NIT, nombre, comprobante...",
                                key=f"buscar_dian_{key}")
        resultado = df
        if columna_fecha and columna_fecha in df.columns:
            fechas = df[columna_fecha].dropna()
            if not fechas.empty and fechas.min() != fechas.max():
                rango = col2.date_input("Rango de fechas", value=(fechas.min(), fechas.max()),
                                        min_value=fechas.min(), max_value=fechas.max(),
                                        key=f"fechas_dian_{key}")
                if isinstance(rango, tuple) and len(rango) == 2:
                    resultado = resultado[(resultado[columna_fecha] >= rango[0])
                                           & (resultado[columna_fecha] <= rango[1])]
        orden = "Fecha"
        if orden_col:
            orden = cols[2].selectbox("Ordenar por", ["Fecha", "Emisor (A-Z)"], key=f"orden_dian_{key}")
        if texto:
            mascara = None
            for c in columnas_texto:
                if c in resultado.columns:
                    coincide = resultado[c].astype(str).str.contains(texto, case=False, na=False)
                    mascara = coincide if mascara is None else (mascara | coincide)
            if mascara is not None:
                resultado = resultado[mascara]
    if orden_col and orden == "Emisor (A-Z)" and orden_col in resultado.columns:
        resultado = resultado.sort_values(orden_col, key=lambda s: s.astype(str).str.upper())
    elif columna_fecha and columna_fecha in resultado.columns:
        resultado = resultado.sort_values(columna_fecha)
    return resultado


def barra_resultado(df_filtrado, df_total, titulo_excel, nombre_archivo, key):
    col_a, col_b = st.columns([3, 1.15])
    with col_a:
        badge(f"Mostrando {len(df_filtrado):,} de {len(df_total):,} registro(s)")
    with col_b:
        st.download_button(
            ":material/download: Descargar esta tabla en Excel",
            data=build_tabla_workbook(df_filtrado, meta, titulo_excel, nombre_hoja=key.capitalize()[:31]),
            file_name=f"{nombre_archivo} - {sufijo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, disabled=df_filtrado.empty, key=f"btn_dl_dian_{key}",
        )


def tabla_con_cruce(df, key, anchos=None, alto=720):
    """Tabla editable con una columna «Causación» al final y un botón que concilia de una
    sola vez todo lo que se haya escrito.

    Se usa en TODAS las hojas que listan documentos DIAN sin cruzar (pendientes, ambiguos,
    notas de crédito): el flujo real es revisar el documento, buscarlo en el correo,
    registrarlo en Avansant y anotar aquí mismo el número — sin tener que irse a otra hoja
    a buscarlo en un desplegable.

    `df` debe traer la columna `id` (el índice del documento en df_dian); se oculta al
    mostrar la tabla, pero es lo que permite saber a qué documento pertenece cada fila."""
    editable = df.copy()
    editable["Causación"] = ""

    config = {
        "id": None,  # índice interno del documento, no se muestra
        "Fecha Emisión": st.column_config.DateColumn(format="DD/MM/YYYY", width=95, disabled=True),
        "Comprobante DIAN": st.column_config.TextColumn(width=140, disabled=True),
        "NIT": st.column_config.TextColumn(width=105, disabled=True),
        "Emisor": st.column_config.TextColumn(width=280, disabled=True),
        "Total": st.column_config.NumberColumn(format="accounting", width=135, disabled=True),
        "Causación": st.column_config.TextColumn(
            "✏️ Causación", width=130,
            help="Escribe aquí el número de causación de Avansant y presiona el botón de abajo."),
    }
    for col, ancho in (anchos or {}).items():
        config[col] = st.column_config.TextColumn(width=ancho, disabled=True)
    # Todo lo que no se configuró arriba se bloquea igual: la única casilla editable de la
    # tabla debe ser «Causación».
    for col in editable.columns:
        config.setdefault(col, st.column_config.TextColumn(disabled=True))

    # `gen_dian` en la key: al conciliar, la tabla se redibuja desde cero y así un número ya
    # aplicado no queda escrito en la casilla de un documento distinto.
    # El alto se ajusta al contenido (con un tope): con 4 filas no tiene sentido dejar media
    # pantalla de renglones vacíos debajo.
    editado = st.data_editor(editable, hide_index=True, use_container_width=True,
                             height=min(alto, 78 + len(editable) * 36),
                             key=f"editor_{key}_{st.session_state['gen_dian']}",
                             column_config=config)

    escritos = [(int(r["id"]), str(r["Causación"]).strip())
                for _, r in editado.iterrows() if str(r["Causación"]).strip()]
    col_a, col_b = st.columns([1.5, 3])
    with col_a:
        if st.button(f":material/playlist_add_check: Conciliar los diligenciados ({len(escritos)})",
                     type="primary", disabled=not escritos, use_container_width=True,
                     key=f"btn_lote_{key}"):
            nuevos, aplicados, problemas = cruzar_en_lote(df_dian, df_avansat, cruces, escritos)
            st.session_state["estado_dian"]["cruces"] = nuevos
            # Un documento que estaba en «Requiere revisión» y ya se resolvió a mano sale de
            # esa lista: si no, seguiría apareciendo ahí como pendiente de decidir.
            resueltos = {i for i, _ in escritos}
            st.session_state["estado_dian"]["ambiguos"] = [
                a for a in ambiguos if a["dian_id"] not in resueltos]
            db.guardar_cambios(periodo, st.session_state["estado_dian"])
            st.session_state["gen_dian"] += 1
            st.session_state["aviso_dian"] = f"{aplicados} documento(s) conciliado(s)."
            st.session_state["problemas_dian"] = problemas
            st.rerun()
    with col_b:
        if escritos:
            st.caption(f"Vas a conciliar {len(escritos)} documento(s). Los que tengan el número "
                       "mal escrito se avisan y el resto sí se aplica.")


# --------------------------------------------------------------- Barra superior --
# Iconos Material (`:material/nombre:`), no emoji: Streamlit los reconoce en las etiquetas
# y los pinta con el mismo trazo sobrio en cualquier equipo, mientras que los emoji los
# dibuja el sistema operativo y cambian de estilo entre Windows, Mac y móvil.
HOJAS = [":material/donut_small: Resumen",
         ":material/link: Cruzados",
         ":material/price_change: Diferencia de valor",
         ":material/rule: Requiere revisión",
         ":material/note_stack: Notas de crédito",
         ":material/content_copy: Duplicados Avansant",
         ":material/pending_actions: Pendientes DIAN",
         ":material/edit_note: Cruce manual"]
vista = st.session_state.get("hoja_activa_dian") or HOJAS[0]

hero("Cruce DIAN", "DIAN (base) cruzado contra las causaciones de Avansant",
     empresa=EMPRESA, periodo=periodo_hero, compacto=(vista != HOJAS[0]))

with st.container(key="nav_hojas_dian"):
    seleccion = st.segmented_control("Hojas", HOJAS, default=HOJAS[0],
                                     label_visibility="collapsed", key="hoja_activa_dian")
vista = seleccion or vista

# ------------------------------------------------------------ Indicadores --
total_dian = float(df_dian["total"].sum())
total_avansat = float(df_avansat["valor"].sum())
# El universo a conciliar son las facturas, más cualquier nota de crédito que sí se haya
# cruzado a mano. Se cuenta con `df_nc` (que ya excluye las cruzadas) y no con el total de
# notas de crédito del archivo: si no, al cruzar una a mano el porcentaje pasaría de 100%.
n_dian_relevante = len(df_dian) - len(df_nc)
n_avansat = len(df_avansat)
n_cruzados = len(df_cruz)
n_pendientes = len(df_sin)
n_ambiguos = len(df_amb)
pct_conciliado = (n_cruzados / n_dian_relevante * 100) if n_dian_relevante else 0

# Cruces correctos (misma factura) pero con el valor distinto entre DIAN y Avansant: una
# retención registrada neta, o un error al digitar. Se separan porque son plata que hay que
# revisar y hasta ahora quedaban escondidos entre los cruces "Alta".
df_dif_valor = (df_cruz[df_cruz["Dif. valor"].abs() > 1].copy().sort_values(
                    "Dif. valor", key=lambda s: s.abs(), ascending=False)
                if not df_cruz.empty else df_cruz)
n_dif_valor = len(df_dif_valor)

if st.session_state.get("aviso_dian"):
    st.success(st.session_state.pop("aviso_dian"))

# Resultado del cruce en lote: lo que no se pudo aplicar (o se aplicó pero con el NIT o el
# valor descuadrado) se lista aquí para poder corregirlo, en vez de perderse en silencio.
if st.session_state.get("problemas_dian"):
    _problemas = st.session_state.pop("problemas_dian")
    with st.expander(f":material/error: {len(_problemas)} documento(s) necesitan tu atención",
                     expanded=True):
        for _p in _problemas:
            st.markdown(f"- **{_p['comprobante']}** (causación escrita: `{_p['numero']}`) — "
                        f"{_p['motivo']}")

if st.session_state.get("_dian_guardado_fallo"):
    st.warning(":material/warning: Este cruce no se pudo guardar en la nube (probablemente falta crear la "
              "tabla `conciliaciones_dian` en Supabase). El cruce en pantalla funciona "
              "normal, pero no vas a poder retomarlo si cierras la app.")

# ==================================================== HOJA 1 · RESUMEN ==
if vista == HOJAS[0]:
    section("Progreso de la revisión",
           f"{n_cruzados} de {n_dian_relevante} documentos ya tienen su causación — "
           "sube al 100% con la hoja **Cruce manual**")
    st.progress(min(pct_conciliado / 100, 1.0), text=f"{pct_conciliado:.1f}% conciliado")

    st.write("")
    section("Cifras generales")
    stat_cards([
        {"label": "Valor total DIAN", "value": f"${total_dian:,.0f}", "icon": icono("documento"),
         "accent": AZUL, "sub": f"{len(df_dian):,} documento(s)"},
        {"label": "Valor total Avansant", "value": f"${total_avansat:,.0f}", "icon": icono("libro"),
         "accent": AZUL, "sub": f"{n_avansat:,} causación(es)"},
        {"label": "Líneas DIAN", "value": f"{n_dian_relevante:,}", "icon": icono("capas"),
         "accent": GRIS, "sub": f"+ {len(df_nc)} nota(s) de crédito aparte"},
        {"label": "Líneas Avansant", "value": f"{n_avansat:,}", "icon": icono("pila"),
         "accent": GRIS, "sub": "causaciones (Tipo = CAUSACION)"},
    ])
    st.write("")
    stat_cards([
        {"label": "Cruzadas", "value": f"{n_cruzados:,}", "icon": icono("check"), "accent": VERDE,
         "sub": f"{n_cruzados/n_dian_relevante*100:.1f}%" if n_dian_relevante else "—"},
        {"label": "Requiere revisión", "value": f"{n_ambiguos:,}", "icon": icono("lupa"),
         "accent": NARANJA, "sub": "más de una causación posible"},
        {"label": "Pendientes (DIAN sin Avansant)", "value": f"{n_pendientes:,}",
         "icon": icono("pendiente"), "accent": ROJO,
         "sub": "es lo que buscabas: facturas sin digitar"},
        {"label": "Duplicados en Avansant", "value": f"{len(duplicados):,}", "icon": icono("alerta"),
         "accent": NARANJA, "sub": "misma referencia, distinto NIT"},
    ])

    if n_dif_valor:
        st.write("")
        st.warning(f"**{n_dif_valor} documento(s) cruzados tienen un valor distinto entre la DIAN "
                   f"y Avansant** (diferencia total ${df_dif_valor['Dif. valor'].sum():,.2f}). "
                   f"Revísalos en la hoja **Diferencia de valor** — suele ser una retención "
                   f"registrada neta o un error al digitar.")

    st.write("")
    niveles = {}
    for c in cruces:
        niveles[c["nivel"]] = niveles.get(c["nivel"], 0) + 1
    section("Cómo cruzaron los que cruzaron")
    stat_cards([
        {"label": "Alta · texto exacto", "value": f"{niveles.get('Alta', 0):,}",
         "icon": icono("check"), "accent": VERDE},
        {"label": "Media · NIT + valor", "value": f"{niveles.get('Media', 0):,}",
         "icon": icono("saldo"), "accent": AZUL},
        {"label": "Baja · NIT + folio", "value": f"{niveles.get('Baja', 0):,}",
         "icon": icono("documento"), "accent": GRIS},
        {"label": "Manual", "value": f"{niveles.get('Manual', 0):,}",
         "icon": icono("editar"), "accent": AZUL},
    ])

# ==================================================== HOJA 2 · CRUZADOS ==
elif vista == HOJAS[1]:
    if df_cruz.empty:
        st.info("Todavía no hay documentos cruzados.")
    else:
        if st.toggle(":material/undo: Deshacer cruces", key="ver_deshacer_dian"):
            st.caption("Selecciona uno o varios cruces y deshazlos. Los documentos vuelven "
                      "a pendientes sin perder ni modificar ningún dato.")
            ev = st.dataframe(
                df_cruz, use_container_width=True, hide_index=True, height=300, row_height=34,
                on_select="rerun", selection_mode="multi-row",
                key=f"sel_deshacer_dian_{st.session_state['gen_dian']}",
                column_config={
                    "Emisor": st.column_config.TextColumn(width=220),
                    "Total": st.column_config.NumberColumn(format="accounting"),
                },
            )
            elegidos = [df_cruz.iloc[i]["ID"] for i in ev.selection.rows]
            if elegidos:
                if st.button(f":material/undo: Deshacer {len(elegidos)} cruce(s)", type="primary",
                            key="btn_deshacer_dian"):
                    st.session_state["estado_dian"]["cruces"] = eliminar_cruces_dian(cruces, elegidos)
                    db.guardar_cambios(periodo, st.session_state["estado_dian"])
                    st.session_state["gen_dian"] += 1
                    st.session_state["aviso_dian"] = f"Se deshicieron {len(elegidos)} cruce(s)."
                    st.rerun()
            else:
                st.caption("Marca las filas de la tabla para habilitar la acción.")
        filtrado = filtrar(df_cruz, "cruzados", "Fecha Emisión",
                           ["Comprobante DIAN", "NIT", "Emisor", "Referencia Avansant",
                            "Tercero Avansant", "Causación"], permitir_orden=True,
                           columna_orden_alt="Emisor")
        barra_resultado(filtrado, df_cruz, "CRUCE DIAN — CRUZADOS", "CRUCE DIAN - CRUZADOS", "cruzados")
        tabla(filtrado, height=850)

# ============================================= HOJA 3 · DIFERENCIA DE VALOR ==
elif vista == HOJAS[2]:
    if df_dif_valor.empty:
        st.success("Todos los documentos cruzados tienen el mismo valor en la DIAN y en Avansant.")
    else:
        st.caption("Estos documentos **sí cruzaron** (misma factura, mismo NIT), pero el valor "
                  "que reporta la DIAN y el que quedó causado en Avansant **no son iguales**. "
                  "Suele ser una retención registrada neta o un error al digitar — vale la pena "
                  "revisarlos uno por uno.")
        filtrado = filtrar(df_dif_valor, "dif_valor", "Fecha Emisión",
                           ["Comprobante DIAN", "NIT", "Emisor", "Causación"])
        barra_resultado(filtrado, df_dif_valor, "CRUCE DIAN — DIFERENCIA DE VALOR",
                        "CRUCE DIAN - DIFERENCIA VALOR", "dif_valor")
        tabla(filtrado, height=800)
        st.caption(f"Diferencia total (sin filtrar): "
                   f"${df_dif_valor['Dif. valor'].sum():,.2f}")

# =============================================== HOJA 4 · REQUIERE REVISIÓN ==
elif vista == HOJAS[3]:
    if df_amb.empty:
        st.success("No hay documentos ambiguos — todo lo que tenía más de un candidato ya se resolvió.")
    else:
        st.caption("Estos documentos tienen **más de una** causación posible en Avansant — la app "
                  "no elige a ciegas. Mira la columna **Candidatos**, decide cuál corresponde y "
                  "escribe su número en la última columna.")
        filtrado = filtrar(df_amb, "ambiguos", "Fecha Emisión",
                           ["Comprobante DIAN", "NIT", "Emisor", "Candidatos"])
        barra_resultado(filtrado.drop(columns=["id"]), df_amb.drop(columns=["id"]),
                        "CRUCE DIAN — REQUIERE REVISIÓN",
                        "CRUCE DIAN - REQUIERE REVISION", "ambiguos")
        tabla_con_cruce(filtrado, "ambiguos",
                        anchos={"Emisor": 190, "Motivo": 175, "Candidatos": 300}, alto=560)

# =============================================== HOJA 4 · NOTAS DE CRÉDITO ==
elif vista == HOJAS[4]:
    st.caption("Notas de crédito electrónicas de la DIAN. No cruzan solas porque en Avansant no "
              "generan una causación propia — por eso no cuentan como \"pendientes\". Si alguna "
              "sí quedó registrada, escribe su número de causación en la última columna.")
    if df_nc.empty:
        st.info("No hay notas de crédito en este período.")
    else:
        mostrar = df_nc[["id", "fecha_emision", "comprobante", "nit_emisor", "nombre_emisor",
                         "total"]].rename(
            columns={"fecha_emision": "Fecha Emisión", "comprobante": "Comprobante DIAN",
                     "nit_emisor": "NIT", "nombre_emisor": "Emisor", "total": "Total"})
        filtrado = filtrar(mostrar, "notas_credito", "Fecha Emisión", ["Comprobante DIAN", "NIT", "Emisor"])
        barra_resultado(filtrado.drop(columns=["id"]), mostrar.drop(columns=["id"]),
                        "CRUCE DIAN — NOTAS DE CRÉDITO",
                        "CRUCE DIAN - NOTAS CREDITO", "notas_credito")
        tabla_con_cruce(filtrado, "notas_credito", alto=420)

# ============================================= HOJA 5 · DUPLICADOS AVANSANT ==
elif vista == HOJAS[5]:
    st.caption("La misma referencia aparece escrita en 2 o más causaciones de Avansant, con "
              "NIT distinto en cada una — no hay forma de saber a cuál factura corresponde "
              "cada una solo con el texto, así que ninguna se cruzó automáticamente por esta vía.")
    if df_dup.empty:
        st.success("No se encontraron referencias duplicadas en Avansant.")
    else:
        filtrado = filtrar(df_dup, "duplicados", None, ["Referencia", "NIT", "Tercero", "Detalle"])
        barra_resultado(filtrado, df_dup, "CRUCE DIAN — DUPLICADOS EN AVANSANT",
                        "CRUCE DIAN - DUPLICADOS AVANSANT", "duplicados")
        tabla(filtrado, height=500)

# =================================================== HOJA 6 · PENDIENTES DIAN ==
elif vista == HOJAS[6]:
    if df_sin.empty:
        st.success("No quedan documentos de DIAN sin causación en Avansant.")
    else:
        st.caption("Documentos de la **DIAN** que no tienen ninguna causación en Avansant — "
                  "esta es la lista de trabajo: revisa el correo, regístralos en Avansant y "
                  "escribe aquí mismo el número de causación en la última columna. "
                  "Al final presiona **Conciliar los diligenciados** y se cruzan todos juntos.")
        mostrar = df_sin[["id", "fecha_emision", "comprobante", "nit_emisor", "nombre_emisor",
                          "total"]].rename(
            columns={"fecha_emision": "Fecha Emisión", "comprobante": "Comprobante DIAN",
                     "nit_emisor": "NIT", "nombre_emisor": "Emisor", "total": "Total"})
        filtrado = filtrar(mostrar, "pendientes_dian", "Fecha Emisión",
                           ["Comprobante DIAN", "NIT", "Emisor"], permitir_orden=True,
                           columna_orden_alt="Emisor")
        barra_resultado(filtrado.drop(columns=["id"]), mostrar,
                        "CRUCE DIAN — PENDIENTES", "CRUCE DIAN - PENDIENTES", "pendientes_dian")
        tabla_con_cruce(filtrado, "pendientes", alto=760)
        st.caption(f"Suma total (sin filtrar): ${df_sin['total'].sum():,.2f}")

# ======================================================= HOJA 7 · CRUCE MANUAL ==
else:
    section("Cruce manual",
           "Escribe el número de causación que encontraste en Avansant — la app valida que "
           "el NIT y el valor tengan sentido antes de confirmar")

    pendientes_y_ambiguos = list(df_sin["id"]) + [a["dian_id"] for a in ambiguos]
    if not pendientes_y_ambiguos:
        st.success("No queda ningún documento pendiente ni ambiguo por cruzar a mano.")
    else:
        opciones_doc = {
            f"{df_dian.at[did, 'comprobante']} — {df_dian.at[did, 'nombre_emisor'][:40]} "
            f"— ${df_dian.at[did, 'total']:,.0f} ({df_dian.at[did, 'fecha_emision']})": did
            for did in pendientes_y_ambiguos
        }
        etiqueta_doc = st.selectbox("Documento DIAN a cruzar", list(opciones_doc.keys()),
                                    index=None, placeholder="Busca por comprobante, proveedor o valor…",
                                    key="doc_dian_manual")
        did = opciones_doc.get(etiqueta_doc)

        if did is not None:
            d = df_dian.loc[did]
            amb = next((a for a in ambiguos if a["dian_id"] == did), None)
            if amb:
                st.info("Este documento tiene varias causaciones candidatas — mira la columna "
                       "**Candidatos** en la hoja *Requiere revisión* para ver cuál corresponde, "
                       "y escribe su número aquí.")

            col_a, col_b = st.columns([1, 1.4])
            with col_a:
                st.markdown(f"**Comprobante:** {d['comprobante']}")
                st.markdown(f"**Emisor:** {d['nombre_emisor']}  \n**NIT:** {d['nit_emisor']}")
                st.markdown(f"**Total:** ${d['total']:,.2f}  \n**Fecha emisión:** {d['fecha_emision']}")

            with col_b:
                numero = st.text_input("Número de causación en Avansant", key=f"causacion_{did}")
                fila_causacion = buscar_causacion(df_avansat, numero) if numero else None

                if numero and fila_causacion is None:
                    st.error(f"No se encontró ninguna causación con el número «{numero}» en Avansant.")
                elif fila_causacion is not None:
                    nit_ok = (not fila_causacion["nit"]) or fila_causacion["nit"] == d["nit_emisor"]
                    valor_ok = abs(float(fila_causacion["valor"]) - float(d["total"])) < 1
                    if nit_ok and valor_ok:
                        st.success("El NIT y el valor coinciden con la factura de DIAN.")
                    else:
                        motivos = []
                        if not nit_ok:
                            motivos.append(f"el NIT no coincide (Avansant: {fila_causacion['nit']})")
                        if not valor_ok:
                            motivos.append(f"el valor no coincide (Avansant: ${fila_causacion['valor']:,.2f})")
                        st.warning("Revisa antes de confirmar: " + " y ".join(motivos) + ".")
                    st.markdown(f"**Tercero en Avansant:** {fila_causacion['tercero']}  \n"
                               f"**Detalle:** {fila_causacion['detalle']}  \n"
                               f"**Fecha contable:** {fila_causacion['fecha_contable']}")

                    if st.button(":material/check_circle: Confirmar cruce", type="primary", key=f"confirmar_{did}"):
                        nuevos_cruces, nuevo_id = crear_cruce_manual_dian(
                            cruces, did, int(fila_causacion["id"]), df_avansat)
                        st.session_state["estado_dian"]["cruces"] = nuevos_cruces
                        if amb:
                            st.session_state["estado_dian"]["ambiguos"] = [
                                a for a in ambiguos if a["dian_id"] != did]
                        db.guardar_cambios(periodo, st.session_state["estado_dian"])
                        st.session_state["gen_dian"] += 1
                        st.session_state["aviso_dian"] = (
                            f"Cruce **{nuevo_id}** creado: {d['comprobante']} con la causación "
                            f"{fila_causacion['causacion']}.")
                        st.rerun()

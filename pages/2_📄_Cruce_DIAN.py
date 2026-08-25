"""Interfaz del Cruce DIAN. Motor: `conciliacion_dian.py`. Base de datos: `db_dian.py`.
Completamente aparte de Cruce Bancario: no comparte estado, datos, ni tabla de Supabase."""
import os
from datetime import datetime

import streamlit as st

import db_dian as db
from conciliacion_dian import (buscar_causacion, construir_vistas_dian, crear_cruce_manual_dian,
                                eliminar_cruces_dian, load_avansant, load_dian, reconciliar_dian,
                                vista_duplicados)
from config import CLAVE_ACCESO
from excel_export import EMPRESA, LOGO_PATH, NIT, build_tabla_workbook, periodo_desde_fechas
from ui import (AZUL, GRIS, NARANJA, ROJO, VERDE, acceso_permitido, badge, hero, inject_css,
                 loader, section, sidebar_brand, sidebar_step, stat_cards, tabla)

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


if not acceso_permitido(CLAVE_ACCESO, EMPRESA, "Cruce DIAN",
                        "Documentos de la DIAN contra las causaciones de Avansant"):
    st.stop()

# ---------------------------------------------------------------- Sidebar --
sidebar_brand("ISTHO S.A.S.", "Cruce DIAN")
st.sidebar.page_link("app.py", label="← Volver al menú", icon="🏠")
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
        st.sidebar.caption("☁️ Hay cruces guardados")
        opciones = {f"{p['periodo']} — actualizado {_fmt_dt(p['actualizado_en'])}": p["periodo"]
                    for p in periodos}
        etiqueta = st.sidebar.selectbox(
            "Retomar cruce guardado", list(opciones.keys()),
            index=None, placeholder="Elige un período…", key="periodo_dian_a_retomar")
        elegido = opciones.get(etiqueta)
        if elegido and st.sidebar.button("↻ Retomar", use_container_width=True, key="btn_retomar_dian"):
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
            "⬇  Descargar esta tabla en Excel",
            data=build_tabla_workbook(df_filtrado, meta, titulo_excel, nombre_hoja=key.capitalize()[:31]),
            file_name=f"{nombre_archivo} - {sufijo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, disabled=df_filtrado.empty, key=f"btn_dl_dian_{key}",
        )


# --------------------------------------------------------------- Barra superior --
HOJAS = ["📊 Resumen", "🔗 Cruzados", "🟠 Requiere revisión", "⬜ Notas de crédito",
         "⚠️ Duplicados Avansant", "🔴 Pendientes DIAN", "✍️ Cruce manual"]
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
n_dian_relevante = len(df_dian) - int(df_dian["es_nota_credito"].sum())
n_avansat = len(df_avansat)
n_cruzados = len(df_cruz)
n_pendientes = len(df_sin)
n_ambiguos = len(df_amb)
pct_conciliado = (n_cruzados / n_dian_relevante * 100) if n_dian_relevante else 0

if st.session_state.get("aviso_dian"):
    st.success(st.session_state.pop("aviso_dian"))

if st.session_state.get("_dian_guardado_fallo"):
    st.warning("⚠️ Este cruce no se pudo guardar en la nube (probablemente falta crear la "
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
        {"label": "Valor total DIAN", "value": f"${total_dian:,.0f}", "icon": "📄", "accent": AZUL,
         "sub": f"{len(df_dian):,} documento(s)"},
        {"label": "Valor total Avansant", "value": f"${total_avansat:,.0f}", "icon": "🧾", "accent": AZUL,
         "sub": f"{n_avansat:,} causación(es)"},
        {"label": "Líneas DIAN", "value": f"{n_dian_relevante:,}", "icon": "📥", "accent": GRIS,
         "sub": f"+ {int(df_dian['es_nota_credito'].sum())} nota(s) de crédito aparte"},
        {"label": "Líneas Avansant", "value": f"{n_avansat:,}", "icon": "📥", "accent": GRIS,
         "sub": "causaciones (Tipo = CAUSACION)"},
    ])
    st.write("")
    stat_cards([
        {"label": "Cruzadas", "value": f"{n_cruzados:,}", "icon": "✔️", "accent": VERDE,
         "sub": f"{n_cruzados/n_dian_relevante*100:.1f}%" if n_dian_relevante else "—"},
        {"label": "Requiere revisión", "value": f"{n_ambiguos:,}", "icon": "🟠", "accent": NARANJA,
         "sub": "más de una causación posible"},
        {"label": "Pendientes (DIAN sin Avansant)", "value": f"{n_pendientes:,}", "icon": "🔴",
         "accent": ROJO, "sub": "es lo que buscabas: facturas sin digitar"},
        {"label": "Duplicados en Avansant", "value": f"{len(duplicados):,}", "icon": "⚠️",
         "accent": NARANJA, "sub": "misma referencia, distinto NIT"},
    ])

    st.write("")
    niveles = {}
    for c in cruces:
        niveles[c["nivel"]] = niveles.get(c["nivel"], 0) + 1
    section("Cómo cruzaron los que cruzaron")
    stat_cards([
        {"label": "🟢 Alta (texto exacto)", "value": f"{niveles.get('Alta', 0):,}", "accent": VERDE},
        {"label": "🟡 Media (NIT + valor)", "value": f"{niveles.get('Media', 0):,}", "accent": AZUL},
        {"label": "🔵 Baja (NIT + folio)", "value": f"{niveles.get('Baja', 0):,}", "accent": GRIS},
        {"label": "✍️ Manual", "value": f"{niveles.get('Manual', 0):,}", "accent": AZUL},
    ])

# ==================================================== HOJA 2 · CRUZADOS ==
elif vista == HOJAS[1]:
    if df_cruz.empty:
        st.info("Todavía no hay documentos cruzados.")
    else:
        if st.toggle("↩️  Deshacer cruces", key="ver_deshacer_dian"):
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
                if st.button(f"↩️  Deshacer {len(elegidos)} cruce(s)", type="primary",
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

# =============================================== HOJA 3 · REQUIERE REVISIÓN ==
elif vista == HOJAS[2]:
    if df_amb.empty:
        st.success("No hay documentos ambiguos — todo lo que tenía más de un candidato ya se resolvió.")
    else:
        st.caption("Estos documentos tienen **más de una** causación posible en Avansant — la app "
                  "no elige a ciegas. Revísalos en **Cruce manual** escribiendo el número correcto.")
        filtrado = filtrar(df_amb, "ambiguos", "Fecha Emisión",
                           ["Comprobante DIAN", "NIT", "Emisor", "Candidatos"])
        barra_resultado(filtrado, df_amb, "CRUCE DIAN — REQUIERE REVISIÓN",
                        "CRUCE DIAN - REQUIERE REVISION", "ambiguos")
        tabla(filtrado, height=850)

# =============================================== HOJA 4 · NOTAS DE CRÉDITO ==
elif vista == HOJAS[3]:
    st.caption("Notas de crédito electrónicas de la DIAN. Hoy no tienen nada con qué cruzar en "
              "Avansant (no generan causación ni reversión registrada), así que se muestran "
              "aparte — no cuentan como \"pendientes\".")
    if df_nc.empty:
        st.info("No hay notas de crédito en este período.")
    else:
        mostrar = df_nc[["fecha_emision", "comprobante", "nit_emisor", "nombre_emisor", "total"]].rename(
            columns={"fecha_emision": "Fecha Emisión", "comprobante": "Comprobante DIAN",
                     "nit_emisor": "NIT", "nombre_emisor": "Emisor", "total": "Total"})
        filtrado = filtrar(mostrar, "notas_credito", "Fecha Emisión", ["Comprobante DIAN", "NIT", "Emisor"])
        barra_resultado(filtrado, mostrar, "CRUCE DIAN — NOTAS DE CRÉDITO",
                        "CRUCE DIAN - NOTAS CREDITO", "notas_credito")
        tabla(filtrado, height=600)

# ============================================= HOJA 5 · DUPLICADOS AVANSANT ==
elif vista == HOJAS[4]:
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
elif vista == HOJAS[5]:
    if df_sin.empty:
        st.success("No quedan documentos de DIAN sin causación en Avansant.")
    else:
        st.caption("Documentos de la **DIAN** que no tienen ninguna causación en Avansant — "
                  "esto es lo que contabilidad todavía tiene que digitar.")
        mostrar = df_sin[["fecha_emision", "comprobante", "nit_emisor", "nombre_emisor", "total"]].rename(
            columns={"fecha_emision": "Fecha Emisión", "comprobante": "Comprobante DIAN",
                     "nit_emisor": "NIT", "nombre_emisor": "Emisor", "total": "Total"})
        filtrado = filtrar(mostrar, "pendientes_dian", "Fecha Emisión",
                           ["Comprobante DIAN", "NIT", "Emisor"], permitir_orden=True,
                           columna_orden_alt="Emisor")
        barra_resultado(filtrado, mostrar, "CRUCE DIAN — PENDIENTES", "CRUCE DIAN - PENDIENTES",
                        "pendientes_dian")
        tabla(filtrado, height=850)
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

                    if st.button("✔️  Confirmar cruce", type="primary", key=f"confirmar_{did}"):
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

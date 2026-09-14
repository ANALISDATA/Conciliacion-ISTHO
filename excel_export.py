"""Generación de los Excel finales: CONCILIACION BANCARIA APROBADA y PENDIENTE."""
import io
import os

import pandas as pd
import xlsxwriter

# Logo del membrete: PNG con este nombre exacto en esta misma carpeta.
LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo_istho.png")

# El nombre de la empresa, el NIT y la cuenta ya no están escritos aquí: se leen de la
# configuración privada (ver config.py), para que el código pueda publicarse sin exponerlos.
from config import CUENTA_DEFECTO, EMPRESA, NIT  # noqa: E402  (re-exportados por compatibilidad)

MESES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
         7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

# Misma paleta que la app (ver AZUL/AZUL_OSC/NARANJA en ui.py) — antes el membrete y los
# encabezados de tabla usaban un verde/oscuro que no correspondía con nada de la interfaz.
COLOR_OSCURO = "#16307A"   # = AZUL_OSC en ui.py
COLOR_VERDE = "#1D4ED8"    # = AZUL en ui.py (el nombre queda por no tocar el resto del archivo)
COLOR_NARANJA = "#E08700"  # = NARANJA en ui.py
COLOR_GRIS_CLARO = "#F2F2F2"


def periodo_desde_fechas(fechas):
    fechas = [f for f in fechas if f is not None]
    if not fechas:
        return "N/A"
    fmin, fmax = min(fechas), max(fechas)
    if (fmin.year, fmin.month) == (fmax.year, fmax.month):
        return f"{MESES[fmin.month]} {fmin.year}"
    return f"{MESES[fmin.month]} {fmin.year} - {MESES[fmax.month]} {fmax.year}"


def _formats(workbook):
    # El logo ya no comparte fila con el texto, así que basta una sangría pequeña.
    base = {"valign": "vcenter", "indent": 1}
    return {
        # Franja superior en blanco donde va el logo (membrete), separada de la banda oscura.
        "banda": workbook.add_format({"bg_color": "white"}),
        "titulo": workbook.add_format({**base, "bold": True, "font_size": 16, "font_color": "white",
                                        "bg_color": COLOR_OSCURO}),
        "subtitulo": workbook.add_format({**base, "bold": True, "font_size": 11, "font_color": "white",
                                           "bg_color": COLOR_OSCURO}),
        "meta": workbook.add_format({**base, "font_size": 10, "font_color": "white", "bg_color": COLOR_OSCURO}),
        "header": workbook.add_format({"bold": True, "font_color": "white", "bg_color": COLOR_VERDE,
                                        "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "celda": workbook.add_format({"border": 1, "valign": "vcenter"}),
        "celda_gris": workbook.add_format({"border": 1, "valign": "vcenter", "bg_color": COLOR_GRIS_CLARO}),
        "moneda": workbook.add_format({"border": 1, "num_format": "#,##0.00", "valign": "vcenter"}),
        "moneda_gris": workbook.add_format({"border": 1, "num_format": "#,##0.00", "valign": "vcenter",
                                             "bg_color": COLOR_GRIS_CLARO}),
        "fecha": workbook.add_format({"border": 1, "num_format": "dd/mm/yyyy", "valign": "vcenter"}),
        "fecha_gris": workbook.add_format({"border": 1, "num_format": "dd/mm/yyyy", "valign": "vcenter",
                                            "bg_color": COLOR_GRIS_CLARO}),
        "total_label": workbook.add_format({"bold": True, "border": 1, "bg_color": COLOR_NARANJA,
                                             "font_color": "white"}),
        "total_valor": workbook.add_format({"bold": True, "border": 1, "bg_color": COLOR_NARANJA,
                                             "font_color": "white", "num_format": "#,##0.00"}),
    }


def _encabezado(workbook, worksheet, titulo, meta, n_columnas):
    """Membrete: el logo ocupa una fila propia arriba del texto, de modo que nunca se
    superpone con el título ni con los datos (antes se insertaba sobre la fila del título)."""
    fmts = _formats(workbook)
    ultima = max(n_columnas - 1, 1)

    # Fila 0: banda oscura reservada solo para el logo.
    worksheet.set_row(0, 40)
    worksheet.merge_range(0, 0, 0, ultima, "", fmts["banda"])
    if os.path.exists(LOGO_PATH):
        worksheet.insert_image(0, 0, LOGO_PATH,
                                {"x_offset": 10, "y_offset": 6, "x_scale": 0.13, "y_scale": 0.13})

    worksheet.set_row(1, 26)
    worksheet.merge_range(1, 0, 1, ultima, titulo, fmts["titulo"])
    worksheet.merge_range(2, 0, 2, ultima,
                           f"{meta['empresa']}     |     NIT: {meta['nit']}", fmts["subtitulo"])
    worksheet.merge_range(3, 0, 3, ultima,
                           f"Cuenta: {meta['cuenta']}     |     Periodo: {meta['periodo']}", fmts["meta"])
    worksheet.merge_range(4, 0, 4, ultima,
                           f"Generado: {meta['generado']}     |     Tolerancia de fecha usada: "
                           f"{meta['tolerancia']} día(s)", fmts["meta"])
    return fmts, 6


def _escribir_tabla(worksheet, fmts, start_row, columnas, df, congelar_y_filtrar=True):
    """columnas: lista de tuplas (titulo, campo, tipo) tipo en {'texto','moneda','fecha','entero'}.

    `congelar_y_filtrar=False` en el informe completo (`build_resumen_completo_workbook`):
    ahí se escriben 3 tablas en la MISMA hoja, y tanto `autofilter` como `freeze_panes` son de
    hoja completa — llamarlos una vez por sección hace que solo quede el de la ÚLTIMA, y como
    esa suele caer casi al final de la hoja, el panel "congelado" terminaba tapando casi todo
    el informe (se veía la primera pantalla y no dejaba bajar más). El informe completo congela
    una sola vez, debajo del membrete, ver `build_resumen_completo_workbook`."""
    for c, (titulo, _, _) in enumerate(columnas):
        worksheet.write(start_row, c, titulo, fmts["header"])
    worksheet.set_row(start_row, 24)

    r = start_row + 1
    for i, (_, row) in enumerate(df.iterrows()):
        gris = (i % 2 == 1)
        for c, (_, campo, tipo) in enumerate(columnas):
            valor = row.get(campo) if hasattr(row, "get") else row[campo]
            if tipo == "moneda":
                worksheet.write_number(r, c, float(valor) if pd.notna(valor) else 0.0,
                                        fmts["moneda_gris" if gris else "moneda"])
            elif tipo == "fecha":
                if pd.notna(valor):
                    worksheet.write_datetime(r, c, valor, fmts["fecha_gris" if gris else "fecha"])
                else:
                    worksheet.write(r, c, "", fmts["celda_gris" if gris else "celda"])
            elif tipo == "entero":
                worksheet.write_number(r, c, int(valor) if pd.notna(valor) else 0,
                                        fmts["celda_gris" if gris else "celda"])
            else:
                # write_string() a propósito, no write(): las descripciones vienen del
                # extracto y del libro auxiliar (texto externo, no escrito por la app) y
                # write() genérico interpreta como fórmula cualquier valor que empiece con
                # "=" — un beneficiario o detalle contable que por casualidad empezara así
                # se ejecutaría como fórmula al abrir el Excel en vez de mostrarse como texto.
                worksheet.write_string(r, c, "" if pd.isna(valor) else str(valor),
                                        fmts["celda_gris" if gris else "celda"])
        r += 1

    if congelar_y_filtrar:
        if r > start_row + 1:
            worksheet.autofilter(start_row, 0, r - 1, len(columnas) - 1)
        worksheet.freeze_panes(start_row + 1, 0)
    return r


def _fila_total(worksheet, fmts, row, n_columnas, etiqueta, columnas_valor):
    """columnas_valor: dict {indice_columna: valor_total}."""
    worksheet.write(row, 0, etiqueta, fmts["total_label"])
    for c in range(1, n_columnas):
        if c in columnas_valor:
            worksheet.write_number(row, c, columnas_valor[c], fmts["total_valor"])
        else:
            worksheet.write(row, c, "", fmts["total_label"])


def _ajustar_anchos(worksheet, anchos):
    for c, ancho in enumerate(anchos):
        worksheet.set_column(c, c, ancho)


# --------------------------------------------------------------------------------
# Exportación genérica: convierte cualquier tabla de la app en un Excel con membrete
# --------------------------------------------------------------------------------
_COLS_FECHA = {"fecha", "fecha banco", "fecha contabilidad",
               # Cruce DIAN (ver conciliacion_dian.py) — agregado aparte para no tocar el
               # formato de ninguna columna que ya usa BANCARIO.
               "fecha emisión", "fecha contable"}
_COLS_MONEDA = {"valor", "valor banco", "valor contabilidad", "diferencia", "total", "valor avansant", "dif. valor"}
_COLS_ENTERO = {"dif. días"}

_ANCHOS_EXCEL = {
    "grupo": 8, "dif. días": 9, "tipo": 10,
    "id": 9, "origen": 12, "motivo": 40, "conciliado el": 18,
    "fecha": 13, "fecha banco": 13, "fecha contabilidad": 16,
    "valor": 16, "valor banco": 16, "valor contabilidad": 18, "diferencia": 15,
    "comprobante": 13, "documento": 15,
    "descripción banco": 32, "descripción contabilidad": 42, "descripción": 46,
    "confianza": 34,
    # Cruce DIAN:
    "nivel": 10, "fecha emisión": 13, "comprobante dian": 16, "nit": 13,
    "emisor": 32, "total": 16, "causación": 12, "referencia avansant": 18,
    "valor avansant": 16, "dif. valor": 15,
    "tercero avansant": 28, "fecha contable": 14, "cruzado el": 18, "candidatos": 60,
}


def _tipo_columna(nombre):
    key = nombre.lower()
    if key in _COLS_FECHA:
        return "fecha"
    if key in _COLS_MONEDA:
        return "moneda"
    if key in _COLS_ENTERO:
        return "entero"
    return "texto"


def build_tabla_workbook(df, meta, titulo, nombre_hoja="Detalle"):
    """Genera un Excel de una sola hoja con el membrete de ISTHO y la tabla recibida
    tal como se ve en la app (respetando el filtro aplicado)."""
    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    ws = workbook.add_worksheet(nombre_hoja[:31])

    columnas = [(c, c, _tipo_columna(c)) for c in df.columns]
    fmts, siguiente = _encabezado(workbook, ws, titulo, meta, len(columnas))
    siguiente += 1
    fin = _escribir_tabla(ws, fmts, siguiente, columnas, df)

    # Fila de totales para la primera columna monetaria que exista (la columna 0 la ocupa
    # la etiqueta "TOTAL", así que solo se totaliza si el valor está de la 1 en adelante).
    col_moneda = next((i for i, (_, _campo, tipo) in enumerate(columnas) if tipo == "moneda" and i >= 1), None)
    if not df.empty and col_moneda is not None:
        campo = columnas[col_moneda][1]
        _fila_total(ws, fmts, fin, len(columnas), f"TOTAL ({len(df)} movimientos)",
                    {col_moneda: float(pd.to_numeric(df[campo], errors="coerce").fillna(0).sum())})

    _ajustar_anchos(ws, [_ANCHOS_EXCEL.get(str(c).lower(), 20) for c in df.columns])
    workbook.close()
    return buffer.getvalue()


# --------------------------------------------------------------------------------
# Informe completo: Conciliados + Cruzados con diferencia + Por revisar, uno debajo del
# otro en UNA sola hoja, con las mismas columnas — para que quien lo reciba pueda borrar
# las filas de encabezado de sección y quedarse con una sola tabla unificada, sin tener
# que descargar cada hoja por separado y pegarlas a mano.
# --------------------------------------------------------------------------------
_COLUMNAS_RESUMEN = [
    ("Estado", "estado", "texto"),
    ("ID", "ID", "texto"),
    ("Fecha Banco", "Fecha Banco", "fecha"),
    ("Fecha Contabilidad", "Fecha Contabilidad", "fecha"),
    ("Valor Banco", "Valor Banco", "moneda"),
    ("Valor Contabilidad", "Valor Contabilidad", "moneda"),
    ("Diferencia", "Diferencia", "moneda"),
    ("Descripción Banco", "Descripción Banco", "texto"),
    ("Descripción Contabilidad", "Descripción Contabilidad", "texto"),
    ("Comprobante", "Comprobante", "texto"),
    ("Documento", "Documento", "texto"),
    ("Motivo", "Motivo", "texto"),
]


def _a_columnas_resumen(df, estado, campo_valor=None):
    """Lleva cualquiera de las 3 tablas (Conciliados/Diferencia/Posibles) al mismo juego de
    columnas, para que las tres se puedan apilar como una sola tabla. `campo_valor`: en
    Conciliados no hay "Valor Banco"/"Valor Contabilidad" separados (el cruce cuadra exacto),
    así que la única columna "Valor" que trae alimenta las dos, con Diferencia en 0."""
    out = pd.DataFrame(index=df.index)
    out["estado"] = estado
    out["ID"] = df["ID"] if "ID" in df.columns else ""
    out["Fecha Banco"] = df.get("Fecha Banco")
    out["Fecha Contabilidad"] = df.get("Fecha Contabilidad")
    if campo_valor:
        out["Valor Banco"] = df[campo_valor]
        out["Valor Contabilidad"] = df[campo_valor]
        out["Diferencia"] = 0.0
    else:
        out["Valor Banco"] = df.get("Valor Banco")
        out["Valor Contabilidad"] = df.get("Valor Contabilidad")
        out["Diferencia"] = df.get("Diferencia")
    out["Descripción Banco"] = df.get("Descripción Banco", "")
    out["Descripción Contabilidad"] = df.get("Descripción Contabilidad", "")
    out["Comprobante"] = df.get("Comprobante", "")
    out["Documento"] = df.get("Documento", "")
    out["Motivo"] = df.get("Motivo", "")
    return out


def _escribir_seccion(worksheet, fmts, start_row, titulo, df, n_columnas):
    """Una franja con el nombre de la sección (fila que se puede borrar entera) seguida de
    su tabla. Sin autofilter/freeze propios (`congelar_y_filtrar=False`): eso se decide una
    sola vez para toda la hoja, ver `build_resumen_completo_workbook`. Devuelve la fila
    siguiente libre."""
    worksheet.set_row(start_row, 22)
    worksheet.merge_range(start_row, 0, start_row, n_columnas - 1,
                           f"{titulo}  ({len(df)} movimiento(s))", fmts["total_label"])
    fin = _escribir_tabla(worksheet, fmts, start_row + 1, _COLUMNAS_RESUMEN, df,
                           congelar_y_filtrar=False)
    return fin + 1  # una fila de aire antes de la siguiente sección


def build_resumen_completo_workbook(df_conc, df_dif, df_posibles, meta):
    """Informe completo de un solo cruce: las tres tablas (Conciliados, Cruzados con
    diferencia, Por revisar) apiladas en una sola hoja, cada una con su propio encabezado de
    sección. Se salta la sección que venga vacía, para no dejar un encabezado sin nada debajo."""
    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    ws = workbook.add_worksheet("Informe completo")

    n_columnas = len(_COLUMNAS_RESUMEN)
    fmts, siguiente = _encabezado(workbook, ws, "CONCILIACIÓN BANCARIA — INFORME COMPLETO",
                                   meta, n_columnas)
    siguiente += 1
    # Un solo congelado para TODA la hoja, justo debajo del membrete — así el membrete queda
    # fijo arriba y el resto (las 3 secciones completas) se desplaza libremente. Antes cada
    # sección congelaba la suya y solo quedaba la última, que caía casi al final de la hoja y
    # dejaba ver "solo la pantalla principal" sin poder bajar más.
    ws.freeze_panes(siguiente, 0)

    secciones = [
        ("CONCILIADOS", _a_columnas_resumen(df_conc, "Conciliado", campo_valor="Valor")),
        ("CRUZADOS CON DIFERENCIA", _a_columnas_resumen(df_dif, "Cruzado con diferencia")),
        ("POR REVISAR", _a_columnas_resumen(df_posibles, "Por revisar")),
    ]
    for titulo, df in secciones:
        if df.empty:
            continue
        siguiente = _escribir_seccion(ws, fmts, siguiente, titulo, df, n_columnas)

    _ajustar_anchos(ws, [_ANCHOS_EXCEL.get(campo.lower(), 20) for _, campo, _ in _COLUMNAS_RESUMEN])
    workbook.close()
    return buffer.getvalue()


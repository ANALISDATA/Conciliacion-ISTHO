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

COLOR_OSCURO = "#132A20"
COLOR_VERDE = "#4CAF50"
COLOR_NARANJA = "#E8772E"
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


def _escribir_tabla(worksheet, fmts, start_row, columnas, df):
    """columnas: lista de tuplas (titulo, campo, tipo) tipo en {'texto','moneda','fecha','entero'}."""
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
_COLS_MONEDA = {"valor", "valor banco", "valor contabilidad", "diferencia", "total"}
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


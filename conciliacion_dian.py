"""Motor del cruce DIAN vs Avansant.

Completamente independiente de `conciliacion.py` (el motor de BANCARIO): no comparte
estado, no comparte estructuras de datos, y no lo modifica. La única función que se
reutiliza es `parse_money()`, una utilidad pura (sin efectos secundarios) para no duplicar
el parseo de dinero.

## La idea del cruce

DIAN es la base: cada factura/nota que emitieron los proveedores y que llegó por la DIAN.
Avansant es donde contabilidad la registra como "causación" (cuenta por pagar). El cruce
busca, para cada documento de DIAN, si existe su causación correspondiente en Avansant —
y si no la encuentra, es una factura que contabilidad todavía no ha digitado.

## Niveles de confianza (de más a menos seguro)

1. **Alta**: el comprobante de DIAN (`Prefijo-Folio`) coincide, letra por letra
   (normalizado), con la referencia que alguien escribió a mano en Avansant.
2. **Media**: el texto no calza (típicamente un error de digitación), pero el NIT del
   proveedor y el valor exacto sí — y solo hay UNA causación posible con esa combinación.
3. **Baja**: tampoco calza el valor, pero el NIT y el número de folio (sin el prefijo) sí
   coinciden, y también hay un único candidato.

En cualquier nivel, si hay MÁS DE UN candidato posible, no se elige a ciegas: queda en
"Ambiguos" para que alguien lo revise. Ver `documentos/DOCUMENTACION_DIAN.md` para el
detalle completo y las cifras con las que se validó.
"""
import re
import unicodedata
from collections import defaultdict
from datetime import datetime

import pandas as pd

from conciliacion import parse_money

_APP_RESPONSE = "application response"
_NOTA_CREDITO = "nota de crédito electrónica"


def _norm_texto(v):
    """Mayúsculas, sin tildes, solo letras y números (sin guiones/espacios/puntos): así un
    espacio de más o un punto residual no rompen una coincidencia que en el fondo es la
    misma referencia."""
    s = unicodedata.normalize("NFKD", str(v or "").upper())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", "", s)


def _folio_digitos(v):
    """El número de folio dentro de una referencia tipo 'PREFIJO-12345': los dígitos
    después del último guion (o del texto completo si no trae guion)."""
    s = str(v or "").strip()
    if "-" in s:
        s = s.rsplit("-", 1)[1]
    s = re.sub(r"\D", "", s)
    return s.lstrip("0") or ("0" if s else "")


def _centavos(v):
    try:
        return round(float(v) * 100)
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------------
# Carga de archivos
# --------------------------------------------------------------------------------
def load_dian(file):
    """Informe de documentos recibidos de la DIAN (.xlsx, tal como se descarga).

    Descarta las filas `Tipo de documento = "Application response"` (son acuses de
    recibo automáticos, no facturas ni notas reales) y arma la llave de cruce:
    `Comprobante = Prefijo-Folio` (o solo el Folio si el documento no trae prefijo)."""
    df = pd.read_excel(file, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    col_emision = next((c for c in df.columns if c.lower().startswith("fecha emisi")), None)
    faltantes = [c for c in ("Tipo de documento", "Folio", "NIT Emisor", "Nombre Emisor", "Total")
                 if c not in df.columns]
    if not col_emision:
        faltantes.append("Fecha Emisión")
    if faltantes:
        raise ValueError(
            f"Al informe DIAN le faltan las columnas {faltantes}. "
            f"Columnas encontradas: {list(df.columns)}"
        )
    col_recepcion = next((c for c in df.columns if c.lower().startswith("fecha recepci")), None)

    out = pd.DataFrame()
    out["tipo_doc"] = df["Tipo de documento"].astype(str).str.strip()
    out = out[out["tipo_doc"].str.lower() != _APP_RESPONSE].copy()
    df = df.loc[out.index]

    out["cufe"] = df.get("CUFE/CUDE", "").fillna("").astype(str).str.strip()
    out["folio"] = df["Folio"].apply(lambda v: str(v).strip() if pd.notna(v) else "")
    out["prefijo"] = df.get("Prefijo", "").fillna("").astype(str).str.strip()
    out["comprobante"] = [f"{p}-{f}" if p else f for p, f in zip(out["prefijo"], out["folio"])]
    out["fecha_emision"] = pd.to_datetime(df[col_emision], dayfirst=True, errors="coerce").dt.date
    out["fecha_recepcion"] = (pd.to_datetime(df[col_recepcion], dayfirst=True, errors="coerce")
                               if col_recepcion else pd.NaT)
    out["nit_emisor"] = df["NIT Emisor"].apply(lambda v: str(v).strip() if pd.notna(v) else "")
    out["nombre_emisor"] = df["Nombre Emisor"].astype(str).str.strip()
    out["total"] = df["Total"].apply(parse_money)
    out["estado"] = df.get("Estado", "").fillna("")
    out["es_nota_credito"] = out["tipo_doc"].str.lower() == _NOTA_CREDITO

    out = out.dropna(subset=["fecha_emision"]).reset_index(drop=True)
    out["id"] = out.index
    return out


# Posición (0-indexada) de cada columna útil dentro del bloque "Contabilización" del
# informe de Avansant. Igual que hace el Power Query actual: el "Nro." aparece dos veces
# con el mismo título (una vez como ID de la causación, otra como referencia de texto que
# alguien escribió a mano), así que no se puede leer por nombre de columna — hay que leerlo
# por posición.
_COLS_AVANSAT = {
    "anio": 0, "tipo": 3, "causacion": 4, "documento": 5, "referencia": 6,
    "fecha_creacion": 7, "detalle": 8, "nit": 11, "tercero": 12, "valor": 14,
    "fecha_contable": 15, "estado": 16, "usuario": 17,
}
# Para poder avisar temprano y claro si el formato del reporte cambió, en vez de leer
# columnas equivocadas en silencio (que es lo que le pasaría al Power Query actual).
_ENCABEZADOS_ESPERADOS = {0: "año", 3: "tipo", 4: "nro", 6: "nro", 8: "detalle",
                          12: "tercero", 14: "valor"}


def load_avansant(file):
    """Informe de comprobantes de Avansant (.xlsx, tal como se descarga): trae 3 filas de
    encabezado (ruta de navegación, secciones "Contabilización"/"Reversión", y los
    títulos reales de columna) antes de los datos. Se queda solo con las filas
    `Tipo = CAUSACION` (cuentas por pagar), que es contra lo que se cruza DIAN."""
    crudo = pd.read_excel(file, header=None)

    fila_encabezado = next(
        (i for i in range(min(10, len(crudo))) if str(crudo.iat[i, 0]).strip() == "Año"), None)
    if fila_encabezado is None:
        raise ValueError(
            "No se encontró la fila de encabezados ('Año') en el informe de Avansant. "
            "¿Es el archivo de comprobantes correcto?"
        )

    encabezados = crudo.iloc[fila_encabezado]
    for pos, esperado in _ENCABEZADOS_ESPERADOS.items():
        real = str(encabezados[pos]).strip().lower()
        if not real.startswith(esperado):
            raise ValueError(
                f"El informe de Avansant no tiene el formato esperado: en la columna {pos + 1} "
                f"se esperaba algo como «{esperado.title()}» y se encontró «{encabezados[pos]}». "
                "Puede que el reporte haya cambiado de estructura."
            )

    datos = crudo.iloc[fila_encabezado + 1:].reset_index(drop=True)
    out = pd.DataFrame()
    out["anio"] = pd.to_numeric(datos[_COLS_AVANSAT["anio"]], errors="coerce")
    out = out[out["anio"].notna()].copy()
    datos = datos.loc[out.index]

    out["tipo"] = datos[_COLS_AVANSAT["tipo"]].astype(str).str.strip()
    out["causacion"] = datos[_COLS_AVANSAT["causacion"]].apply(
        lambda v: "" if pd.isna(v) else str(v).strip())
    out["documento"] = datos[_COLS_AVANSAT["documento"]].apply(
        lambda v: "" if pd.isna(v) else str(v).strip())
    out["referencia"] = datos[_COLS_AVANSAT["referencia"]].apply(
        lambda v: "" if pd.isna(v) else str(v).strip())
    out["fecha_creacion"] = pd.to_datetime(datos[_COLS_AVANSAT["fecha_creacion"]], errors="coerce")
    out["detalle"] = datos[_COLS_AVANSAT["detalle"]].apply(
        lambda v: "" if pd.isna(v) else str(v).strip())
    out["nit"] = datos[_COLS_AVANSAT["nit"]].apply(lambda v: "" if pd.isna(v) else str(v).strip())
    out["tercero"] = datos[_COLS_AVANSAT["tercero"]].apply(
        lambda v: "" if pd.isna(v) else str(v).strip())
    out["valor"] = datos[_COLS_AVANSAT["valor"]].apply(parse_money)
    out["fecha_contable"] = pd.to_datetime(datos[_COLS_AVANSAT["fecha_contable"]], errors="coerce").dt.date
    out["estado"] = datos[_COLS_AVANSAT["estado"]].apply(lambda v: "" if pd.isna(v) else str(v).strip())
    out["usuario"] = datos[_COLS_AVANSAT["usuario"]].apply(lambda v: "" if pd.isna(v) else str(v).strip())

    out = out[out["tipo"].str.upper() == "CAUSACION"].reset_index(drop=True)
    out["id"] = out.index
    return out


# --------------------------------------------------------------------------------
# El cruce
# --------------------------------------------------------------------------------
def reconciliar_dian(df_dian, df_avansat):
    """Cruza DIAN (base) contra las causaciones de Avansant, en los 3 niveles descritos
    arriba del archivo. Las notas de crédito NO se intentan cruzar (hoy no tienen nada
    contra qué cruzar en Avansant — ver documentación) y quedan en su propia categoría.

    Devuelve `(cruces, ambiguos, duplicados)`:
    - `cruces`: cada uno con `id`, `origen`, `nivel`, `motivo`, `dian_id`, `avansat_id`.
    - `ambiguos`: documentos DIAN con más de un candidato posible — requieren revisión.
    - `duplicados`: referencias de Avansant que aparecen en 2+ causaciones con distinto
      NIT (no se puede saber a cuál factura corresponde cada una)."""
    av = df_avansat.copy()
    av["matched"] = False

    por_texto = defaultdict(list)
    por_nit_valor = defaultdict(list)
    por_nit_folio = defaultdict(list)
    for i, row in av.iterrows():
        nit = row["nit"]
        texto = _norm_texto(row["referencia"])
        if texto:
            por_texto[texto].append(i)
        por_nit_valor[(nit, _centavos(row["valor"]))].append(i)
        f = _folio_digitos(row["referencia"])
        if f:
            por_nit_folio[(nit, f)].append(i)

    duplicados = []
    for texto, idxs in por_texto.items():
        if len(idxs) < 2:
            continue
        nits = {av.at[i, "nit"] for i in idxs}
        if len(nits) > 1:
            duplicados.append({"referencia": av.loc[idxs[0], "referencia"],
                                "avansat_ids": [int(i) for i in idxs]})

    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    cruces, ambiguos = [], []
    facturas = df_dian[~df_dian["es_nota_credito"]]
    # Pendientes: el orden en que aparecen las filas en el archivo NO puede decidir el
    # resultado — si una factura A resuelve por Media antes de que la factura B (procesada
    # después) alcance a intentar su Alta, B perdería un candidato que en realidad le
    # pertenecía a ella. Por eso cada nivel corre como una PASADA completa sobre todas las
    # facturas que sigan pendientes, igual que hace `reconciliar()` en BANCARIO — así el
    # resultado no depende de en qué orden vengan las filas del Excel.
    pendientes = list(facturas["id"])

    def _pasada(nivel, motivo, obtener_candidatos, motivo_ambiguo):
        nonlocal pendientes
        siguen_pendientes = []
        for did in pendientes:
            d = df_dian.loc[did]
            candidatos = obtener_candidatos(d)
            if len(candidatos) == 1:
                avi = candidatos[0]
                av.at[avi, "matched"] = True
                cruces.append({"origen": "Automática", "nivel": nivel, "motivo": motivo,
                                "fecha_hora": ahora, "dian_id": int(did), "avansat_id": int(avi)})
            elif len(candidatos) > 1:
                ambiguos.append({"dian_id": int(did), "avansat_ids": [int(i) for i in candidatos],
                                  "motivo": motivo_ambiguo})
            else:
                siguen_pendientes.append(did)
        pendientes = siguen_pendientes

    # Nivel Alta: texto exacto (normalizado). Si el candidato SÍ trae NIT y no coincide con
    # el de la factura, no se acepta a ciegas: se descarta como candidato de este nivel
    # (puede rescatarse en la pasada de Media o Baja).
    _pasada(
        "Alta", "El comprobante coincide exacto con la referencia en Avansant",
        lambda d: [i for i in por_texto.get(_norm_texto(d["comprobante"]), [])
                   if not av.at[i, "matched"]
                   and (not av.at[i, "nit"] or av.at[i, "nit"] == d["nit_emisor"])],
        "Varias causaciones con la misma referencia y el mismo NIT")

    # Nivel Media: mismo NIT y mismo valor exacto (la referencia no calzó por texto,
    # típicamente un error de digitación).
    _pasada(
        "Media", "Mismo NIT y mismo valor exacto (la referencia no coincidía)",
        lambda d: [i for i in por_nit_valor.get((d["nit_emisor"], _centavos(d["total"])), [])
                   if not av.at[i, "matched"]],
        "Varias causaciones del mismo NIT con el mismo valor")

    # Nivel Baja: mismo NIT y mismo número de folio, sin el prefijo (por si el prefijo se
    # escribió distinto o se omitió).
    _pasada(
        "Baja", "Mismo NIT y mismo número de folio (el prefijo no coincidía)",
        lambda d: ([i for i in por_nit_folio.get((d["nit_emisor"], _folio_digitos(d["comprobante"])), [])
                    if not av.at[i, "matched"]] if _folio_digitos(d["comprobante"]) else []),
        "Varias causaciones del mismo NIT con el mismo folio")

    # Lo que sigue en `pendientes` tras las 3 pasadas queda sin cruzar — se deriva en
    # construir_vistas_dian() a partir de lo que no aparece ni en `cruces` ni en `ambiguos`.

    for n, c in enumerate(cruces, start=1):
        c["id"] = f"D-{n:04d}"

    return cruces, ambiguos, duplicados


def construir_vistas_dian(df_dian, df_avansat, cruces, ambiguos):
    """Deriva todas las tablas que ve el usuario a partir de la lista de cruces — igual
    filosofía que BANCARIO: los datos originales (`df_dian`, `df_avansat`) nunca se tocan,
    así que deshacer un cruce manual es tan simple como quitarlo de la lista."""
    dian_usados = {c["dian_id"] for c in cruces} | {a["dian_id"] for a in ambiguos}
    avansat_usados = {c["avansat_id"] for c in cruces}

    filas = []
    for c in cruces:
        d, a = df_dian.loc[c["dian_id"]], df_avansat.loc[c["avansat_id"]]
        filas.append({
            "ID": c["id"], "Origen": c["origen"], "Nivel": c["nivel"],
            "Fecha Emisión": d["fecha_emision"], "Comprobante DIAN": d["comprobante"],
            "NIT": d["nit_emisor"], "Emisor": d["nombre_emisor"], "Total": d["total"],
            "Causación": a["causacion"], "Valor Avansant": a["valor"],
            # Un cruce puede ser correcto (misma factura, mismo NIT, misma referencia) y aun
            # así traer un valor distinto: una retención que la contabilidad registró neta, o
            # sencillamente un error al digitar. Antes no se veía en ninguna parte — el cruce
            # salía como "Alta" y nadie se enteraba de la diferencia. Ahora va en su propia
            # columna para que contabilidad la revise.
            "Dif. valor": float(d["total"]) - float(a["valor"]),
            "Referencia Avansant": a["referencia"],
            "Tercero Avansant": a["tercero"], "Fecha Contable": a["fecha_contable"],
            "Motivo": c["motivo"], "Cruzado el": c["fecha_hora"],
        })
    df_cruzados = pd.DataFrame(filas)

    filas = []
    for a in ambiguos:
        d = df_dian.loc[a["dian_id"]]
        candidatos = "; ".join(
            f"Causación {df_avansat.at[i, 'causacion']} (${df_avansat.at[i, 'valor']:,.0f}, "
            f"ref «{df_avansat.at[i, 'referencia']}», {df_avansat.at[i, 'tercero']})"
            for i in a["avansat_ids"])
        filas.append({
            "Fecha Emisión": d["fecha_emision"], "Comprobante DIAN": d["comprobante"],
            "NIT": d["nit_emisor"], "Emisor": d["nombre_emisor"], "Total": d["total"],
            "Motivo": a["motivo"], "Candidatos": candidatos,
        })
    df_ambiguos = pd.DataFrame(filas)

    df_notas_credito = df_dian[df_dian["es_nota_credito"]].reset_index(drop=True)

    mascara_pendiente = (~df_dian.index.isin(dian_usados)) & (~df_dian["es_nota_credito"])
    df_sin_coincidencia = df_dian[mascara_pendiente].reset_index(drop=True)

    df_avansat_pendiente = df_avansat[~df_avansat.index.isin(avansat_usados)].reset_index(drop=True)

    return df_cruzados, df_ambiguos, df_notas_credito, df_sin_coincidencia, df_avansat_pendiente


def vista_duplicados(df_avansat, duplicados):
    """Una fila por cada causación involucrada en una referencia duplicada, para revisar
    cuál corresponde a cuál (no se puede saber solo con los datos disponibles)."""
    filas = []
    for dup in duplicados:
        for i in dup["avansat_ids"]:
            a = df_avansat.loc[i]
            filas.append({"Referencia": dup["referencia"], "Causación": a["causacion"],
                          "NIT": a["nit"], "Tercero": a["tercero"], "Valor": a["valor"],
                          "Fecha Contable": a["fecha_contable"], "Detalle": a["detalle"]})
    return pd.DataFrame(filas)


# --------------------------------------------------------------------------------
# Cruce manual: la persona teclea el número de causación que ya encontró buscando a
# mano en Avansant (así es como ellas mismas confirman una factura hoy).
# --------------------------------------------------------------------------------
def buscar_causacion(df_avansat, numero_causacion):
    """Busca una causación por su número. Devuelve la fila (para mostrarla y que la
    persona confirme que es la correcta) o `None` si no existe."""
    numero = str(numero_causacion).strip()
    if not numero:
        return None
    coincide = df_avansat[df_avansat["causacion"].astype(str).str.strip() == numero]
    return coincide.iloc[0] if not coincide.empty else None


def crear_cruce_manual_dian(cruces, dian_id, avansat_id, df_avansat):
    """Agrega un cruce manual a la lista y la devuelve (no muta la original). Queda con
    nivel «Manual» para distinguirlo a simple vista de los automáticos."""
    a = df_avansat.loc[avansat_id]
    usados = [c["id"] for c in cruces if str(c.get("id", "")).startswith("M-")]
    nuevo = {
        "id": f"M-{len(usados) + 1:04d}", "origen": "Manual", "nivel": "Manual",
        "motivo": f"Cruzado a mano con la causación {a['causacion']}",
        "fecha_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "dian_id": int(dian_id), "avansat_id": int(avansat_id),
    }
    return cruces + [nuevo], nuevo["id"]


def eliminar_cruces_dian(cruces, ids):
    """Deshace los cruces indicados. Los documentos vuelven solos a pendientes porque las
    vistas se derivan de esta lista."""
    ids = set(ids)
    return [c for c in cruces if c["id"] not in ids]

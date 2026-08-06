"""Lógica de carga y cruce (fecha + valor + nombre) entre extracto bancario y libro auxiliar."""
import re
import unicodedata
from collections import defaultdict
from datetime import datetime, timedelta
from functools import lru_cache
from itertools import combinations

import pandas as pd

_STOPWORDS = {
    "PAGO", "PAGOS", "PROVE", "PROVEEDOR", "PROVEEDORES", "TRANSFERENCIA", "TRANSFERENCIAS",
    "CTA", "CUENTA", "SUC", "SUCURSAL", "VIRTUAL", "TRANSPORTE", "TRANSPORTES", "FACTURA",
    "FACTURAS", "ANTICIPO", "ANTICIPOS", "MANIFIESTO", "PLACAX", "PLACA", "SOBRE", "NOMIN",
    "NOMINA", "SERVICIO", "SERVICIOS", "BANCO", "BANCOS", "OTROS", "OTRO", "DE", "LA", "EL",
    "LOS", "LAS", "Y", "A", "AL", "DEL", "RECAUDO", "LIQUIDACION", "LIQUIDACIONES", "PSE",
    "IMPTO", "GOBIERNO", "COBRO", "IVA", "AUTOMATICOS", "ABONO", "INTERESES", "AHORROS",
    "GASTOS", "BANCARIOS", "BANCARIA", "NRO", "RECIBO", "RECIBOS", "EGRESO", "INGRESO",
    "COMPROBANTE", "CIA", "SAS", "LTDA", "S", "PROV",
    # El nombre de la propia empresa no identifica a ninguna contraparte: la contabilidad lo
    # usa como beneficiario en retiros, traslados y movimientos propios. Tratarlo como nombre
    # bloqueaba cruces exactos (ej. "RETIRO CAJERO ..." contra un asiento a nombre de ISTHO).
    "ISTHO",
    # El canal de la transferencia tampoco identifica a la contraparte. Sin esto, "PAGO
    # INTERBANC <empresa>" no cruzaba contra "<empresa> INDUSTRIAL S.A." aunque el valor
    # fuera exacto y la fecha estuviera a solo unos días: "INTERBANC" quedaba en el conjunto
    # de palabras del banco sin nada equivalente del lado contable, y eso basta para que
    # `_mismo_nombre` rechace el cruce por completo (exige que TODAS las palabras casen).
    "INTERBANC", "INTERBANCARIO", "INTERBANCARIA",
}

# Longitud a partir de la cual se asume que el extracto cortó la descripción a mitad de
# palabra. Bancolombia trunca alrededor de los 28-30 caracteres.
_LARGO_TRUNCADO = 26


@lru_cache(maxsize=4096)
def _tokens(texto):
    """Palabras significativas del texto, en mayúsculas y SIN TILDES: el banco escribe
    'LOGISTICOS' donde la contabilidad escribe 'LOGÍSTICOS' y son la misma empresa."""
    plano = unicodedata.normalize("NFKD", str(texto).upper())
    plano = "".join(c for c in plano if not unicodedata.combining(c)).replace("Ñ", "N")
    palabras = re.split(r"[^A-Z0-9]+", plano)
    return frozenset(p for p in palabras if len(p) >= 3 and p not in _STOPWORDS)


@lru_cache(maxsize=4096)
def _tokens_ordenados(texto):
    """Igual que `_tokens` pero conservando el orden, para poder descartar la última palabra
    cuando el extracto viene cortado a mitad de palabra."""
    plano = unicodedata.normalize("NFKD", str(texto).upper())
    plano = "".join(c for c in plano if not unicodedata.combining(c)).replace("Ñ", "N")
    return [p for p in re.split(r"[^A-Z0-9]+", plano)
            if len(p) >= 3 and p not in _STOPWORDS]


@lru_cache(maxsize=400_000)
def _mismo_nombre(t1, t2):
    """Compara dos conjuntos de palabras SIN importar el orden. Devuelve True solo si pueden
    ser la misma persona o entidad: todas las palabras del nombre más corto deben aparecer
    en el más largo (aceptando prefijos, porque el banco trunca)."""
    if not t1 or not t2:
        return None           # no hay nada que validar
    pequeno, grande = (t1, t2) if len(t1) <= len(t2) else (t2, t1)
    usados, emparejadas = set(), 0
    for a in pequeno:
        for b in grande:
            if b in usados:
                continue
            if a == b or (min(len(a), len(b)) >= 3 and (a.startswith(b) or b.startswith(a))):
                usados.add(b)
                emparejadas += 1
                break
    if emparejadas != len(pequeno):
        return False          # sobra alguna palabra: son personas distintas
    # Con una sola palabra en común solo se acepta si es larga y por tanto distintiva
    # (una razón social como "BANCOOMEVA"), nunca un nombre de pila suelto como "JUAN".
    return emparejadas >= 2 or len(next(iter(pequeno))) >= 5


# Caché grande a propósito: con ~1.000 movimientos de banco contra ~500 asientos se
# consultan cientos de miles de pares y una caché pequeña se desaloja sola.
@lru_cache(maxsize=400_000)
def _validar_nombre(desc_banco, nombre_libro):
    """VALIDACIÓN OBLIGATORIA antes de cualquier cruce automático: si en los dos lados se
    puede leer un nombre, tienen que ser la misma persona.

    - Se separa en palabras, se pasa a mayúsculas y se quitan conectores y tecnicismos.
    - **No importa el orden**: se comparan como conjuntos, así que "Juan Carlos Marín" y
      "Marín Juan Carlos" son la misma persona.
    - Todas las palabras del nombre más corto deben estar en el más largo. Por eso
      "Juan Carlos Restrepo" y "Juan Carlos Marín" NO cruzan: sobra un apellido distinto.
    - Se tolera el truncamiento del banco ("TOBON EDWIN ALB" por "TOBON JARAMILLO EDWIN
      ALBERTO"), aceptando que una palabra sea prefijo de la otra.

    Si en alguno de los dos lados no hay ningún nombre identificable (ej. el extracto dice
    solo "TRANSFERENCIA CTA SUC VIRTUAL", o la contabilidad pone a la propia ISTHO como
    beneficiario), no hay nada que validar y se deja que decidan el valor y la fecha.
    """
    resultado = _mismo_nombre(_tokens(desc_banco), _tokens(nombre_libro))
    if resultado is not None and resultado:
        return True
    if resultado is None:
        return True

    # Segundo intento por truncamiento: si la descripción del extracto viene cortada, su
    # última palabra suele estar partida a la mitad ("comfenalco cred" por "COMFENALCO
    # CREDITOS") y ensucia la comparación. Se reintenta sin ella.
    if len(str(desc_banco).strip()) >= _LARGO_TRUNCADO:
        lista = _tokens_ordenados(desc_banco)
        if len(lista) >= 2:
            return _mismo_nombre(frozenset(lista[:-1]), _tokens(nombre_libro)) is True
    return False


@lru_cache(maxsize=400_000)
def _similitud_nombre(desc_banco, beneficiario_libro):
    """Cuenta cuántas palabras (ya sin conectores/tecnicismos) tienen en común, dando
    también crédito cuando una es prefijo de la otra (los bancos truncan los nombres,
    ej. 'CELULA' por 'CELULAR', 'ALB' por 'ALBERTO')."""
    t1, t2 = _tokens(desc_banco), _tokens(beneficiario_libro)
    if not t1 or not t2:
        return 0
    usados = set()
    hits = 0
    for a in t1:
        for b in t2:
            if b in usados:
                continue
            if a == b or (len(a) >= 4 and len(b) >= 4 and (a.startswith(b) or b.startswith(a))):
                hits += 1
                usados.add(b)
                break
    return hits


# --------------------------------------------------------------------------------
# Conceptos: el banco itemiza (un movimiento por persona/cobro) mientras la contabilidad
# registra un solo asiento consolidado. Reconocer el concepto permite sumar todos los
# movimientos del banco de ese tipo y cruzarlos contra el asiento único.
# Ej: 3 "PAGO A NOMIN <empleado>" que suman lo mismo que 1 "PAGO NOMINA 31 MAYO 2026".
# --------------------------------------------------------------------------------
_CONCEPTOS = [
    ("Nómina", (r"\bNOMIN\w*", r"\bPAGO\s+DE\s+NOMINA")),
    ("Seguridad social", (r"\bPILA\b", r"SEGURIDAD\s+SOCIAL", r"\bAPORTES?\b", r"\bEPS\b",
                           r"\bARL\b", r"\bPENSION\w*", r"\bCESANTIAS?\b", r"\bPARAFISCAL\w*")),
    ("4x1000 (GMF)", (r"4\s*X\s*1000", r"\bGMF\b", r"IMPTO\s+GOBIERNO")),
    ("IVA", (r"\bIVA\b",)),
    ("Intereses", (r"\bINTERES\w*",)),
    ("Retenciones", (r"\bRETEFUENTE\b", r"\bRETENCION\w*", r"\bRETE\s?ICA\b", r"\bRETE\s?IVA\b")),
    ("Comisiones y servicios bancarios", (r"\bCOMISION\w*", r"CUOTA\s+MANEJO", r"GASTOS?\s+BANCARIOS?",
                                           r"CUOTA\s+PLAN", r"SERVICIO\s+PAGO", r"\bCHEQUERA\b")),
]
_CONCEPTOS_COMPILADOS = [(etiqueta, [re.compile(p) for p in patrones])
                         for etiqueta, patrones in _CONCEPTOS]


@lru_cache(maxsize=8192)
def _concepto(texto):
    """Etiqueta del concepto contable al que pertenece un movimiento, o None si no aplica."""
    t = str(texto).upper()
    for etiqueta, patrones in _CONCEPTOS_COMPILADOS:
        if any(p.search(t) for p in patrones):
            return etiqueta
    return None


def parse_money(value):
    """Convierte valores tipo '168,900.00', '1.234.567,89' o numéricos a float."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace("$", "").replace(" ", "").strip()
    if s in ("", "-", "nan", "None", "NaN"):
        return 0.0
    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo = True
        s = s[1:-1]
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")  # 1.234.567,89 -> 1234567.89
        else:
            s = s.replace(",", "")  # 1,234,567.89 -> 1234567.89
    elif "," in s:
        entero, _, dec = s.rpartition(",")
        s = entero.replace(",", "") + "." + dec if len(dec) == 2 else s.replace(",", "")
    try:
        v = float(s)
        return -v if negativo else v
    except ValueError:
        return 0.0


def _nombre_archivo(file):
    return str(getattr(file, "name", file)).lower()


def _es_csv(file):
    return _nombre_archivo(file).endswith(".csv")


def _leer_csv(file, header):
    """Intenta varias combinaciones de separador/codificación típicas de exportes bancarios/contables."""
    intentos = [
        {"sep": None, "engine": "python", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "latin-1"},
        {"sep": ",", "encoding": "latin-1"},
    ]
    ultimo_error = None
    for kwargs in intentos:
        try:
            if hasattr(file, "seek"):
                file.seek(0)
            return pd.read_csv(file, header=header, dtype=str, **kwargs)
        except Exception as e:  # noqa: BLE001
            ultimo_error = e
    raise ValueError(f"No se pudo leer el archivo CSV: {ultimo_error}")


def _leer_tabla(file, header):
    if _es_csv(file):
        return _leer_csv(file, header)
    if hasattr(file, "seek"):
        file.seek(0)
    return pd.read_excel(file, header=header)


def load_extracto(file):
    """Extracto Bancolombia (xlsx o csv): col3=fecha, col5=valor, col7=descripción, sin encabezado."""
    df = _leer_tabla(file, header=None)
    if df.shape[1] < 8:
        raise ValueError(
            f"El archivo tiene {df.shape[1]} columnas y se esperaban al menos 8 "
            "(formato de extracto Bancolombia sin encabezado)."
        )

    # Si la primera fila no parece un movimiento real (ej. trae encabezados), se descarta.
    primera_fecha = pd.to_datetime(df.iloc[0, 3], errors="coerce", dayfirst=True)
    if pd.isna(primera_fecha):
        df = df.iloc[1:].reset_index(drop=True)

    df = df.iloc[:, [3, 5, 7]].copy()
    df.columns = ["fecha", "valor", "descripcion"]
    df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce").dt.date
    df["valor"] = df["valor"].apply(parse_money)
    df["descripcion"] = df["descripcion"].astype(str).str.strip()
    df = df.dropna(subset=["fecha"]).reset_index(drop=True)
    df["id"] = df.index
    df["tipo"] = df["valor"].apply(lambda v: "Entrada" if v >= 0 else "Salida")
    return df


def load_libro_auxiliar(file):
    """Libro auxiliar contable (xlsx o csv): valor = DÉBITO (entra al banco) - CRÉDITO (sale del banco)."""
    df = _leer_tabla(file, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = [c for c in ("FECHA", "DETALLE", "DÉBITO", "CRÉDITO") if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"Al libro auxiliar le faltan las columnas {faltantes}. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    debito = df["DÉBITO"].apply(parse_money)
    credito = df["CRÉDITO"].apply(parse_money)

    out = pd.DataFrame()
    out["fecha"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce").dt.date
    out["valor"] = debito - credito
    out["debito"] = debito
    out["credito"] = credito
    out["detalle"] = df["DETALLE"].astype(str).str.strip()
    out["beneficiario"] = df.get("NOMBRE BENEFICIARIO", "").astype(str).str.strip()
    out["descripcion"] = out["detalle"] + " - " + out["beneficiario"]
    out["documento"] = df.get("DOCUMENTO")
    out["comprobante"] = df.get("COMPROBANTE")
    out["tipo_comprobante"] = df.get("TIPO COMPROBANTE")
    out["cuenta_nombre"] = df.get("NOMBRE AUXILIAR")

    out = out.dropna(subset=["fecha"]).reset_index(drop=True)
    out["id"] = out.index
    out["tipo"] = out["valor"].apply(lambda v: "Entrada" if v >= 0 else "Salida")
    return out


def _centavos(v):
    return int(round(v * 100))


def _buscar_subconjunto(valores, objetivo_centavos, max_tam=6):
    """valores: lista de (idx, valor_centavos) del mismo signo que el objetivo.
    Devuelve la lista de idx cuya suma coincide exactamente con el objetivo, o None."""
    n = len(valores)
    if n < 2:
        return None
    for tam in range(2, min(max_tam, n) + 1):
        for combo in combinations(range(n), tam):
            if sum(valores[i][1] for i in combo) == objetivo_centavos:
                return [valores[i][0] for i in combo]
    return None


def _agrupar_por_fecha(banco, libro, tolerancia_dias=3, tolerancia_dias_nombre=15, max_grupo=6,
                        max_candidatos_sin_nombre=18):
    """Busca casos donde varios movimientos de un lado suman exactamente el valor de uno o
    varios movimientos del otro lado. Si hay coincidencia de nombre/beneficiario, busca dentro
    de una ventana de fecha amplia (`tolerancia_dias_nombre`, igual que hace el cruce 1 a 1 con
    nombre); si no hay señal de nombre, solo agrupa dentro de la tolerancia normal y nada más
    cuando hay pocos candidatos, para no disparar el tiempo de cálculo."""
    grupos = []

    banco_por_fecha = defaultdict(list)
    for i in banco.index:
        banco_por_fecha[banco.at[i, "fecha"]].append(i)
    libro_por_fecha = defaultdict(list)
    for i in libro.index:
        libro_por_fecha[libro.at[i, "fecha"]].append(i)

    def _ventana(bucket, fecha_centro, dias):
        out = []
        for dd in range(-dias, dias + 1):
            out.extend(bucket.get(fecha_centro + timedelta(days=dd), []))
        return out

    # A) varios movimientos de banco -> un solo registro contable
    for li in [i for i in libro.index if not libro.at[i, "matched"]]:
        if libro.at[li, "matched"]:
            continue
        objetivo = _centavos(libro.at[li, "valor"])
        fecha_l = libro.at[li, "fecha"]
        beneficiario = libro.at[li, "beneficiario"]

        amplios = [i for i in _ventana(banco_por_fecha, fecha_l, tolerancia_dias_nombre)
                   if not banco.at[i, "matched"] and (banco.at[i, "valor"] >= 0) == (libro.at[li, "valor"] >= 0)
                   and _similitud_nombre(banco.at[i, "descripcion"], beneficiario) >= 1
                   and _validar_nombre(banco.at[i, "descripcion"], beneficiario)]
        candidatos = amplios
        if not candidatos:
            estrictos = [i for i in _ventana(banco_por_fecha, fecha_l, tolerancia_dias)
                         if not banco.at[i, "matched"] and (banco.at[i, "valor"] >= 0) == (libro.at[li, "valor"] >= 0)]
            candidatos = estrictos if len(estrictos) <= max_candidatos_sin_nombre else []

        valores = [(i, _centavos(banco.at[i, "valor"])) for i in candidatos]
        combo = _buscar_subconjunto(valores, objetivo, max_grupo)
        if combo:
            for i in combo:
                banco.at[i, "matched"] = True
            libro.at[li, "matched"] = True
            grupos.append({"tipo": "banco->libro", "banco_ids": combo, "libro_ids": [li]})

    # B) varios registros contables -> un solo movimiento de banco
    for bi in [i for i in banco.index if not banco.at[i, "matched"]]:
        if banco.at[bi, "matched"]:
            continue
        objetivo = _centavos(banco.at[bi, "valor"])
        fecha_b = banco.at[bi, "fecha"]
        descripcion = banco.at[bi, "descripcion"]

        amplios = [i for i in _ventana(libro_por_fecha, fecha_b, tolerancia_dias_nombre)
                   if not libro.at[i, "matched"] and (libro.at[i, "valor"] >= 0) == (banco.at[bi, "valor"] >= 0)
                   and _similitud_nombre(descripcion, libro.at[i, "beneficiario"]) >= 1
                   and _validar_nombre(descripcion, libro.at[i, "beneficiario"])]
        candidatos = amplios
        if not candidatos:
            estrictos = [i for i in _ventana(libro_por_fecha, fecha_b, tolerancia_dias)
                         if not libro.at[i, "matched"] and (libro.at[i, "valor"] >= 0) == (banco.at[bi, "valor"] >= 0)]
            candidatos = estrictos if len(estrictos) <= max_candidatos_sin_nombre else []

        valores = [(i, _centavos(libro.at[i, "valor"])) for i in candidatos]
        combo = _buscar_subconjunto(valores, objetivo, max_grupo)
        if combo:
            for i in combo:
                libro.at[i, "matched"] = True
            banco.at[bi, "matched"] = True
            grupos.append({"tipo": "libro->banco", "banco_ids": [bi], "libro_ids": combo})

    return grupos


def _agrupar_por_concepto(banco, libro, tolerancia_dias=3, max_combinatoria=14):
    """Cruza los casos en que el banco itemiza y la contabilidad consolida (o al revés):
    todos los movimientos del mismo concepto (nómina, 4x1000, IVA, comisiones...) dentro
    de una ventana de fechas se suman y se comparan contra el asiento único del otro lado.

    Primero intenta la suma de TODOS los movimientos del concepto (el caso real de un pago
    de nómina en lote); si no cuadra y son pocos, busca un subconjunto exacto."""
    grupos = []

    banco_por_fecha = defaultdict(list)
    for i in banco.index:
        banco_por_fecha[banco.at[i, "fecha"]].append(i)
    libro_por_fecha = defaultdict(list)
    for i in libro.index:
        libro_por_fecha[libro.at[i, "fecha"]].append(i)

    def _ventana(bucket, fecha_centro, dias):
        out = []
        for dd in range(-dias, dias + 1):
            out.extend(bucket.get(fecha_centro + timedelta(days=dd), []))
        return out

    def _resolver(objetivo, candidatos, valores_cent, fechas):
        """Devuelve la lista de índices que suma exactamente el objetivo, o None.
        Prueba tres estrategias, de la más simple a la más costosa."""
        if len(candidatos) < 2:
            return None

        # 1) Todos los movimientos del concepto suman el asiento (lote completo).
        if sum(valores_cent) == objetivo:
            return list(candidatos)

        # 2) Subconjunto de días COMPLETOS: la contabilidad consolida uno o varios lotes
        #    diarios (ej. la nómina del 13 y la del 16 en un solo asiento del 15). Se
        #    buscan combinaciones de fechas, no de movimientos sueltos: son pocas y es
        #    además como lo razona una persona al conciliar.
        por_dia = defaultdict(list)
        for idx, cent, f in zip(candidatos, valores_cent, fechas):
            por_dia[f].append((idx, cent))
        if 2 <= len(por_dia) <= 18:
            dias = sorted(por_dia)
            totales = [(d, sum(c for _, c in por_dia[d])) for d in dias]
            for tam in range(1, min(len(totales), 6) + 1):
                for combo in combinations(range(len(totales)), tam):
                    if sum(totales[i][1] for i in combo) == objetivo:
                        elegidos = []
                        for i in combo:
                            elegidos.extend(idx for idx, _ in por_dia[totales[i][0]])
                        if len(elegidos) >= 2:
                            return elegidos

        # 3) Subconjunto de movimientos sueltos (solo si son pocos, es combinatorio).
        if len(candidatos) <= max_combinatoria:
            return _buscar_subconjunto(list(zip(candidatos, valores_cent)), objetivo,
                                        max_tam=min(len(candidatos), 8))
        return None

    # A) varios movimientos del banco (mismo concepto) -> un asiento contable
    for li in [i for i in libro.index if not libro.at[i, "matched"]]:
        if libro.at[li, "matched"]:
            continue
        concepto = _concepto(libro.at[li, "descripcion"])
        if not concepto:
            continue
        objetivo = _centavos(libro.at[li, "valor"])
        signo = libro.at[li, "valor"] >= 0
        cands = [i for i in _ventana(banco_por_fecha, libro.at[li, "fecha"], tolerancia_dias)
                 if not banco.at[i, "matched"] and (banco.at[i, "valor"] >= 0) == signo
                 and _concepto(banco.at[i, "descripcion"]) == concepto]
        combo = _resolver(objetivo, cands, [_centavos(banco.at[i, "valor"]) for i in cands],
                           [banco.at[i, "fecha"] for i in cands])
        if combo:
            for i in combo:
                banco.at[i, "matched"] = True
            libro.at[li, "matched"] = True
            grupos.append({"tipo": "banco->libro", "banco_ids": combo, "libro_ids": [li],
                            "concepto": concepto})

    # B) varios asientos contables (mismo concepto) -> un movimiento del banco
    for bi in [i for i in banco.index if not banco.at[i, "matched"]]:
        if banco.at[bi, "matched"]:
            continue
        concepto = _concepto(banco.at[bi, "descripcion"])
        if not concepto:
            continue
        objetivo = _centavos(banco.at[bi, "valor"])
        signo = banco.at[bi, "valor"] >= 0
        cands = [i for i in _ventana(libro_por_fecha, banco.at[bi, "fecha"], tolerancia_dias)
                 if not libro.at[i, "matched"] and (libro.at[i, "valor"] >= 0) == signo
                 and _concepto(libro.at[i, "descripcion"]) == concepto]
        combo = _resolver(objetivo, cands, [_centavos(libro.at[i, "valor"]) for i in cands],
                           [libro.at[i, "fecha"] for i in cands])
        if combo:
            for i in combo:
                libro.at[i, "matched"] = True
            banco.at[bi, "matched"] = True
            grupos.append({"tipo": "libro->banco", "banco_ids": [bi], "libro_ids": combo,
                            "concepto": concepto})

    return grupos


def _agrupar_lotes_por_dia(banco, libro, max_dias_combinados=4, max_dias_distancia=20):
    """Cruza los lotes que el banco paga por días y la contabilidad registra en un asiento
    global, incluso cuando la descripción contable no dice de qué se trata.

    Ej. real: 'BANCOS - ISTHO SAS' del 15/06 por 115.107.631 = toda la nómina que el banco
    pagó el 13/06 (111.580.741) más la del 16/06 (3.526.890).

    Trabaja al revés que `_agrupar_por_concepto`: el concepto se toma del banco (que sí lo
    describe) y del lado contable solo se exige que el valor calce exacto y la fecha esté
    cerca. Para que sea rápido, precalcula de una vez las sumas de todas las combinaciones
    de hasta `max_dias_combinados` días por concepto y luego consulta por valor."""
    grupos = []

    por_concepto = defaultdict(lambda: defaultdict(list))
    for i in banco.index:
        if banco.at[i, "matched"]:
            continue
        c = _concepto(banco.at[i, "descripcion"])
        if c:
            por_concepto[c][banco.at[i, "fecha"]].append(i)

    for concepto, dias_dict in por_concepto.items():
        dias = sorted(dias_dict)
        if len(dias) < 1:
            continue
        totales = [(d, sum(_centavos(banco.at[i, "valor"]) for i in dias_dict[d])) for d in dias]

        # Índice suma -> combinación de días (la primera que aparezca).
        indice = {}
        for tam in range(1, min(len(totales), max_dias_combinados) + 1):
            for combo in combinations(range(len(totales)), tam):
                indice.setdefault(sum(totales[i][1] for i in combo), combo)

        for li in libro.index:
            if libro.at[li, "matched"]:
                continue
            combo = indice.get(_centavos(libro.at[li, "valor"]))
            if not combo:
                continue

            ids, fechas_combo = [], []
            for i in combo:
                dia = totales[i][0]
                ids.extend(dias_dict[dia])
                fechas_combo.append(dia)
            # Los movimientos pueden haber sido tomados por un cruce anterior.
            if len(ids) < 2 or any(banco.at[x, "matched"] for x in ids):
                continue
            # La fecha del asiento debe estar cerca del lote, si no es coincidencia.
            fecha_l = libro.at[li, "fecha"]
            if min(abs((fecha_l - f).days) for f in fechas_combo) > max_dias_distancia:
                continue

            for x in ids:
                banco.at[x, "matched"] = True
            libro.at[li, "matched"] = True
            grupos.append({"tipo": "banco->libro", "banco_ids": ids, "libro_ids": [li],
                            "concepto": f"{concepto}, lote de {len(fechas_combo)} día(s)"})

    return grupos


def _cruces_posibles_por_margen(banco, libro, margen_valor=100000, tolerancia_dias=3, min_similitud=2):
    """Para lo que sigue sin cruzar: busca pares con la MISMA fecha (o dentro de la
    tolerancia) y coincidencia de nombre/beneficiario, cuyo valor difiera hasta
    `margen_valor` (no exacto, si no ya habría cruzado antes). Son señales fuertes
    de que es el mismo movimiento pero con un valor mal digitado o un descuento/
    comisión de por medio: se listan aparte para que una persona los confirme."""
    posibles = []
    libro_por_fecha = defaultdict(list)
    for i in libro.index:
        if not libro.at[i, "matched"]:
            libro_por_fecha[libro.at[i, "fecha"]].append(i)

    idxs_banco = [i for i in banco.index if not banco.at[i, "matched"]]
    for bi in idxs_banco:
        if banco.at[bi, "matched"]:
            continue
        fecha_b = banco.at[bi, "fecha"]
        candidatos = []
        for dd in range(-tolerancia_dias, tolerancia_dias + 1):
            candidatos.extend(libro_por_fecha.get(fecha_b + timedelta(days=dd), []))

        mejor, mejor_score = None, None
        for li in candidatos:
            if libro.at[li, "matched"]:
                continue
            if (banco.at[bi, "valor"] >= 0) != (libro.at[li, "valor"] >= 0):
                continue
            if not _validar_nombre(banco.at[bi, "descripcion"], libro.at[li, "beneficiario"]):
                continue
            sim = _similitud_nombre(banco.at[bi, "descripcion"], libro.at[li, "beneficiario"])
            if sim < min_similitud:
                continue
            diff = abs(banco.at[bi, "valor"] - libro.at[li, "valor"])
            if diff > margen_valor or diff == 0:
                continue
            score = (-sim, diff)
            if mejor is None or score < mejor_score:
                mejor, mejor_score = li, score

        if mejor is not None:
            banco.at[bi, "matched"] = True
            libro.at[mejor, "matched"] = True
            posibles.append({"banco_id": bi, "libro_id": mejor, "diferencia": banco.at[bi, "valor"] - libro.at[mejor, "valor"]})

    return posibles


def reconciliar(df_banco, df_libro, tolerancia_dias=3, tolerancia_dias_nombre=15, agrupar_por_fecha=True,
                 max_grupo=6, buscar_posibles=True, margen_valor=100000):
    """Cruza por valor exacto y fecha (exacta primero, luego dentro de la tolerancia); si con el valor
    exacto el nombre del beneficiario también coincide, se acepta una ventana de fecha más amplia
    (`tolerancia_dias_nombre`). Luego agrupa (varios movimientos de un lado, en la misma fecha, que
    suman exactamente el valor de uno o varios del otro) y, por último, marca como 'posibles' los
    pares que coinciden en fecha y nombre pero cuyo valor difiere hasta `margen_valor`."""
    banco = df_banco.copy()
    libro = df_libro.copy()
    banco["matched"] = False
    libro["matched"] = False

    libro_by_valor = defaultdict(list)
    for idx, row in libro.iterrows():
        libro_by_valor[round(row["valor"], 2)].append(idx)

    matches = []

    # Pasada 1: fecha exacta + valor exacto (siempre validando primero el nombre)
    for idx, row in banco.iterrows():
        val = round(row["valor"], 2)
        candidatos = [i for i in libro_by_valor.get(val, []) if not libro.at[i, "matched"]
                       and _validar_nombre(row["descripcion"], libro.at[i, "beneficiario"])]
        exactos = [i for i in candidatos if libro.at[i, "fecha"] == row["fecha"]]
        if exactos:
            match_idx = exactos[0]
            libro.at[match_idx, "matched"] = True
            banco.at[idx, "matched"] = True
            matches.append({"banco_id": idx, "libro_id": match_idx, "dias_diferencia": 0,
                             "confianza": "Alta (fecha y valor exactos)"})

    # Pasada 2: valor exacto + fecha dentro de la tolerancia (la más cercana gana)
    if tolerancia_dias > 0:
        for idx, row in banco.iterrows():
            if banco.at[idx, "matched"]:
                continue
            val = round(row["valor"], 2)
            candidatos = [i for i in libro_by_valor.get(val, []) if not libro.at[i, "matched"]
                           and _validar_nombre(row["descripcion"], libro.at[i, "beneficiario"])]
            mejor, mejor_diff = None, None
            for i in candidatos:
                diff = abs((row["fecha"] - libro.at[i, "fecha"]).days)
                if diff <= tolerancia_dias and (mejor is None or diff < mejor_diff):
                    mejor, mejor_diff = i, diff
            if mejor is not None:
                libro.at[mejor, "matched"] = True
                banco.at[idx, "matched"] = True
                matches.append({"banco_id": idx, "libro_id": mejor, "dias_diferencia": mejor_diff,
                                 "confianza": f"Media (valor exacto, {mejor_diff} día(s) de diferencia)"})

    # Pasada 2b: valor exacto + nombre coincide, con una ventana de fecha más amplia
    # (ej. el banco contabiliza el giro días después de la liquidación contable).
    if tolerancia_dias_nombre > tolerancia_dias:
        for idx, row in banco.iterrows():
            if banco.at[idx, "matched"]:
                continue
            val = round(row["valor"], 2)
            candidatos = [i for i in libro_by_valor.get(val, []) if not libro.at[i, "matched"]]
            mejor, mejor_score = None, None
            for i in candidatos:
                diff = abs((row["fecha"] - libro.at[i, "fecha"]).days)
                if diff > tolerancia_dias_nombre:
                    continue
                if not _validar_nombre(row["descripcion"], libro.at[i, "beneficiario"]):
                    continue
                sim = _similitud_nombre(row["descripcion"], libro.at[i, "beneficiario"])
                if sim < 1:
                    continue
                score = (-sim, diff)
                if mejor is None or score < mejor_score:
                    mejor, mejor_score = i, score
            if mejor is not None:
                dias = abs((row["fecha"] - libro.at[mejor, "fecha"]).days)
                libro.at[mejor, "matched"] = True
                banco.at[idx, "matched"] = True
                matches.append({"banco_id": idx, "libro_id": mejor, "dias_diferencia": dias,
                                 "confianza": f"Alta (valor exacto y nombre coincide, {dias} día(s) de diferencia)"})

    # Pasada 3: agrupación por fecha (varios movimientos que suman el mismo valor)
    grupos = []
    if agrupar_por_fecha:
        # Para la agrupación (sumas) se limita la ventana ampliada por nombre: es una búsqueda
        # cuadrática y una ventana de 15+ días compara cada movimiento contra medio mes de
        # candidatos. Los casos reales de "varios movimientos = 1 solo" caen a pocos días.
        ventana_grupo = min(tolerancia_dias_nombre, 6)
        grupos = _agrupar_por_fecha(banco, libro, tolerancia_dias=tolerancia_dias,
                                     tolerancia_dias_nombre=ventana_grupo, max_grupo=max_grupo)

        # Pasada 3b: lo que el banco itemiza y la contabilidad consolida (nómina, 4x1000,
        # IVA, comisiones...). Va después para que los cruces por nombre tengan prioridad.
        grupos += _agrupar_por_concepto(banco, libro, tolerancia_dias=ventana_grupo)

        # Pasada 3c: lotes completos de días (el concepto se lee del banco, no del asiento).
        grupos += _agrupar_lotes_por_dia(banco, libro)

    # Pasada 4: coinciden en fecha y nombre, pero el valor difiere hasta el margen permitido
    posibles = []
    if buscar_posibles and margen_valor > 0:
        posibles = _cruces_posibles_por_margen(banco, libro, margen_valor=margen_valor,
                                                tolerancia_dias=tolerancia_dias, min_similitud=2)

    # --- Se arma la lista de cruces: cada uno es una conciliación con identidad propia,
    # --- lo que permite deshacerla después sin tocar los datos originales.
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    cruces = []
    for m in matches:
        cruces.append({
            "origen": "Automática",
            "motivo": m["confianza"],
            "fecha_hora": ahora,
            "banco_ids": [int(m["banco_id"])],
            "libro_ids": [int(m["libro_id"])],
        })
    for g in grupos:
        concepto = g.get("concepto")
        prefijo = f"Agrupado por concepto «{concepto}»" if concepto else "Agrupado por fecha"
        if g["tipo"] == "banco->libro":
            motivo = f"{prefijo} ({len(g['banco_ids'])} mov. banco = 1 mov. contabilidad)"
        else:
            motivo = f"{prefijo} ({len(g['libro_ids'])} mov. contabilidad = 1 mov. banco)"
        cruces.append({
            "origen": "Automática",
            "motivo": motivo,
            "fecha_hora": ahora,
            "banco_ids": [int(i) for i in g["banco_ids"]],
            "libro_ids": [int(i) for i in g["libro_ids"]],
        })

    for n, c in enumerate(cruces, start=1):
        c["id"] = f"A-{n:04d}"

    posibles = [{"banco_id": int(p["banco_id"]), "libro_id": int(p["libro_id"]),
                 "diferencia": float(p["diferencia"])} for p in posibles]

    return cruces, posibles


# --------------------------------------------------------------------------------
# Estado de la conciliación
#
# La verdad del sistema es la lista de `cruces` (cada uno con su id, origen, fecha y
# los movimientos que relaciona). Todas las tablas que se muestran se derivan de ahí,
# así que crear o deshacer una conciliación nunca modifica ni elimina los datos
# originales: solo cambia esa lista.
# --------------------------------------------------------------------------------
def siguiente_id_manual(cruces):
    usados = [c["id"] for c in cruces if str(c.get("id", "")).startswith("M-")]
    return f"M-{len(usados) + 1:04d}"


def crear_cruce_manual(cruces, banco_ids, libro_ids):
    """Agrega una conciliación manual a la lista y devuelve la nueva lista (no muta la original)."""
    nuevo = {
        "id": siguiente_id_manual(cruces),
        "origen": "Manual",
        "motivo": f"Conciliación manual ({len(banco_ids)} mov. banco ↔ {len(libro_ids)} mov. contabilidad)",
        "fecha_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "banco_ids": [int(i) for i in banco_ids],
        "libro_ids": [int(i) for i in libro_ids],
    }
    return cruces + [nuevo], nuevo["id"]


def eliminar_cruces(cruces, ids):
    """Deshace las conciliaciones indicadas. Los movimientos vuelven solos a pendientes
    porque las vistas se derivan de esta lista."""
    ids = set(ids)
    return [c for c in cruces if c["id"] not in ids]


def _fila_cruce(c, b, l, valor):
    return {
        "ID": c["id"],
        "Origen": c["origen"],
        "Fecha Banco": b["fecha"] if b is not None else None,
        "Fecha Contabilidad": l["fecha"] if l is not None else None,
        "Dif. días": (abs((b["fecha"] - l["fecha"]).days) if b is not None and l is not None else 0),
        "Valor": valor,
        "Descripción Banco": b["descripcion"] if b is not None else "",
        "Descripción Contabilidad": l["descripcion"] if l is not None else "",
        "Comprobante": l["comprobante"] if l is not None else "",
        "Documento": l["documento"] if l is not None else "",
        "Motivo": c["motivo"],
        "Conciliado el": c["fecha_hora"],
    }


def construir_vistas(df_banco, df_libro, cruces, posibles):
    """Deriva todas las tablas que ve el usuario a partir de la lista de cruces."""
    banco_usados, libro_usados = set(), set()
    for c in cruces:
        banco_usados.update(c["banco_ids"])
        libro_usados.update(c["libro_ids"])
    for p in posibles:
        banco_usados.add(p["banco_id"])
        libro_usados.add(p["libro_id"])

    filas = []
    for c in cruces:
        bs, ls = c["banco_ids"], c["libro_ids"]
        if len(ls) == 1 and len(bs) >= 1:
            l = df_libro.loc[ls[0]]
            for bi in bs:
                b = df_banco.loc[bi]
                filas.append(_fila_cruce(c, b, l, b["valor"]))
        elif len(bs) == 1:
            b = df_banco.loc[bs[0]]
            for li in ls:
                l = df_libro.loc[li]
                filas.append(_fila_cruce(c, b, l, l["valor"]))
        else:  # varios contra varios (solo puede venir de una conciliación manual)
            for bi in bs:
                b = df_banco.loc[bi]
                filas.append(_fila_cruce(c, b, None, b["valor"]))
            for li in ls:
                l = df_libro.loc[li]
                filas.append(_fila_cruce(c, None, l, l["valor"]))
    df_conciliados = pd.DataFrame(filas)

    filas_pos = []
    for p in posibles:
        b, l = df_banco.loc[p["banco_id"]], df_libro.loc[p["libro_id"]]
        filas_pos.append({
            "Fecha Banco": b["fecha"], "Fecha Contabilidad": l["fecha"],
            "Valor Banco": b["valor"], "Valor Contabilidad": l["valor"],
            "Diferencia": p["diferencia"],
            "Descripción Banco": b["descripcion"], "Descripción Contabilidad": l["descripcion"],
            "Comprobante": l["comprobante"], "Documento": l["documento"],
        })
    df_posibles = pd.DataFrame(filas_pos)

    df_solo_banco = df_banco[~df_banco["id"].isin(banco_usados)].reset_index(drop=True)
    df_solo_libro = df_libro[~df_libro["id"].isin(libro_usados)].reset_index(drop=True)
    return df_conciliados, df_posibles, df_solo_banco, df_solo_libro


def resumen_cruces(df_banco, df_libro, cruces):
    """Una fila por conciliación, para la pantalla donde se deshacen. Incluye las
    descripciones de ambos lados porque el cruce se valida principalmente por el nombre
    del beneficiario, y sin verlo aquí no hay forma de confirmar que la conciliación
    seleccionada es la correcta antes de deshacerla."""
    filas = []
    for c in cruces:
        v_banco = sum(float(df_banco.loc[i, "valor"]) for i in c["banco_ids"])
        v_libro = sum(float(df_libro.loc[i, "valor"]) for i in c["libro_ids"])
        desc_banco = "; ".join(str(df_banco.loc[i, "descripcion"]) for i in c["banco_ids"])
        desc_libro = "; ".join(str(df_libro.loc[i, "descripcion"]) for i in c["libro_ids"])
        filas.append({
            "ID": c["id"],
            "Origen": c["origen"],
            "Descripción Banco": desc_banco,
            "Descripción Contabilidad": desc_libro,
            "Mov. banco": len(c["banco_ids"]),
            "Mov. contabilidad": len(c["libro_ids"]),
            "Valor banco": v_banco,
            "Valor contabilidad": v_libro,
            "Motivo": c["motivo"],
            "Conciliado el": c["fecha_hora"],
        })
    return pd.DataFrame(filas)

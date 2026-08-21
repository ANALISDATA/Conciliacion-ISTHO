# Conciliación Bancaria — ISTHO S.A.S.

Aplicación local que cruza el **extracto bancario** de Bancolombia contra el **libro auxiliar
contable**, encuentra las diferencias y genera los informes en Excel con el membrete de la
empresa. Reemplaza un proceso que se hacía a mano.

> Este documento es la referencia completa del proyecto: qué hace, cómo está construido, por
> qué se tomó cada decisión y qué trampas ya se encontraron. Si hay que retomar el desarrollo
> después de un tiempo, leer esto alcanza para ponerse al día.

---

## 1. Cómo se usa

1. Doble clic en **`Iniciar_App.bat`**. Se abre una ventana negra (déjala abierta, es el
   servidor) y el navegador en `http://localhost:8501`.
2. En el panel izquierdo:
   - **① Archivos** → cargar el extracto bancario y el libro auxiliar (`.xlsx` o `.csv`).
   - **② Parámetros de cruce** → tolerancias y opciones (los valores por defecto funcionan bien).
   - **③ Saldo inicial** → escribir el «SALDO ANTERIOR» que aparece en el extracto.
3. Botón **Conciliar**. Tarda entre 15 y 35 segundos con un mes completo; mientras tanto se
   ve una pantalla de carga animada con el avance por etapas.
4. Navegar entre las 5 hojas con la barra superior y descargar los Excel que se necesiten.

> ⚠️ **Al cambiar el código hay que cerrar la consola y volver a abrir el `.bat`.** Refrescar
> el navegador solo recarga `app.py`; los módulos importados (`conciliacion.py`, `ui.py`,
> `excel_export.py`) quedan en memoria con la versión anterior.

---

## 2. Archivos del proyecto

| Archivo | Responsabilidad |
|---|---|
| `app.py` | Interfaz: sidebar, navegación, las 5 hojas, filtros y botones. No contiene lógica de cruce. |
| `conciliacion.py` | **El motor.** Lectura de archivos, algoritmo de cruce y modelo de estado. Es el corazón del proyecto. |
| `ui.py` | Tema visual: CSS, encabezado, tarjetas de indicadores, pantalla de carga y el componente de tabla. |
| `excel_export.py` | Generación de los Excel con membrete, formato contable y fila de totales. |
| `logo_istho.png` | Logo que se inserta en la app y en los Excel. |
| `Iniciar_App.bat` | Lanzador. |
| `.streamlit/config.toml` | Colores base del tema y desactivación de la telemetría de Streamlit. |

**Dependencias:** `streamlit`, `pandas`, `openpyxl`, `xlsxwriter`.
(`playwright` se usó solo para verificar la interfaz durante el desarrollo, no hace falta para correr la app.)

---

## 3. Formato de los archivos de entrada

### Extracto bancario (Bancolombia)
**No tiene fila de encabezados.** Se leen por posición:

| Columna | Contenido |
|---|---|
| 3 | Fecha |
| 5 | Valor — **con signo**: positivo = entrada, negativo = salida |
| 7 | Descripción |

⚠️ **El banco trunca las descripciones a ~28 caracteres.** `"PAGO A PROV Estructuras y Sold"`
es el texto completo que entrega Bancolombia, no un recorte de la app. Esto condiciona todo el
algoritmo de comparación de nombres.

### Libro auxiliar contable
Sí tiene encabezados. Se usan: `FECHA`, `DETALLE`, `NOMBRE BENEFICIARIO`, `DÉBITO`, `CRÉDITO`,
`COMPROBANTE`, `DOCUMENTO`, `NOMBRE AUXILIAR`.

**El valor comparable se calcula como `DÉBITO − CRÉDITO`**, para que quede en la misma
convención de signos del extracto (débito = entra dinero al banco, crédito = sale).

Los importes vienen como texto (`"1,137,260.00"`); `parse_money()` maneja formato con coma o
con punto decimal, paréntesis para negativos y símbolos de moneda.

---

## 4. El motor de cruce

`reconciliar()` ejecuta varias pasadas **de la más estricta a la más flexible**. Cada
movimiento que cruza queda marcado y ya no se considera en las pasadas siguientes, así que
nunca se usa dos veces.

### Validación previa obligatoria: el nombre

Antes de cualquier cruce automático corre `_validar_nombre(descripcion_banco, beneficiario_libro)`:

- Separa en palabras, pasa a mayúsculas y descarta conectores y tecnicismos (`_STOPWORDS`:
  PAGO, PROVE, TRANSFERENCIA, ANTICIPO, etc.).
- **No importa el orden**: se comparan como conjuntos → `"Juan Carlos Marín"` y
  `"Marín Juan Carlos"` son la misma persona.
- **Todas las palabras del nombre más corto deben estar en el más largo** → `"Juan Carlos
  Restrepo"` y `"Juan Carlos Marín"` NO cruzan: sobra un apellido distinto.
- **Tolera el truncamiento del banco** aceptando que una palabra sea prefijo de la otra
  (mínimo 3 letras): `"TOBON EDWIN ALB"` ≡ `"TOBON JARAMILLO EDWIN ALBERTO"`.
- Con una sola palabra en común solo acepta si tiene ≥5 letras (una razón social como
  `"BANCOOMEVA"`), nunca un nombre de pila suelto como `"JUAN"`.
- **Si en alguno de los dos lados no hay nombre identificable** (ej. el extracto dice solo
  `"TRANSFERENCIA CTA SUC VIRTUAL"`), no hay nada que validar y deja pasar: deciden el valor
  y la fecha.

Dos correcciones que se agregaron después de medir cuántos cruces exactos estaba bloqueando
de más:

- **`ISTHO` está en las stopwords.** La contabilidad pone a la propia empresa como
  beneficiario en retiros, traslados y movimientos internos; eso no identifica a ninguna
  contraparte. Comparar contra ese nombre bloqueaba cruces con fecha y valor exactos
  (ej. `"RETIRO CAJERO AUTOSERVICIOS"` contra un asiento a nombre de ISTHO SAS).
- **Reintento por truncamiento.** Si la validación falla y la descripción del extracto mide
  ≥26 caracteres (o sea, viene cortada), se reintenta **descartando su última palabra**, que
  suele estar partida a la mitad: `"PAGO A PROVE comfenalco cred"` es en realidad
  `"COMFENALCO CREDITOS"` y corresponde a `"CAJA DE COMPENSACION FAMILIAR COMFENALCO
  ANTIOQUIA"`. Sin esto, la palabra rota ensuciaba la comparación y descartaba la entidad
  correcta.

Los casos que deben seguir **sin** cruzar en las pasadas 1-3 siguen bloqueados:
`"Jorge Luis Garc"` contra `"CAMELO BELTRAN JHON SEBASTIAN"` no comparte ninguna palabra,
aunque el valor y la fecha coincidan por casualidad. **Esta pareja sí termina cruzando en la
segunda ronda** (pasada 4, sección siguiente), que ya no valida nombre — a propósito, y por
eso queda marcada Baja y es fácil de deshacer si al revisar resulta ser la coincidencia
equivocada.

> Se exige **subconjunto** y no igualdad exacta de palabras justamente porque el banco trunca.
> Con igualdad estricta se perderían casi todos los cruces reales.

### Pasadas, en orden

| # | Función | Qué busca |
|---|---|---|
| 1 | (en `reconciliar`) | Fecha exacta + valor exacto |
| 2 | (en `reconciliar`) | Valor exacto + fecha dentro de `tolerancia_dias` (3 por defecto) |
| 2b | (en `reconciliar`) | Valor exacto + **nombre coincide** → acepta ventana más amplia, `tolerancia_dias_nombre` (15 días). Resuelve el caso del banco que gira días después de la liquidación contable. |
| 3 | `_agrupar_por_fecha` | Varios movimientos de un lado que **suman** el valor de uno del otro, acotando candidatos por nombre. Ej: un anticipo y su «pago sobre anticipo» = 1 solo giro bancario. |
| 3b | `_agrupar_por_concepto` | Cuando el banco **itemiza** y la contabilidad **consolida**. Reconoce el concepto (Nómina, Seguridad social, 4x1000, IVA, Intereses, Retenciones, Comisiones) y suma todos los del mismo tipo. |
| 3c | `_agrupar_lotes_por_dia` | **Lotes de días completos.** El concepto se lee del lado del banco (que sí lo describe) y del contable solo se exige que el valor calce y la fecha esté cerca. |
| 4 | `_cruces_posibles_por_margen` | Coinciden en fecha y nombre pero el valor difiere hasta `margen_valor` ($100.000). **No se dan por conciliados**: quedan aparte para revisión manual. |

Aparte de estas 4 (todas dentro de `reconciliar()`, disparadas por el botón **Conciliar**), existe
una pasada más que **no** corre automáticamente — ver la sección siguiente.

**Ejemplos reales que resuelven las pasadas de agrupación:**

- *3b*: `1.573.547 + 1.117.134 + 1.035.400 = 3.726.081` — tres pagos de nómina individuales en
  el banco contra un solo asiento «PAGO NOMINA 31 MAYO 2026».
- *3c*: el asiento «BANCOS - ISTHO SAS» del 15/06 por `115.107.631` es toda la nómina que el
  banco pagó el 13/06 (`111.580.741`) más la del 16/06 (`3.526.890`) — 96 movimientos en un
  solo cruce. La descripción contable no dice «nómina» por ningún lado, por eso esta pasada
  lee el concepto del banco.

### Segunda ronda: cruzar lo pendiente solo por valor (acción aparte)

Después de las pasadas 1-4 (todas exigen que el nombre valide, salvo la de margen que también
lo exige), muchos movimientos quedaban pendientes con **el mismo valor en los dos lados** y lo
único que fallaba era el nombre: el extracto viene cortado a 28 caracteres, describe el canal
(`"TRANSFERENCIA CTA SUC VIRTUAL"`) en vez de la contraparte, o la contabilidad registró el
asiento a nombre de otro tercero. Medido en Junio 2026: **75 cruces adicionales**, pendientes
de banco 615→540 y de contabilidad 151→76 (en Mayo 2026: 76 cruces, 443→367 y 156→80).

**No corre dentro de `reconciliar()` ni con el botón Conciliar, a propósito.** Es
`aplicar_segunda_ronda(df_banco, df_libro, cruces, posibles)`, una función y un botón aparte
(«🔁 Cruzar pendientes por valor», debajo de Conciliar en el sidebar) para que quien concilia
pueda **revisar lo que quedó pendiente antes** de decidir si vale la pena aplicar un criterio
menos estricto — no es automático porque no es igual de confiable que el resto.

- Se habilita solo cuando ya hay una conciliación cargada (recién hecha o retomada).
- No repite las pasadas de `reconciliar()` ni toca los cruces existentes: calcula qué sigue
  pendiente (lo que ningún cruce ni ningún «posible» está usando) y le corre
  `_cruces_solo_valor()` encima, agregando sus resultados a la lista de cruces.
- **Valor exacto al centavo y mismo signo**, nunca con margen ni sumando varios movimientos.
- **Sin ningún límite de fecha**: no importa si un lado quedó registrado un mes distinto al
  del otro, si el valor calza, cruza. La fecha solo interviene para decidir el orden cuando
  hay más de un candidato con el mismo valor — se arman todos los pares posibles y se
  resuelven de fechas más parecidas a más lejanas, para que un par casi exacto en fecha nunca
  pierda su candidato ideal frente a uno que se resolvió antes por el orden de recorrido.
- Sus cruces quedan con motivo **«Baja (solo valor exacto, sin validar nombre...)»**, para
  distinguirlos a simple vista en la hoja Conciliados y poder deshacerlos si al revisar
  resultan ser coincidencia y no el mismo movimiento.
- Es **idempotente**: darle clic dos veces no duplica nada, porque la segunda vez ya no
  encuentra pendientes nuevos que cruzar.

**Es una decisión de diseño deliberada, no un descuido**: el caso que la sección anterior
documenta como "no debe cruzar" —`"Jorge Luis Garc"` contra `"CAMELO BELTRAN JHON
SEBASTIAN"`— sí termina cruzando aquí, porque el valor y la fecha sí son idénticos. Se acepta
el riesgo de alguna coincidencia porque, revisado contra los datos reales, la enorme mayoría
de lo que cruza esta pasada son pares genuinos que la validación de nombre no pudo confirmar
(nombre truncado, descripción del canal, tercero distinto en la contabilidad), no coincidencias.

### Notas de rendimiento
- Las búsquedas de subconjuntos son combinatorias. Hay topes deliberados
  (`max_grupo`, `max_combinatoria`, ventanas de días) para que el proceso no se dispare.
- La ventana de la agrupación se limita a 6 días aunque `tolerancia_dias_nombre` sea mayor:
  con 15 días el proceso pasaba de ~33 s a ~87 s sin encontrar prácticamente más cruces.
- Los importes se comparan **en centavos enteros** (`_centavos`), nunca con flotantes.
- Las funciones de nombre (`_validar_nombre`, `_similitud_nombre`, `_mismo_nombre`) usan
  `lru_cache` de **400.000 entradas**: con ~1.000 movimientos contra ~500 asientos se
  consultan cientos de miles de pares y una caché chica se desaloja sola. Bajarla vuelve a
  costar ~15 s por corrida.
- Tiempo actual: **~50 s** por mes completo.

---

## 5. Modelo de estado: conciliar y desconciliar

Esta es la decisión de diseño más importante del proyecto.

`reconciliar()` **no devuelve tablas**, devuelve dos listas:

```python
cruces = [
  {
    "id": "A-0001",              # A- automático, M- manual
    "origen": "Automática",      # o "Manual"
    "motivo": "Alta (fecha y valor exactos)",
    "fecha_hora": "31/07/2026 13:46",
    "banco_ids": [17],           # índices de df_banco
    "libro_ids": [473],          # índices de df_libro
  },
  ...
]
posibles = [{"banco_id": .., "libro_id": .., "diferencia": ..}, ...]
```

**Todas las tablas que ve el usuario se derivan de esa lista** con `construir_vistas()`:
lo conciliado son los movimientos referenciados por algún cruce; lo pendiente es todo lo demás.

Como consecuencia:
- **Conciliar a mano** = agregar un elemento a la lista (`crear_cruce_manual`).
- **Desconciliar** = quitar un elemento de la lista (`eliminar_cruces`). Los movimientos
  reaparecen solos en pendientes.
- **Nunca se modifican ni se eliminan los datos originales.** `df_banco` y `df_libro` se
  cargan una vez y quedan intactos en `st.session_state`.

### Prueba de integridad
Cada vez que se toca el motor hay que verificar tres cosas: que ningún movimiento esté en dos
cruces, que la suma de conciliados + pendientes cubra el total, y que el cuadre cierre:

```
diferencia = total_banco − total_contabilidad
esperado   = (suma solo_banco − suma solo_contabilidad) + suma diferencias de «posibles»
```

Con Junio 2026 el cuadre da **4.016.309,19** y **se mantiene idéntico** después de conciliar,
desconciliar y volver a conciliar a mano. Si ese número cambia, algo se rompió.

---

## 6. La interfaz

### Sidebar
Marca ISTHO arriba y tres pasos numerados: Archivos · Parámetros de cruce · Saldo inicial,
más el botón Conciliar.

### Barra superior y hojas
Franja azul con el título, y debajo la barra de navegación centrada (`st.segmented_control`).

| Hoja | Contenido |
|---|---|
| 📊 **Resumen** | KPIs financieros + resumen del cruce + línea de control del cuadre |
| 🔗 **Conciliados** | Tabla de cruces + panel para **desconciliar** (con confirmación) |
| 🔍 **Por revisar** | Los «posibles»: coinciden en fecha y nombre, difieren en valor |
| 🏦 **Pend. extracto** | Movimientos del extracto sin registro contable |
| 📘 **Pend. libro auxiliar** | Registros contables que no aparecen en el extracto |
| ⚖️ **Cruce manual** | Dos paneles con selección múltiple, totales en vivo y botón Cruzar |

Los dos tipos de pendiente tienen **hoja propia** (antes compartían una con sub-pestañas):
son las dos listas que más se trabajan a mano y cada una necesita la pantalla completa.

**El filtro de Pend. extracto/Pend. libro auxiliar y el de su panel correspondiente en Cruce
manual son el mismo** (comparten `key` en `filtrar()`, ver `app.py`): filtras un pendiente en
una hoja para identificarlo y, al pasar a Cruce manual a cruzarlo a mano, ya aparece filtrado
ahí también, sin repetir la búsqueda.

En las hojas de tablas el encabezado usa su versión **compacta** (una línea) para que los
registros arranquen lo más arriba posible.

### KPIs financieros (hoja Resumen)

```
Saldo final = Saldo inicial + Total ingresos − Total salidas
```

Ingresos y salidas se determinan **solo por el signo** del movimiento, sin mirar el concepto.

Validado contra el extracto real de Junio 2026 — **coincide al centavo**:

| KPI | App | Extracto |
|---|---|---|
| Saldo inicial | 133.174.991,10 | SALDO ANTERIOR |
| Total ingresos | 1.606.508.447,63 | TOTAL ABONOS |
| Total salidas | 1.538.074.252,49 | TOTAL CARGOS |
| **Saldo final** | **201.609.186,24** | **SALDO ACTUAL** |

### Cruce manual
Permite cualquier combinación (1↔1, 1↔N, N↔1, N↔M). Los totales de cada lado se recalculan al
seleccionar y **el botón Cruzar solo se habilita cuando ambos totales son exactamente iguales**;
si no, muestra la diferencia en rojo. Al cruzar se genera un ID `M-xxxx` con fecha y hora.

---

## 7. Exportación a Excel

Cada hoja tiene su propio botón **«Descargar esta tabla en Excel»**, que exporta **exactamente
lo que se está viendo, con el filtro aplicado**.

`build_tabla_workbook()` arma un archivo de una hoja con:
- Membrete: fila blanca con el logo, luego banda azul con título, empresa, NIT, cuenta,
  período, fecha de generación y tolerancia usada.
- Encabezados de tabla, filas alternadas, autofiltro y panel congelado.
- Formato contable en los importes y `dd/mm/aaaa` en las fechas.
- Fila de TOTAL al final.

Las columnas que se ocultan en pantalla (`Comprobante`, `Documento`, `Conciliado el`) **sí van
completas en el Excel**.

---

## 8. Tema visual

Paleta corporativa en **azul rey**, con colores semánticos consistentes:

| Constante (`ui.py`) | Color | Significado |
|---|---|---|
| `AZUL` | `#1D4ED8` | Marca y dato principal |
| `AZUL_OSC` | `#16307A` | Fondos y botones |
| `CIAN` | `#22D3EE` | Acento tecnológico del patrón |
| `VERDE` | `#0EA36B` | Positivo: ingresos, conciliados |
| `NARANJA` | `#E08700` | Requiere revisión |
| `ROJO` | `#DC2743` | Pendiente, salidas |
| `GRIS` | `#64748B` | Neutro |

> `VERDE_OSC` es un alias de `AZUL`. Se conserva el nombre para no tocar `app.py`.

Elementos propios: encabezado con patrón de circuito en SVG, tarjetas de indicadores con
tooltips en CSS puro, pestañas tipo píldora y pantalla de carga animada.

---

## 9. Trampas ya encontradas (leer antes de tocar nada)

Estas costaron tiempo. Están documentadas para no repetirlas.

**Pandas 3**
- `astype(str)` **no convierte los NaN a texto**; quedan como NaN y Streamlit los pinta como
  `"None"` en la tabla. Hay que convertir valor por valor con `pd.isna`.
- El texto ya no usa dtype `object`, así que `df[col].dtype == "object"` falla. Filtrar por
  `is_numeric_dtype` en su lugar.

**Streamlit / CSS**
- El encabezado es un **`<header>`**, no un `<div>`: el selector `div[data-testid="stHeader"]`
  no aplica. Además su barra de herramientas tapa la franja superior e impide pulsar la
  navegación → hay que quitarle `pointer-events` y devolvérselos solo a sus botones.
- Los contenedores de Streamlit son **flex en columna**: lo que centra en horizontal es
  `align-items`, no `justify-content`.
- Esta versión monta las pestañas con **react-aria**, no con baseweb: los selectores son
  `[role="tablist"]` y `[data-testid="stTab"]`.
- El HTML pasado a `st.markdown` **debe ir aplanado** (helper `_flat`): con indentación de 4+
  espacios se interpreta como bloque de código y se muestra el HTML crudo.
- Un `st.expander` se cierra en cada recarga, y la propia selección de filas dispara una
  recarga → para paneles con tablas seleccionables hay que usar `st.toggle`.
- **Los tooltips de las tarjetas necesitan elevar también a sus contenedores.** Un `z-index`
  alto en el globo no basta: cada bloque de Streamlit se pinta según su orden en el
  documento, así que el segundo grupo de tarjetas tapaba el globo del primero. Hay que subir
  con `:has(.istho-card:hover)` la tarjeta **y** los contenedores que la envuelven, y
  dejarlos con `overflow: visible`.
- No sirve verificar el apilamiento con `document.elementFromPoint`: el globo tiene
  `pointer-events: none`, así que ese método siempre lo atraviesa y reporta lo que hay
  detrás. Hay que comprobarlo visualmente.
- Las animaciones de la pantalla de carga son **CSS puro** a propósito: siguen corriendo
  mientras Python está bloqueado procesando.

**Datos**
- El extracto viene truncado a ~28 caracteres (ver sección 3). Condiciona toda la
  comparación de nombres.
- Las columnas de dinero deben ir anchas: si se recortan se pierde el **primer dígito** y el
  número deja de ser legible.

**PowerShell (entorno de desarrollo)**
- `Set-Content` re-codifica y **rompe los acentos** de los archivos `.py`. Usar la herramienta
  de escritura o `[System.IO.File]::WriteAllLines` con UTF-8.

---

## 10. Diferencias reales detectadas en Junio 2026

Estas **no se fuerzan a cruzar** porque no suman exacto. Son partidas conciliatorias legítimas
que alguien de contabilidad debe investigar:

| Asiento contable | Contabilidad | Suma en banco | Diferencia |
|---|---|---|---|
| PAGO NOMINA JUNIO DE 2026 (30/06) | 113.142.358 | 111.342.382 | **1.799.976** |
| PAGO DE PRIMAS SERVICIOS JUNIO (20/06) | 100.778.748 | 94.523.608 | **6.255.140** |
| GASTOS BANCARIOS JUNIO 2026 (30/06) | 8.144.829 | 7.592.572 | **552.257** |

Forzar estos cruces escondería una diferencia real.

---

## 11. Ideas pendientes

- Guardar el estado de la conciliación en disco para retomarla en otra sesión (hoy vive en
  memoria y se pierde al cerrar la app).
- Conciliar varios meses seguidos arrastrando las partidas pendientes del mes anterior.
- Permitir confirmar directamente un «posible» como conciliado desde su propia hoja.

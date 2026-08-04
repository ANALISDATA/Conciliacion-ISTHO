# Necesidad que resuelve la aplicación y beneficio esperado

**ISTHO S.A.S. — Área Financiera**
Conciliación bancaria: cuenta Bancolombia 255-000119-91

---

## 1. La situación anterior

La conciliación entre el **extracto bancario** y el **libro auxiliar contable** se hacía
**a mano**, comparando dos archivos de Excel línea por línea.

Cada mes había que revisar **cerca de 1.400 registros**:

| Mes medido | Movimientos del banco | Registros contables | Total a revisar |
|---|---:|---:|---:|
| Mayo 2026 | 832 | 507 | **1.339** |
| Junio 2026 | 1.006 | 488 | **1.494** |

**Tiempo que tomaba: entre 6 y 8 horas** por cada cierre mensual.

### Por qué era tan lento

El trabajo no consistía solo en buscar valores iguales. Había que resolver a ojo
situaciones como estas, todas reales de estos dos meses:

- **El banco detalla y la contabilidad agrupa.** Tres pagos de nómina del banco
  (1.573.547 + 1.117.134 + 1.035.400) corresponden a **un solo asiento** contable de
  3.726.081. Había que sumarlos mentalmente para descubrirlo.
- **Lotes de varios días.** Un asiento de 115.107.631 resultó ser toda la nómina que el
  banco pagó el 13 de junio más la del 16. Encontrar eso a mano es buscar una aguja en
  un pajar.
- **Los nombres no coinciden.** El extracto escribe `"PAGO A PROVE MONTOYA JULIAN"` y la
  contabilidad `"MONTOYA ARBOLEDA JULIAN ALONSO"`. Es la misma persona, en otro orden y
  con el nombre cortado.
- **Fechas desfasadas.** El banco registra el giro días después de que la contabilidad
  causó la liquidación.

### Riesgos de hacerlo así

- **Cansancio y error humano** tras horas comparando cifras.
- **Cruces equivocados** entre personas distintas que casualmente tienen el mismo valor.
- El trabajo **dependía de una sola persona** y de su método propio.
- **Sin rastro** de por qué se cruzó cada partida.

---

## 2. Lo que hace la aplicación

Cruza los dos archivos automáticamente validando **tres cosas a la vez**: el nombre, el
valor y la fecha. Reconoce los casos que antes había que descubrir a ojo:

- Suma varios movimientos del banco contra un único asiento contable.
- Identifica lotes completos de días (la nómina de varias fechas contra un asiento global).
- Compara nombres **sin importar el orden** y tolerando que el banco los corte.
- Separa aparte lo que coincide en fecha y nombre pero **difiere en el valor**, para que
  una persona lo revise en vez de darlo por bueno.

Lo que no logra cruzar queda organizado en listas de pendientes, listo para resolverlo a
mano dentro de la misma aplicación.

---

## 3. Resultados medidos

Cifras reales obtenidas al procesar los archivos de mayo y junio de 2026:

| | Mayo 2026 | Junio 2026 |
|---|---:|---:|
| Tiempo de proceso automático | **16 segundos** | **18 segundos** |
| Movimientos conciliados solos | 439 | 431 |
| Conciliaciones creadas | 287 | 291 |
| **Porcentaje del dinero resuelto** | **69 %** | **79 %** |
| Pendientes por revisar a mano | 605 | 768 |

### Por qué el porcentaje del dinero es lo que importa

Aunque en número de líneas la aplicación resuelve alrededor de un tercio, en **valor
resuelve entre el 69 % y el 79 %**. La razón: lo que cruza solo son los movimientos
**grandes** (pagos a proveedores, nómina, transferencias), mientras que lo pendiente son
en su mayoría **cobros pequeños** que el banco cobra uno por uno:

| Concepto pendiente (junio) | Movimientos |
|---|---:|
| Nómina | 198 |
| IVA sobre pagos automáticos | 131 |
| Comisiones y servicios bancarios | 123 |
| Intereses | 27 |
| 4x1000 (GMF) | 27 |
| Otros | 110 |

Estos **no se revisan uno por uno**: se agrupan por concepto y por día. Al agruparlos, los
768 movimientos pendientes de junio se reducen a **unas 240 decisiones** reales.

---

## 4. Tiempo estimado ahora

| Etapa | Tiempo |
|---|---|
| Cargar los dos archivos | ~1 minuto |
| Proceso automático de cruce | ~20 segundos |
| Revisión y cruces manuales de lo pendiente | ~45 a 60 minutos |
| **Total estimado** | **alrededor de 1 hora** |

### Cómo se llegó a esa cifra

Quedan **unas 240 decisiones** por resolver al mes. Una hora son 60 minutos, es decir
**cerca de 15 segundos por decisión** — un ritmo razonable, porque la aplicación ya
presenta los movimientos agrupados, con los totales calculados y el botón de cruzar
habilitado solo cuando los valores coinciden exactamente.

> **Advertencia honesta:** esta cifra es una **estimación**, no una medición. El tiempo
> automático (16-18 segundos) sí está medido; el tiempo manual depende de la persona.
> **El primer mes será más lento** mientras se aprende la herramienta. Conviene
> cronometrar los primeros dos cierres para reemplazar esta estimación por un dato real.

---

## 5. Beneficio

| | Antes | Ahora (estimado) |
|---|---:|---:|
| Tiempo por cierre mensual | 6 a 8 horas | **~1 hora** |
| Reducción | — | **cerca del 85 %** |
| Horas al año (12 cierres) | 72 a 96 horas | **~12 horas** |
| **Horas liberadas al año** | | **entre 60 y 84** |

Equivale a recuperar **más de una semana y media de trabajo al año** de una persona del
área financiera.

### Beneficios que no se miden en horas

- **Menos errores:** el cruce exige que el valor coincida **al centavo** y que el nombre
  corresponda a la misma persona. Se acabaron los cruces por coincidencia de cifras.
- **Trazabilidad:** cada conciliación queda con su identificador, su motivo y su fecha, y
  se puede **deshacer** sin alterar ni un dato original.
- **El conocimiento deja de depender de una persona:** el criterio de cruce está escrito
  en la herramienta, no en la cabeza de quien concilia.
- **Informes listos:** cada hoja se descarga en Excel con el membrete de la empresa, sin
  trabajo adicional de formato.
- **Control de cuadre automático:** la aplicación verifica que la diferencia entre banco y
  contabilidad coincida exactamente con las partidas sin cruzar. Si algo no cuadra, avisa.

---

## 6. Lo que todavía falta

Para ser transparentes sobre el estado actual:

| Pendiente | Consecuencia hoy |
|---|---|
| **No guarda la información** | Si se recarga la página se pierde el trabajo manual y hay que volver a conciliar |
| **Clave compartida** | No queda registro de *quién* concilió cada partida |
| **Hospedaje gratuito** | La aplicación puede tardar en responder y reiniciarse sola |

La más importante es la primera. Mientras no se resuelva, la aplicación sirve para
**revisión y pruebas**, pero no debería usarse como registro oficial del cierre.

---

*Documento elaborado con base en el procesamiento real de los archivos de mayo y junio de
2026. Las cifras de tiempo automático están medidas; las de tiempo manual son estimadas.*

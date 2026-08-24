# Contexto para Claude Code

App local de **conciliación bancaria** para ISTHO S.A.S. (Streamlit + pandas).
Cruza el extracto de Bancolombia contra el libro auxiliar contable.

## 📖 Leer primero: `documentos/DOCUMENTACION.md`

Ese archivo tiene todo: formato de los datos, las pasadas del algoritmo de cruce, el modelo
de estado para conciliar/desconciliar, la interfaz, la exportación a Excel y las trampas ya
encontradas. **Consultarlo antes de modificar el motor o la interfaz.**

## Lo mínimo para no romper nada

- **`conciliacion.py` es el motor**; `app.py` es solo interfaz. No mezclar.
- `reconciliar()` devuelve `(cruces, posibles)` — **listas de IDs, no tablas**. Las vistas se
  derivan con `construir_vistas()`. Así conciliar y desconciliar nunca alteran los datos
  originales.
- Tras cualquier cambio en el motor, correr la prueba de integridad: sin movimientos
  duplicados, cobertura completa y **el cuadre de Junio 2026 debe seguir dando 4.016.309,19**
  (Mayo 2026: 12.132.851,78) — antes y después de aplicar la segunda ronda.
- `aplicar_segunda_ronda()` es una función y un botón **aparte** de `reconciliar()`/Conciliar
  (a propósito, para poder revisar lo pendiente antes de aplicarla): cruza lo que queda
  pendiente por valor exacto **sin ningún límite de fecha**, sin exigir nombre, marcada
  «Baja». Ver documentación.
- Dentro de `reconciliar()` (automático, sin botón aparte) hay dos mejoras agregadas después
  de probarlas contra Julio 2026: **pasada 0** `_agrupar_por_manifiesto()` (agrupa
  `C. EGRESO TRANSPORTE` por número `MFxxxx` leído de `DETALLE`, más confiable que el nombre)
  y el **rescate de nombre desde `DETALLE`** en `load_libro_auxiliar()` (cuando
  `NOMBRE BENEFICIARIO` es solo "ISTHO SAS", sin nombre útil, se prueba a leer el nombre real
  de la cola de `DETALLE`, típico en nómina). Cualquier cambio a estas dos zonas necesita
  volver a correr la prueba de integridad en los 3 meses de prueba, no solo en Mayo/Junio.
- El extracto bancario **viene truncado a ~28 caracteres**; por eso la comparación de nombres
  exige subconjunto (no igualdad) y tolera prefijos.
- Al probar: cerrar la consola y reabrir `Iniciar_App.bat`. Refrescar el navegador no recarga
  los módulos importados.

## Archivos de prueba

`..\Extracto Junio 2026.xlsx` y `..\LIBRO AUXILIAR JUNIO 2026.xlsx` (también existen los de Mayo
y de Julio — el de Julio es `LIBRO AUXILIAR JULIO 2026-2.xlsx`, el libro auxiliar oficial
completo, el que sirvió para validar la pasada de manifiesto y el rescate de nombre).

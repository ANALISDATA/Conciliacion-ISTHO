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
  duplicados, cobertura completa y **el cuadre de Junio 2026 debe seguir dando 4.016.309,19**.
- El extracto bancario **viene truncado a ~28 caracteres**; por eso la comparación de nombres
  exige subconjunto (no igualdad) y tolera prefijos.
- Al probar: cerrar la consola y reabrir `Iniciar_App.bat`. Refrescar el navegador no recarga
  los módulos importados.

## Archivos de prueba

`..\Extracto Junio 2026.xlsx` y `..\LIBRO AUXILIAR JUNIO 2026.xlsx` (también existen los de Mayo).

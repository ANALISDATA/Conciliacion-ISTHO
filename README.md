# Conciliación Bancaria — ISTHO S.A.S.

Aplicación web que cruza automáticamente el **extracto bancario** contra el **libro auxiliar
contable**, permite conciliar y desconciliar a mano, y genera los informes en Excel con el
membrete de la empresa.

> **Documentación** (carpeta [`documentos/`](documentos/)):
> - **[NECESIDAD-Y-BENEFICIO.md](documentos/NECESIDAD-Y-BENEFICIO.md)** — qué problema
>   resuelve, cuánto tiempo ahorra y con qué cifras se midió.
> - **[DOCUMENTACION.md](documentos/DOCUMENTACION.md)** — cómo funciona por dentro: el
>   motor de cruce, el modelo de datos y las trampas ya encontradas.

---

## Cómo se usa

1. Cargar el **extracto bancario** y el **libro auxiliar** (`.xlsx` o `.csv`, tal como se
   descargan del banco y del sistema contable).
2. Escribir el **saldo inicial** (el «SALDO ANTERIOR» del extracto del mes pasado).
3. Pulsar **Conciliar**.
4. Revisar las seis hojas y descargar los Excel que se necesiten.

### Las seis hojas

| Hoja | Contenido |
|---|---|
| 📊 Resumen | Saldos del periodo y resultado del cruce |
| 🔗 Conciliados | Cruces encontrados, con opción de **deshacerlos** |
| 🔍 Por revisar | Coinciden en fecha y nombre, pero difieren en valor |
| 🏦 Pend. extracto | Movimientos del banco sin registro contable |
| 📘 Pend. libro auxiliar | Registros contables que no aparecen en el banco |
| ⚖️ Cruce manual | Cruzar a mano lo que quedó pendiente |

---

## Ejecutar en el computador

Con Python instalado:

```bash
pip install -r requirements.txt
streamlit run app.py
```

En Windows también sirve hacer doble clic en **`Iniciar_App.bat`**.

> Al cambiar `conciliacion.py`, `ui.py` o `excel_export.py` hay que **cerrar la consola y
> volver a abrirla**. Refrescar el navegador no basta: Streamlit no recarga los módulos
> importados.

---

## Estructura del proyecto

| Archivo | Qué hace |
|---|---|
| `app.py` | Pantallas, navegación y estado de la sesión |
| `conciliacion.py` | Lectura de archivos y **motor de cruce** |
| `ui.py` | Estilos, encabezado, tarjetas y tablas |
| `excel_export.py` | Generación de los Excel con membrete |
| `logo_istho.png` | Logo que va en la app y en los informes |
| `.streamlit/config.toml` | Colores institucionales |

---

## Privacidad

Este repositorio **debe ser privado**: contiene el logo, el NIT y el número de cuenta de la
empresa.

Los extractos y libros auxiliares **nunca** se suben. El `.gitignore` excluye `*.xlsx`,
`*.xls` y `*.csv` desde el primer commit, porque un archivo que entra al historial de Git no
se elimina con solo borrarlo después.

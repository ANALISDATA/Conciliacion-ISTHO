-- Ejecutar UNA VEZ en Supabase: Dashboard del proyecto -> SQL Editor -> pegar y correr.
--
-- Agrega a la tabla "conciliaciones" (proceso BANCARIO) las dos columnas donde se guardan,
-- al cerrar un mes, los pendientes que quedaron sin cruzar -- para poder arrastrarlos al mes
-- siguiente cuando se carguen sus archivos. Ver conciliacion.agregar_arrastre / db.guardar_cierre.
--
-- No afecta ningún dato existente: las columnas quedan vacías (NULL) en los meses ya
-- guardados, que es exactamente el mismo comportamiento que si no hubiera nada para arrastrar.

alter table conciliaciones
  add column if not exists pendientes_banco text,
  add column if not exists pendientes_libro text;

# Pharmapp

App para el celular del service técnico: registro diario de jornadas y viajes, rendición de gastos con fotos de tickets, órdenes de service con foto de la hoja firmada y cierre mensual con las planillas Excel de la empresa.

- Se instala desde Chrome ("Instalar app") y funciona sin internet.
- Los datos quedan guardados en el celular. Las fotos se respaldan en Google Drive cuando hay conexión (Más → Configurar Drive).
- La planilla de horas tiene diseño propio (`excel.js`); la rendición de gastos se genera sobre la plantilla de la empresa en `templates/` (sin datos personales).
- Feriados nacionales: una vez por semana consulta api.argentinadatos.com (con date.nager.at de respaldo) y pide confirmación.

Archivos: `index.html` (pantalla), `app.js` (lógica), `excel.js` (planillas), `drive.js` (respaldo), `db.js` (base local), `sw.js` (modo sin internet).
Al publicar cambios, subir el número de `VERSION` en `sw.js` para que los celulares bajen la versión nueva.

# Pharmapp – guía para Claude

App web instalable (PWA) para el celular de un técnico de service de MG Pharmacorp (sucursal Bahía Blanca). Sin build: HTML + JS plano, publicada con GitHub Pages desde `main` (raíz) en https://nacho-rochon.github.io/pharmapp/.

## Cómo hablar con la usuaria
- Español rioplatense, simple, sin jerga técnica.
- Antes de crear cuentas, publicar fuera del repo o tocar su Google/Drive, preguntar.

## Archivos
- `index.html`, `styles.css`: pantalla y estilos (5 pestañas: Día, Viajes, Services, Mes, Más).
- `app.js`: toda la lógica de pantallas y acciones (`actions` = botones con `data-act`).
- `util.js`: fechas, períodos (del 26 al 25), tipos de día, cálculo de horas extra.
- `excel.js`: planilla de horas (diseño propio con ExcelJS), rendición de gastos (sobre `templates/rendicion.xlsx`, planilla oficial de la empresa) y resumen de services.
- `drive.js`: respaldo de fotos en Google Drive vía un Apps Script que la usuaria pegó en su cuenta (el código del script se genera en la app, Más → Configurar Drive).
- `db.js`: IndexedDB (days, clients, trips, expenses, services, photos, settings). Todo funciona sin internet.
- `sw.js`: service worker. Con internet trae siempre lo último (network-first, 4 s de espera) y sin internet usa la copia guardada; al instalarse una versión nueva la app se recarga sola. **En cada cambio subir `VERSION`** (y el texto "Pharmapp vN" en app.js).
- `vendor/`: ExcelJS 4.4.0 y JSZip 3.10.1 (locales para modo offline).

## Reglas de negocio
- Jornada normal 8 a 16. Horas extra = lo que pasa de 8 hs en día hábil + todo lo trabajado en fin de semana, feriado o día libre.
- Tipos de día: normal, viaje (ida o ida y vuelta), otra (jornada en otra localidad), regreso, extra, vacaciones, feriado, franco, licencia, estudio (con foto del comprobante).
- En `viaje` y `regreso` la jornada es salida → llegada (no se piden inicio/fin aparte). Solo esos dos llevan "VIAJE: SI" y km en la planilla.
- Colores en la app: lila = viaje / otra localidad / regreso (aunque haya extra); naranja = día con horas extra.
- Rendición: un Excel por viaje y medio de pago (VISA empresa o gasto personal). Saldo VISA al ir (E31) y al volver.
- Feriados: una vez por semana consulta api.argentinadatos.com (respaldo date.nager.at) y pide confirmación antes de marcar.

## Repo público
Nunca subir datos personales (nombres, saldos, clientes reales con datos, fotos). El nombre del empleado se carga en Ajustes de la app.

## Probar
Servir la carpeta (`python3 -m http.server`) y probar con Playwright en un perfil de celular (Pixel 7). Revisar los Excel convirtiéndolos a PDF con LibreOffice.

// Utilidades compartidas: fechas, períodos, horas y formatos.
const U = (() => {
  const DIAS = ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'];
  const DIAS_LARGO = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const pad = n => String(n).padStart(2, '0');
  const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const parse = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
  const today = () => ymd(new Date());
  const addDays = (s, n) => { const d = parse(s); d.setDate(d.getDate() + n); return ymd(d); };
  const dow = s => parse(s).getDay();
  const isWeekend = s => { const w = dow(s); return w === 0 || w === 6; };
  const ddmm = s => { const [, m, d] = s.split('-'); return `${d}/${m}`; };
  const ddmmyy = s => { const [y, m, d] = s.split('-'); return `${d}/${m}/${y.slice(2)}`; };
  const largo = s => { const d = parse(s); return `${DIAS_LARGO[d.getDay()]} ${d.getDate()} de ${MESES[d.getMonth()].toLowerCase()}`; };
  const excelDate = s => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };

  // Horas "08:30" <-> minutos
  const toMin = t => { if (!t) return null; const [h, m] = t.split(':').map(Number); return h * 60 + (m || 0); };
  const frac = t => { const m = toMin(t); return m == null ? null : m / 1440; }; // fracción de día para Excel
  const fmtHs = h => (Math.round(h * 100) / 100).toString().replace('.', ',');

  // Período de la planilla: del (cierre+1) del mes anterior al (cierre) del mes. Clave = 'AAAA-MM' del mes que cierra.
  function periodOf(s, cierre = 25) {
    const d = parse(s);
    if (d.getDate() > cierre) d.setMonth(d.getMonth() + 1, 1);
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
  }
  function periodRange(key, cierre = 25) {
    const [y, m] = key.split('-').map(Number);
    const end = ymd(new Date(y, m - 1, cierre));
    const start = ymd(new Date(y, m - 2, cierre + 1));
    const days = [];
    for (let s = start; s <= end; s = addDays(s, 1)) days.push(s);
    const startMonth = parse(start).getMonth();
    return { key, start, end, days, year: y, mesFin: MESES[m - 1], mesInicio: MESES[startMonth] };
  }
  const shiftPeriod = (key, n) => { const [y, m] = key.split('-').map(Number); const d = new Date(y, m - 1 + n, 1); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; };

  // Tipos de día. css = color en la app; viaje: true = días en otra localidad (color lila)
  const TIPOS = {
    normal: { label: 'Jornada normal', css: '' },
    viaje: { label: 'Viaje (ida, o ida y vuelta)', corto: 'Viaje', css: 'viaje', viaje: true },
    otra: { label: 'Jornada en otra localidad', corto: 'Otra localidad', css: 'viaje', viaje: true },
    regreso: { label: 'Viaje de regreso', corto: 'Regreso', css: 'viaje', viaje: true },
    extra: { label: 'Trabajo en día libre', css: 'extra' },
    vacaciones: { label: 'Vacaciones', css: 'libre', obs: 'VACACIONES' },
    feriado: { label: 'Feriado', css: 'libre', obs: 'FERIADO' },
    franco: { label: 'Franco / no laborable', css: 'libre', obs: 'FRANCO' },
    licencia: { label: 'Licencia / enfermedad', css: 'libre', obs: 'LICENCIA' },
    estudio: { label: 'Licencia por estudio', css: 'libre', obs: 'LICENCIA POR ESTUDIO' }
  };
  const SIN_TRABAJO = ['vacaciones', 'feriado', 'franco', 'licencia', 'estudio'];
  const VIAJA = ['viaje', 'regreso']; // días en los que se maneja: inicio = salida, fin = llegada

  // Horario efectivo del día
  function times(rec) {
    if (!rec || SIN_TRABAJO.includes(rec.tipo)) return { inicio: null, fin: null };
    if (VIAJA.includes(rec.tipo)) return { inicio: rec.salida || rec.inicio || null, fin: rec.llegada || rec.fin || null };
    return { inicio: rec.inicio || null, fin: rec.fin || null };
  }

  // Horas trabajadas y horas extra de un día
  function calcDay(rec, s, cfg) {
    const { inicio, fin } = times(rec);
    const a = toMin(inicio), b0 = toMin(fin);
    if (a == null || b0 == null) return { dur: 0, extra: 0 };
    const b = b0 <= a ? b0 + 1440 : b0; // vuelta pasada la medianoche
    const dur = (b - a) / 60;
    const base = (toMin(cfg.fin) - toMin(cfg.inicio)) / 60;
    const libre = isWeekend(s) || rec.tipo === 'extra' || rec.feriado;
    const extra = libre ? dur : Math.max(0, dur - base);
    return { dur: Math.round(dur * 100) / 100, extra: Math.round(extra * 100) / 100 };
  }

  const money = n => (n || 0).toLocaleString('es-AR', { style: 'currency', currency: 'ARS', minimumFractionDigits: 2 });
  const num = v => { if (v === '' || v == null) return null; const n = Number(String(v).replace(/\./g, '').replace(',', '.')); return isNaN(n) ? null : n; };
  const safe = s => String(s || '').replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // Conceptos de la planilla de rendición (fila en el Excel)
  const CONCEPTOS = [
    ['desayuno', 'Desayuno', 10], ['almuerzo', 'Almuerzo', 11], ['merienda', 'Merienda', 12], ['cena', 'Cena', 13],
    ['hotel', 'Hotel', 14], ['combustible', 'Est. Servicio', 15], ['peajes', 'Peajes', 16], ['estacionamiento', 'Estacionam.', 17],
    ['colectivo', 'Pasajes de colectivo', 18], ['transporte', 'Transporte', 19], ['otros', 'Otros', 20]
  ];
  const TIPOS_SERVICE = ['Mantenimiento correctivo', 'Mantenimiento preventivo', 'Instalación', 'Revisión / diagnóstico', 'Capacitación', 'Otro'];

  return { DIAS, DIAS_LARGO, MESES, pad, ymd, parse, today, addDays, dow, isWeekend, ddmm, ddmmyy, largo, excelDate, toMin, frac, fmtHs,
    periodOf, periodRange, shiftPeriod, TIPOS, SIN_TRABAJO, VIAJA, times, calcDay, money, num, safe, esc, CONCEPTOS, TIPOS_SERVICE };
})();
if (typeof module !== 'undefined') module.exports = U;

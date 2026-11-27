// Generación de los Excel: planilla de horas (diseño propio) y rendición (sobre la planilla de la empresa).
const XL = (() => {
  let env = {
    ExcelJS: typeof ExcelJS !== 'undefined' ? ExcelJS : null,
    loadTemplate: async name => (await fetch(`templates/${name}.xlsx`)).arrayBuffer(),
    logo: async () => (await fetch('templates/logo.jpeg')).arrayBuffer()
  };
  const setup = o => Object.assign(env, o);
  const lib = () => env.ExcelJS || (typeof ExcelJS !== 'undefined' ? ExcelJS : null);

  // ---------- Planilla de horas ----------
  // Arma la fila de un día tal como va al Excel (también la usa la vista previa de la app).
  function resolveDay(s, rec, cfg, fillMissing, svcs) {
    const weekend = U.isWeekend(s);
    let r = rec;
    let auto = false;
    if (!r && !weekend && fillMissing) { r = { tipo: 'normal', inicio: cfg.inicio, fin: cfg.fin }; auto = true; }
    const out = { date: s, tipo: r ? r.tipo : null, auto, missing: !r && !weekend, viaje: 'NO',
      salida: null, llegada: null, localidad: '', km: null, cliente: '', inicio: null, fin: null, obs: '', dur: 0, extra: 0 };
    if (!r) return out;
    const T = U.TIPOS[r.tipo] || U.TIPOS.normal;
    if (U.SIN_TRABAJO.includes(r.tipo)) { out.obs = r.obs ? `${T.obs}: ${r.obs}` : T.obs; return out; }
    const t = U.times(r);
    out.inicio = t.inicio; out.fin = t.fin;
    if (U.VIAJA.includes(r.tipo)) {
      out.viaje = 'SI'; out.salida = t.inicio; out.llegada = t.fin;
      out.localidad = r.localidad || ''; out.km = r.km ?? null; out.cliente = r.cliente || '';
    } else if (r.tipo === 'otra') {
      out.localidad = r.localidad || ''; out.cliente = r.cliente || '';
    } else if (r.cliente) out.cliente = r.cliente;
    const c = U.calcDay(r, s, cfg);
    out.dur = c.dur; out.extra = c.extra;
    let obs = (r.obs || '').trim();
    if (!obs && svcs && svcs.length) obs = svcs.map(v => [v.equipo, v.tareas].filter(Boolean).join(': ')).join(' / ');
    if (r.feriado && !/feriado/i.test(obs)) obs = `Feriado trabajado${r.feriadoNombre ? ' (' + r.feriadoNombre + ')' : ''}` + (obs ? '. ' + obs : '');
    if (r.tipo === 'idavuelta') obs = `Ida y vuelta en el día${r.kmBase ? ` (${r.kmBase} km × 2 = ${r.km} km)` : ''}` + (obs ? '. ' + obs : '');
    if (r.tipo === 'regreso' && !/regreso|vuelta/i.test(obs)) obs = 'Vuelta desde ' + (r.localidad || 'viaje') + (obs ? '. ' + obs : '');
    if (r.tipo === 'otra' && !obs) obs = 'Jornada en ' + (r.localidad || 'otra localidad');
    out.obs = obs;
    return out;
  }

  // Colores de la planilla
  const C = { azul: 'FF0B5CAB', azulClaro: 'FFE7F0FA', gris: 'FFF2F4F7', grisTxt: 'FF64707D', linea: 'FFD5DCE4', lila: 'FFF1E9FB',
    naranja: 'FFFFB45C', naranjaTxt: 'FF8A3D00', libre: 'FFFCE1EA', blanco: 'FFFFFFFF' };
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin', color: { argb: C.linea } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  async function buildHoras({ periodKey, days, services, clients, cfg, fillMissing }) {
    const P = U.periodRange(periodKey, cfg.cierre);
    const wb = new (lib().Workbook)();
    wb.creator = 'Pharmapp';
    const ws = wb.addWorksheet(`${P.mesFin.toUpperCase()} ${P.year}`, {
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } },
      headerFooter: { oddFooter: '&L&8Generado con Pharmapp&R&8Página &P de &N' },
      views: [{ showGridLines: false }]
    });
    const COLS = [
      ['Día', 6], ['Fecha', 10], ['Tipo de día', 17], ['Viaje', 7], ['Hs salida', 9], ['Hs llegada', 9], ['Localidad', 16], ['Km', 7],
      ['Cliente / lugar', 20], ['Inicio jornada', 9], ['Fin jornada', 9], ['Hs trabajadas', 10], ['Hs extra', 9], ['Observaciones', 52]
    ];
    COLS.forEach(([, w], i) => { ws.getColumn(i + 1).width = w; });
    const N = COLS.length, last = String.fromCharCode(64 + N);

    // Encabezado
    ws.mergeCells(`A1:${String.fromCharCode(64 + N - 2)}1`);
    const t = ws.getCell('A1');
    t.value = 'REGISTRO DE HORAS TRABAJADAS Y VIAJES REALIZADOS';
    t.font = { bold: true, size: 15, color: { argb: C.azul } };
    ws.getRow(1).height = 26;
    const info = [['Empleado/a', cfg.empleado || ''], ['Sector', cfg.sector || 'SERVICIO TECNICO'],
      ['Período', `del ${U.ddmmyy(P.start)} al ${U.ddmmyy(P.end)}  (${P.mesInicio} – ${P.mesFin} ${P.year})`], ['Sucursal', cfg.sucursal || '']];
    info.forEach(([k, v], i) => {
      const r = ws.getRow(2 + i);
      r.getCell(1).value = k; r.getCell(1).font = { bold: true, color: { argb: C.grisTxt } };
      ws.mergeCells(2 + i, 1, 2 + i, 2);
      r.getCell(3).value = v; r.getCell(3).font = { bold: true };
      ws.mergeCells(2 + i, 3, 2 + i, 9);
    });
    if (env.logo) {
      try {
        const id = wb.addImage({ buffer: await env.logo(), extension: 'jpeg' });
        ws.addImage(id, { tl: { col: N - 1 + 0.55, row: 0.1 }, ext: { width: 92, height: 94 } });
      } catch (e) { /* sin logo */ }
    }

    // Filas de días
    const svcBy = {};
    for (const v of services || []) (svcBy[v.fecha] = svcBy[v.fecha] || []).push(v);
    const rows = P.days.map(s => resolveDay(s, days[s], cfg, fillMissing, svcBy[s]));
    const total = rows.reduce((a, r) => a + r.extra, 0);
    const kmTot = rows.reduce((a, r) => a + (r.viaje === 'SI' && r.km ? r.km : 0), 0);
    const trab = rows.filter(r => r.inicio).length;
    const viajes = rows.filter(r => r.viaje === 'SI').length;

    // Resumen
    const RS = 7;
    [['Días trabajados', trab], ['Días de viaje', viajes], ['Km recorridos', kmTot], ['Total hs extra', total]].forEach(([k, v], i) => {
      const c0 = 1 + i * 3;
      ws.mergeCells(RS, c0, RS, c0 + 1);
      const a = ws.getRow(RS).getCell(c0); a.value = k; a.font = { size: 9, color: { argb: C.grisTxt } }; a.alignment = { horizontal: 'center' };
      ws.mergeCells(RS + 1, c0, RS + 1, c0 + 1);
      const b = ws.getRow(RS + 1).getCell(c0); b.value = v; b.font = { bold: true, size: 14, color: { argb: i === 3 ? C.naranjaTxt : C.azul } };
      b.alignment = { horizontal: 'center' }; if (i === 3) b.numFmt = '0.##';
      for (let c = c0; c <= c0 + 1; c++) for (const rr of [RS, RS + 1]) { ws.getRow(rr).getCell(c).fill = fill(i === 3 ? C.naranja : C.azulClaro); }
    });

    const H = RS + 3;
    const hr = ws.getRow(H);
    COLS.forEach(([name], i) => {
      const c = hr.getCell(i + 1);
      c.value = name; c.font = { bold: true, color: { argb: C.blanco }, size: 10 }; c.fill = fill(C.azul);
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; c.border = border;
    });
    hr.height = 30;
    ws.views = [{ state: 'frozen', ySplit: H, showGridLines: false }];
    ws.autoFilter = { from: { row: H, column: 1 }, to: { row: H, column: N } };

    rows.forEach((d, i) => {
      const s = d.date;
      const row = ws.getRow(H + 1 + i);
      const T = d.tipo ? U.TIPOS[d.tipo] : null;
      const tipoTxt = d.tipo ? (T.corto || T.label) + (d.auto ? ' (auto)' : '') : (U.isWeekend(s) ? 'Fin de semana' : '—');
      const vals = [U.DIAS[U.dow(s)], U.excelDate(s), tipoTxt, d.viaje === 'SI' ? 'SI' : 'NO', U.frac(d.salida), U.frac(d.llegada), d.localidad || null,
        d.km ?? null, d.cliente || null, U.frac(d.inicio), U.frac(d.fin), d.dur || null, d.extra || null, d.obs || null];
      vals.forEach((v, k) => { row.getCell(k + 1).value = v; });
      let bg = null;
      if (T && T.viaje) bg = C.lila;
      else if (d.extra > 0) bg = C.naranja;
      else if ((d.tipo && U.SIN_TRABAJO.includes(d.tipo)) || (!d.tipo && U.isWeekend(s))) bg = C.libre;
      for (let k = 1; k <= N; k++) {
        const c = row.getCell(k);
        c.border = border; c.font = { size: 10 };
        c.alignment = { vertical: 'middle', horizontal: [1, 2, 4, 5, 6, 8, 10, 11, 12, 13].includes(k) ? 'center' : 'left', wrapText: k === N || k === 9 };
        if (bg) c.fill = fill(bg);
      }
      row.getCell(2).numFmt = 'dd/mm/yy';
      for (const k of [5, 6, 10, 11]) row.getCell(k).numFmt = 'h:mm';
      row.getCell(12).numFmt = '0.##'; row.getCell(13).numFmt = '0.##';
      if (d.extra > 0) row.getCell(13).font = { bold: true, size: 10, color: { argb: C.naranjaTxt } };
      if (U.dow(s) === 0) row.getCell(1).font = { bold: true, size: 10, color: { argb: 'FFC0392B' } };
      if (d.missing && !d.auto) row.getCell(3).font = { italic: true, size: 10, color: { argb: 'FFC0392B' } };
    });
    // Totales
    const TR = H + 1 + rows.length, first = H + 1, lastR = TR - 1;
    const tot = ws.getRow(TR);
    tot.getCell(1).value = 'TOTALES'; ws.mergeCells(TR, 1, TR, 7);
    tot.getCell(8).value = { formula: `SUM(H${first}:H${lastR})`, result: kmTot };
    tot.getCell(12).value = { formula: `SUM(L${first}:L${lastR})`, result: rows.reduce((a, r) => a + r.dur, 0) };
    tot.getCell(13).value = { formula: `SUM(M${first}:M${lastR})`, result: total };
    tot.getCell(14).value = `Total: ${U.fmtHs(total)} hs extra`;
    for (let k = 1; k <= N; k++) {
      const c = tot.getCell(k);
      c.font = { bold: true, color: { argb: C.blanco } }; c.fill = fill(C.azul); c.border = border;
      c.alignment = { horizontal: k === 1 || k === N ? 'left' : 'center', vertical: 'middle' };
      if ([12, 13].includes(k)) c.numFmt = '0.##';
    }
    tot.height = 22;
    // Referencias de colores
    const ref = ws.getRow(TR + 2);
    [['Viaje / otra localidad', C.lila], ['Horas extra', C.naranja], ['Fin de semana / feriado / licencia', C.libre]].forEach(([k, col], i) => {
      const c = ref.getCell(1 + i * 4); c.fill = fill(col); c.border = border;
      const l = ref.getCell(2 + i * 4); l.value = k; l.font = { size: 9, color: { argb: C.grisTxt } };
    });
    ws.pageSetup.printArea = `A1:${last}${TR + 2}`;
    ws.pageSetup.printTitlesRow = `${H}:${H}`;

    // Hoja 2: detalle de horas extra (para rendirlas)
    const he = wb.addWorksheet('HORAS EXTRA', { views: [{ showGridLines: false }] });
    he.columns = [{ width: 6 }, { width: 11 }, { width: 18 }, { width: 9 }, { width: 9 }, { width: 10 }, { width: 70 }];
    he.getCell('A1').value = `Horas extra – ${P.mesFin} ${P.year} – ${cfg.empleado || ''}`; he.getCell('A1').font = { bold: true, size: 13, color: { argb: C.azul } };
    const hh = he.getRow(3);
    ['Día', 'Fecha', 'Tipo', 'Inicio', 'Fin', 'Hs extra', 'Motivo / tareas'].forEach((n, i) => {
      const c = hh.getCell(i + 1); c.value = n; c.font = { bold: true, color: { argb: C.blanco } }; c.fill = fill(C.azul); c.border = border; c.alignment = { horizontal: 'center' };
    });
    const ex = rows.filter(r => r.extra > 0);
    ex.forEach((d, i) => {
      const r = he.getRow(4 + i);
      const T = U.TIPOS[d.tipo];
      [U.DIAS[U.dow(d.date)], U.excelDate(d.date), T ? (T.corto || T.label) : '', U.frac(d.inicio), U.frac(d.fin), d.extra, d.obs || ''].forEach((v, k) => { r.getCell(k + 1).value = v; });
      r.getCell(2).numFmt = 'dd/mm/yy'; r.getCell(4).numFmt = r.getCell(5).numFmt = 'h:mm'; r.getCell(6).numFmt = '0.##';
      for (let k = 1; k <= 7; k++) { const c = r.getCell(k); c.border = border; c.alignment = { vertical: 'top', wrapText: k === 7, horizontal: k === 7 || k === 3 ? 'left' : 'center' }; }
    });
    const ht = he.getRow(4 + ex.length);
    ht.getCell(5).value = 'TOTAL'; ht.getCell(6).value = { formula: ex.length ? `SUM(F4:F${3 + ex.length})` : '0', result: total }; ht.getCell(6).numFmt = '0.##';
    for (let k = 1; k <= 7; k++) { ht.getCell(k).fill = fill(C.naranja); ht.getCell(k).font = { bold: true }; ht.getCell(k).border = border; }
    he.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

    // Hoja 3: localidades
    const loc = wb.addWorksheet('LOCALIDADES');
    loc.columns = [{ header: 'LOCALIDAD', width: 18 }, { header: 'KM', width: 8 }, { header: 'CLIENTE/LUGAR', width: 40 }];
    loc.getRow(1).eachCell(c => { c.font = { bold: true, color: { argb: C.blanco } }; c.fill = fill(C.azul); });
    (clients || []).forEach(c => loc.addRow([c.localidad, c.km ?? null, c.clientes]));

    const parts = (cfg.empleado || '').trim().split(/\s+/);
    const nombre = parts.length > 1 ? [...parts.slice(1), parts[0]].join(' ') : (cfg.empleado || '');
    const filename = U.safe(`${P.mesFin.toUpperCase()} ${nombre.toUpperCase()} Planilla Serv. Tecnico ${cfg.sucursalCorta || 'BB'}`) + '.xlsx';
    return { buffer: await wb.xlsx.writeBuffer(), filename, total, rows, period: P };
  }

  // ---------- Rendición de gastos de un viaje ----------
  const MON = '_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)';
  async function buildRendicion({ trip, expenses, medio, cfg }) {
    const list = expenses.filter(e => e.medio === medio);
    const wb = new (lib().Workbook)();
    await wb.xlsx.load(await env.loadTemplate('rendicion'));
    const ws = wb.worksheets[0];
    const esVisa = medio === 'visa';
    ws.getCell('D2').value = (trip.motivo || '').toUpperCase() + (esVisa ? '' : ' – GASTOS PERSONALES (REINTEGRO)');
    ws.getCell('K2').value = U.excelDate(trip.presentacion || U.today());
    ws.getCell('C3').value = cfg.sucursal || 'Bahia Blanca';
    ws.getCell('C4').value = (trip.localidad || '').toUpperCase();
    ws.getCell('F4').value = U.excelDate(trip.desde);
    ws.getCell('G4').value = U.excelDate(trip.hasta || trip.desde);

    let fechas = [...new Set(list.map(e => e.fecha))].sort();
    if (!fechas.length) fechas = [trip.desde];
    const extraCols = fechas.length > 8;
    if (extraCols) fechas = fechas.slice(0, 7).concat(['resto']);
    const colOf = {};
    fechas.forEach((f, i) => {
      const c = 3 + i; colOf[f] = c;
      const cell = ws.getRow(9).getCell(c);
      cell.value = f === 'resto' ? 'Otros días' : U.excelDate(f);
    });
    const sums = {};
    for (const e of list) {
      const conc = U.CONCEPTOS.find(c => c[0] === e.concepto) || U.CONCEPTOS[10];
      const c = colOf[e.fecha] || colOf.resto;
      const k = conc[2] + ':' + c;
      sums[k] = (sums[k] || 0) + (Number(e.monto) || 0);
    }
    if (Object.keys(sums).some(k => k.startsWith('20:'))) ws.getCell('B20').value = 'Otros';
    for (const [k, v] of Object.entries(sums)) {
      const [r, c] = k.split(':').map(Number);
      const cell = ws.getRow(r).getCell(c);
      cell.value = Math.round(v * 100) / 100; cell.numFmt = MON;
    }
    const L = n => String.fromCharCode(64 + n);
    const f22 = ws.getCell('F22').style;
    for (let c = 3; c <= 10; c++) {
      const cell = ws.getRow(22).getCell(c);
      if (c > 6) cell.style = JSON.parse(JSON.stringify(f22));
      cell.value = { formula: `SUM(${L(c)}10:${L(c)}21)` };
      cell.numFmt = MON;
    }
    const total = list.reduce((a, e) => a + (Number(e.monto) || 0), 0);
    ws.getCell('E32').value = { formula: 'SUM(C22:J22)', result: total };
    if (esVisa) {
      ws.getCell('E31').value = trip.saldoIda ?? null;
      ws.getCell('E33').value = { formula: 'E31-E32', result: (trip.saldoIda || 0) - total };
      if (trip.saldoVuelta != null) {
        for (const c of ['B', 'C', 'D', 'E']) ws.getCell(`${c}34`).style = JSON.parse(JSON.stringify(ws.getCell(`${c}33`).style));
        ws.getCell('B34').value = 'Saldo real al volver';
        ws.getCell('E34').value = trip.saldoVuelta; ws.getCell('E34').numFmt = MON;
      }
    } else {
      ws.getCell('B31').value = 'Total a reintegrar:';
      ws.getCell('E31').value = { formula: 'SUM(C22:J22)', result: total }; ws.getCell('E31').numFmt = MON;
      ws.getCell('E33').value = null; ws.getCell('B33').value = null; ws.getCell('B32').value = null; ws.getCell('E32').value = null;
    }
    const apellido = (cfg.empleado || '').trim().split(/\s+/)[0] || '';
    const dias = [...new Set([trip.desde, trip.hasta].filter(Boolean))].map(s => U.parse(s).getDate());
    const mes = U.parse(trip.hasta || trip.desde).getMonth() + 1;
    const filename = U.safe(`${dias.join(' y ')}-${mes} ${apellido} - ${esVisa ? 'VISA MG PHARMA' : 'GASTO PERSONAL'} gastos ${trip.motivo || ''} ${(trip.localidad || '').toUpperCase()} PLANILLA DE RENDICION DE GASTOS`) + '.xlsx';
    return { buffer: await wb.xlsx.writeBuffer(), filename, total, extraCols };
  }

  // ---------- Resumen de órdenes de service ----------
  async function buildServices({ services, title, photoNames }) {
    const wb = new (lib().Workbook)();
    const ws = wb.addWorksheet('Services');
    ws.addRow([title]).font = { bold: true, size: 14 };
    ws.addRow([]);
    const head = ['Fecha', 'Cliente', 'Localidad', 'Equipo', 'Marca', 'Modelo', 'Nº de Serie', 'Tipo de Servicio', 'Descripción del Problema',
      'Tareas Realizadas', 'Observaciones', 'Solución', 'Piezas Reemplazadas', 'Tiempo Trabajo', 'Tiempo Viaje', 'Viaje', 'Fotos'];
    const h = ws.addRow(head);
    h.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B5CAB' } }; c.alignment = { vertical: 'middle', wrapText: true }; });
    for (const v of services) {
      const row = ws.addRow([U.excelDate(v.fecha), v.cliente, v.localidad, v.equipo, v.marca, v.modelo, v.serie, v.tipo, v.problema, v.tareas,
        v.observaciones, v.solucion, v.piezas, v.tiempoTrabajo, v.tiempoViaje, v.tripName || '', (photoNames[v.id] || []).join(', ')]);
      row.getCell(1).numFmt = 'dd/mm/yy';
      row.eachCell(c => { c.alignment = { vertical: 'top', wrapText: true }; });
    }
    const widths = [10, 22, 16, 16, 14, 14, 14, 20, 36, 44, 30, 9, 26, 12, 12, 22, 40];
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.views = [{ state: 'frozen', ySplit: 3 }];
    return wb.xlsx.writeBuffer();
  }

  return { setup, resolveDay, buildHoras, buildRendicion, buildServices };
})();
if (typeof module !== 'undefined') module.exports = XL;

// Generación de los Excel sobre las planillas originales de la empresa.
const XL = (() => {
  let env = {
    ExcelJS: typeof ExcelJS !== 'undefined' ? ExcelJS : null,
    loadTemplate: async name => (await fetch(`templates/${name}.xlsx`)).arrayBuffer()
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
      salida: null, llegada: null, localidad: '', km: null, cliente: '', inicio: null, fin: null, obs: '', extra: 0 };
    if (!r) return out;
    if (U.SIN_TRABAJO.includes(r.tipo)) { out.obs = r.obs || U.TIPOS[r.tipo].obs; return out; }
    out.inicio = r.inicio; out.fin = r.fin;
    if (r.tipo === 'viaje') {
      out.viaje = 'SI'; out.salida = r.salida || null; out.llegada = r.llegada || null;
      out.localidad = r.localidad || ''; out.km = r.km ?? null; out.cliente = r.cliente || '';
    } else if (r.cliente) out.cliente = r.cliente;
    const { extra } = U.calcDay(r, s, cfg);
    out.extra = extra;
    let obs = (r.obs || '').trim();
    if (!obs && svcs && svcs.length) obs = svcs.map(v => [v.equipo, v.tareas].filter(Boolean).join(': ')).join(' / ');
    if (extra > 0 && cfg.prefijoExtra !== false && !/\bhs\b/i.test(obs)) obs = `${U.fmtHs(extra)} hs extra${obs ? ': ' + obs : ''}`;
    out.obs = obs;
    return out;
  }

  const COLS = { A: 1, B: 2, C: 3, F: 6, I: 9, K: 11, Q: 17, S: 19, X: 24, AA: 27, AD: 30, AI: 35 };
  const MERGES = [['C', 'E'], ['F', 'H'], ['I', 'J'], ['K', 'P'], ['Q', 'R'], ['S', 'W'], ['X', 'Z'], ['AA', 'AC'], ['AD', 'AI']];

  async function buildHoras({ periodKey, days, services, clients, cfg, fillMissing }) {
    const P = U.periodRange(periodKey, cfg.cierre);
    const wb = new (lib().Workbook)();
    await wb.xlsx.load(await env.loadTemplate('horas'));
    const ws = wb.getWorksheet('PLANILLA');
    ws.name = `${P.mesFin.toUpperCase()} ${P.year}`;
    ws.getCell('C2').value = cfg.empleado || '';
    ws.getCell('C3').value = cfg.sector || 'SERVICIO TECNICO';
    ws.getCell('C4').value = P.mesInicio;
    ws.getCell('G4').value = P.mesFin;
    ws.getCell('L4').value = P.year;
    ws.getCell('E5').value = U.excelDate(P.start);
    ws.getCell('K5').value = U.excelDate(P.end);

    const svcBy = {};
    for (const v of services || []) (svcBy[v.fecha] = svcBy[v.fecha] || []).push(v);
    const rows = [];
    let total = 0;
    P.days.forEach((s, i) => {
      const r = 8 + i;
      if (r > 37) { // 31 días: la plantilla trae 30 filas con formato, copiamos el de la anterior
        const src = ws.getRow(37), dst = ws.getRow(r);
        dst.height = src.height;
        for (let c = 1; c <= 35; c++) dst.getCell(c).style = JSON.parse(JSON.stringify(src.getCell(c).style || {}));
        for (const [a, b] of MERGES) { try { ws.mergeCells(`${a}${r}:${b}${r}`); } catch (e) { /* ya combinada */ } }
        const w = U.dow(s);
        if (w === 0 || w === 6) for (let c = 1; c <= 35; c++) dst.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: w === 0 ? 'FFFF5050' : 'FFFFF2CC' } };
      }
      const d = resolveDay(s, days[s], cfg, fillMissing, svcBy[s]);
      rows.push(d);
      total += d.extra;
      const row = ws.getRow(r);
      const set = (col, v) => { row.getCell(COLS[col]).value = v; };
      set('A', U.DIAS[U.dow(s)]);
      set('B', U.excelDate(s)); row.getCell(2).numFmt = 'dd/mm/yy';
      set('C', d.viaje);
      set('F', U.frac(d.salida)); set('I', U.frac(d.llegada));
      set('K', d.localidad || null); set('Q', d.km ?? null); set('S', d.cliente || null);
      set('X', U.frac(d.inicio)); set('AA', U.frac(d.fin));
      set('AD', d.obs || null);
      for (const c of ['F', 'I', 'X', 'AA']) row.getCell(COLS[c]).numFmt = 'h:mm';
    });
    const totRow = 8 + P.days.length;
    if (totRow !== 38) { ws.getCell(`AI${totRow}`).style = JSON.parse(JSON.stringify(ws.getCell('AI38').style || {})); }
    ws.getCell(`AI${totRow}`).value = `Total: ${U.fmtHs(total)} hs extra`;
    ws.getCell(`AI${totRow}`).font = { bold: true };
    ws.getCell(`AI${totRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

    // Hoja LOCALIDADES con la base actual
    const loc = wb.getWorksheet('LOCALIDADES');
    if (loc) {
      for (let r = 2; r <= Math.max(loc.rowCount, 2); r++) for (let c = 1; c <= 3; c++) loc.getRow(r).getCell(c).value = null;
      (clients || []).forEach((c, i) => { const row = loc.getRow(i + 2); row.getCell(1).value = c.localidad; row.getCell(2).value = c.km ?? null; row.getCell(3).value = c.clientes; });
    }
    const apellido = (cfg.empleado || '').trim().split(/\s+/);
    const nombreArchivo = apellido.length > 1 ? [...apellido.slice(1), apellido[0]].join(' ') : (cfg.empleado || '');
    const filename = U.safe(`${P.mesFin.toUpperCase()} ${nombreArchivo.toUpperCase()} Planilla Serv. Tecnico ${cfg.sucursalCorta || 'BB'}`) + '.xlsx';
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

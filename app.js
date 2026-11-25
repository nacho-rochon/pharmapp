// Pharmapp – pantalla principal
(() => {
  const $ = s => document.querySelector(s);
  const E = U.esc;
  const DEF_CFG = {
    empleado: '', sector: 'SERVICIO TECNICO', sucursal: 'Bahia Blanca', sucursalCorta: 'BB',
    inicio: '08:00', fin: '16:00', cierre: 25, prefijoExtra: true,
    driveUrl: '', driveToken: '', driveRoot: 'Pharmapp', lastBackup: ''
  };
  const SEED_CLIENTS = [
    ['BAHIA BLANCA', 0, 'TALLER'], ['VIEDMA', 278, 'RUF'], ['OLAVARRIA', 298, 'HOSP CURA'],
    ['MAR DEL PLATA', 464, 'TETAMANTI, ALENDE'], ['SANTA ROSA', 327, 'POLYMEDIC, LABSIS'], ['PEDRO LURO', 122, 'PRADERE']
  ];
  const S = { tab: 'dia', day: U.today(), tripId: null, period: null, cfg: null, urls: [], fillMissing: true, installEvt: null };

  // ---------- utilidades de interfaz ----------
  function toast(msg, ms = 2600) {
    const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden');
    clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.add('hidden'), ms);
  }
  function openSheet(title, html, onReady) {
    $('#sheetBox').innerHTML = `<div class="sheet-head"><h2>${E(title)}</h2><button data-act="closeSheet" aria-label="Cerrar">×</button></div>${html}`;
    $('#sheet').classList.remove('hidden');
    $('#sheetBox').scrollTop = 0;
    if (onReady) onReady($('#sheetBox'));
  }
  const closeSheet = () => { $('#sheet').classList.add('hidden'); $('#sheetBox').innerHTML = ''; pendingPhotos = []; };
  const val = id => { const el = document.getElementById(id); return el ? el.value.trim() : ''; };
  const objURL = blob => { const u = URL.createObjectURL(blob); S.urls.push(u); return u; };
  const freeURLs = () => { S.urls.forEach(u => URL.revokeObjectURL(u)); S.urls = []; };
  const ask = msg => window.confirm(msg);

  async function saveCfg() { await DB.put('settings', { key: 'cfg', value: S.cfg }); }

  // ---------- fotos ----------
  let pendingPhotos = []; // fotos tomadas en el formulario abierto, antes de guardar
  async function compress(file) {
    try {
      const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
      const max = 2000, k = Math.min(1, max / Math.max(bmp.width, bmp.height));
      const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
      c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
      return await new Promise(r => c.toBlob(r, 'image/jpeg', 0.82));
    } catch (e) { return file; }
  }
  function pickPhotos(camera) {
    return new Promise(res => {
      const inp = camera ? $('#fileCam') : $('#fileGal');
      inp.value = '';
      inp.onchange = async () => { const out = []; for (const f of inp.files) out.push(await compress(f)); res(out); };
      inp.click();
    });
  }
  function thumbsHTML(existing, pend) {
    const a = (existing || []).map(p => `<div class="th"><img src="${objURL(p.blob)}" alt=""><button class="x" data-act="delPhoto" data-id="${p.id}">×</button><span class="cl">${p.synced ? '☁️' : '⏳'}</span></div>`);
    const b = (pend || []).map((blob, i) => `<div class="th"><img src="${objURL(blob)}" alt=""><button class="x" data-act="delPend" data-i="${i}">×</button></div>`);
    return a.concat(b).join('');
  }
  async function storePhotos(kind, refId) {
    for (const blob of pendingPhotos) await DB.put('photos', { id: DB.id(), kind, refId, blob, created: Date.now(), synced: false });
    pendingPhotos = [];
    syncSoon();
  }
  const photosOf = async refId => (await DB.all('photos')).filter(p => p.refId === refId).sort((a, b) => a.created - b.created);

  // ---------- descarga / compartir ----------
  function downloadBlob(blob, filename) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  }
  let lastFile = null;
  function offerFile(blob, filename, driveFolder) {
    lastFile = { blob, filename, driveFolder };
    const canShare = navigator.canShare && navigator.canShare({ files: [new File([blob], filename, { type: blob.type })] });
    openSheet('Archivo listo', `
      <div class="card"><b>${E(filename)}</b><div class="mute small">${Math.round(blob.size / 1024)} KB</div></div>
      <div class="btns">
        <button class="btn pri big" data-act="fileDownload">⬇️ Descargar</button>
        ${canShare ? '<button class="btn" data-act="fileShare">📤 Compartir (WhatsApp, mail…)</button>' : ''}
        ${Drive.configured(S.cfg) && driveFolder ? '<button class="btn" data-act="fileDrive">☁️ Guardar en Drive</button>' : ''}
      </div>`);
  }
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const xblob = buf => new Blob([buf], { type: XLSX_MIME });

  // ---------- datos ----------
  async function clients() { return (await DB.all('clients')).sort((a, b) => a.localidad.localeCompare(b.localidad)); }
  function clientOptions(list) {
    const out = [];
    for (const c of list) for (const n of String(c.clientes || '').split(',').map(x => x.trim()).filter(Boolean)) out.push({ nombre: n, localidad: c.localidad, km: c.km });
    return out;
  }
  async function daysMap() { const m = {}; for (const d of await DB.all('days')) m[d.date] = d; return m; }
  async function trips() { return (await DB.all('trips')).sort((a, b) => (b.desde || '').localeCompare(a.desde || '')); }
  const tripName = t => `${t.motivo || 'Viaje'} · ${t.localidad || ''}`.trim();

  // Dónde va cada foto en Drive
  async function describePhoto(p) {
    const root = '';
    if (p.kind === 'service') {
      const v = await DB.get('services', p.refId); if (!v) return null;
      const P = U.periodRange(U.periodOf(v.fecha, S.cfg.cierre), S.cfg.cierre);
      const n = (await photosOf(v.id)).findIndex(x => x.id === p.id) + 1;
      return { folder: `${root}${P.key} ${P.mesFin}/Services`, name: U.safe(`${v.fecha} ${v.cliente || ''} ${v.equipo || ''} hoja ${n}`) + '.jpg' };
    }
    if (p.kind === 'ticket') {
      const x = await DB.get('expenses', p.refId); if (!x) return null;
      const t = await DB.get('trips', x.tripId); if (!t) return null;
      const P = U.periodRange(U.periodOf(t.desde, S.cfg.cierre), S.cfg.cierre);
      const conc = (U.CONCEPTOS.find(c => c[0] === x.concepto) || [, 'Otros'])[1];
      return { folder: `${root}${P.key} ${P.mesFin}/Viajes/${U.safe(t.desde + ' ' + (t.localidad || '') + ' - ' + (t.motivo || ''))}`,
        name: U.safe(`${x.fecha} ${conc} ${x.monto} ${x.medio === 'visa' ? 'VISA' : 'personal'} ${p.id.slice(-4)}`) + '.jpg' };
    }
    return null;
  }
  async function backupData() {
    const out = { app: 'pharmapp', version: 1, fecha: new Date().toISOString() };
    for (const s of DB.stores) if (s !== 'photos') out[s] = await DB.all(s);
    out.photos = (await DB.all('photos')).map(({ blob, ...m }) => m);
    return out;
  }
  let syncTimer;
  function syncSoon(ms = 1500) { clearTimeout(syncTimer); syncTimer = setTimeout(() => Drive.sync(S.cfg, describePhoto, backupData), ms); }

  // ---------- estado de Drive en la cabecera ----------
  function paintPill(st) {
    const p = $('#syncPill');
    if (!navigator.onLine) p.textContent = `📴 Sin señal${st.pending ? ` · ${st.pending} foto(s) en cola` : ''}`;
    else if (!Drive.configured(S.cfg)) p.textContent = st.pending ? `⚠️ ${st.pending} foto(s) sin respaldo` : 'Drive sin configurar';
    else if (st.uploading) p.textContent = `⬆️ Subiendo${st.pending ? ` ${st.pending}` : ''}…`;
    else if (st.lastError) p.textContent = '⚠️ Error de Drive';
    else p.textContent = st.pending ? `⏳ ${st.pending} en cola` : '☁️ Drive al día';
  }

  // =====================================================================
  // PESTAÑA DÍA
  // =====================================================================
  function dayStatus(rec, s) {
    if (!rec) return U.isWeekend(s) ? '<span class="status libre">fin de semana</span>' : (s <= U.today() ? '<span class="status falta">sin cargar</span>' : '<span class="status libre">—</span>');
    const t = U.TIPOS[rec.tipo];
    const hs = rec.inicio && rec.fin ? ` ${rec.inicio}–${rec.fin}` : '';
    return `<span class="status ${t.css}">${E(t.label)}${hs}</span>`;
  }
  async function viewDia() {
    const s = S.day, rec = await DB.get('days', s);
    const svcs = (await DB.all('services')).filter(v => v.fecha === s);
    const openTrips = (await trips()).filter(t => t.estado !== 'cerrado');
    const sinSaldo = openTrips.filter(t => t.saldoIda == null);
    const calc = U.calcDay(rec, s, S.cfg);
    const P = U.periodRange(U.periodOf(s, S.cfg.cierre), S.cfg.cierre);
    const dm = await daysMap();
    let h = '';
    if (!S.cfg.empleado) h += `<div class="banner warn">👋 Primero completá tu nombre en <a href="#" data-act="tab" data-tab="mas">Más → Ajustes</a> para que salga en las planillas.</div>`;
    if (S.installEvt) h += `<div class="banner ok">📲 Podés instalar Pharmapp en el celu. <button class="btn pri" style="margin-top:8px" data-act="install">Instalar app</button></div>`;
    for (const t of sinSaldo) h += `<div class="banner warn">💳 <b>Anotá el saldo de la VISA</b> antes de salir a ${E(t.localidad || 'viaje')}. <a href="#" data-act="openTrip" data-id="${t.id}">Anotar ahora</a></div>`;
    h += `<div class="daynav"><button data-act="dayPrev" aria-label="Día anterior">‹</button>
      <label class="d" style="margin:0"><b>${E(U.largo(s))}</b><span class="mute small">${s === U.today() ? 'hoy · ' : ''}toca para cambiar</span>
      <input type="date" id="dayPick" value="${s}" style="position:absolute;opacity:0;width:1px;height:1px"></label>
      <button data-act="dayNext" aria-label="Día siguiente">›</button></div>`;
    if (!rec) {
      h += `<div class="btns">
        <button class="btn ok big" data-act="quickNormal">✅ Jornada normal ${S.cfg.inicio} a ${S.cfg.fin}</button>
        <button class="btn pri big" data-act="editDay" data-tipo="viaje">🚗 Día de viaje</button>
        <button class="btn" data-act="editDay" data-tipo="${U.isWeekend(s) ? 'extra' : 'normal'}">🕒 Otro horario / horas extra</button></div>
        <div class="chips" style="margin-top:10px">
        ${['vacaciones', 'feriado', 'franco', 'licencia'].map(t => `<button class="chip" data-act="quickTipo" data-tipo="${t}">${U.TIPOS[t].label}</button>`).join('')}</div>`;
    } else {
      h += `<div class="card">${dayStatus(rec, s)}
        ${rec.tipo === 'viaje' ? `<div style="margin-top:8px"><b>${E(rec.localidad || '')}</b> ${rec.km ? `· ${rec.km} km` : ''} ${rec.cliente ? `· ${E(rec.cliente)}` : ''}<div class="mute small">Salida ${E(rec.salida || '—')} · Llegada ${E(rec.llegada || '—')}</div></div>` : ''}
        ${calc.extra > 0 ? `<div style="margin-top:8px" class="status extra">+${U.fmtHs(calc.extra)} hs extra</div>` : ''}
        ${rec.obs ? `<div style="margin-top:8px">${E(rec.obs)}</div>` : ''}
        <div class="btns two"><button class="btn" data-act="editDay">✏️ Editar</button><button class="btn bad" data-act="delDay">Borrar</button></div></div>`;
    }
    h += `<h3>Services de este día</h3><div class="card">${svcs.length ? `<ul class="list">${svcs.map(v => `<li data-act="editService" data-id="${v.id}"><div class="grow"><div class="t">${E(v.cliente || 'Sin cliente')}</div><div class="s">${E([v.equipo, v.tipo].filter(Boolean).join(' · '))}</div></div>›</li>`).join('')}</ul>` : '<div class="mute">Ninguno todavía.</div>'}
      <div class="btns"><button class="btn" data-act="newService" data-fecha="${s}">🛠️ Nueva orden de service</button></div></div>`;
    // período
    const pend = P.days.filter(d => d <= U.today() && !dm[d] && !U.isWeekend(d));
    h += `<h3>Período ${U.ddmm(P.start)} al ${U.ddmm(P.end)}</h3>`;
    if (pend.length) h += `<div class="banner warn">Hay ${pend.length} día(s) hábiles sin cargar. <button class="btn" style="margin-top:8px" data-act="fillPending">Marcar todos como jornada normal</button></div>`;
    h += `<div class="card"><ul class="list">${P.days.slice().reverse().filter(d => d <= U.addDays(U.today(), 7)).map(d => `<li data-act="goDay" data-d="${d}"><div class="grow"><div class="t">${U.DIAS[U.dow(d)]} ${U.ddmm(d)}</div></div>${dayStatus(dm[d], d)}</li>`).join('')}</ul></div>`;
    return h;
  }
  async function editDay(tipo) {
    const s = S.day;
    const rec = (await DB.get('days', s)) || {};
    const t = tipo || rec.tipo || 'normal';
    const cl = await clients();
    const opts = clientOptions(cl);
    const openTrips = (await trips()).filter(x => x.estado !== 'cerrado' || (x.desde <= s && (x.hasta || x.desde) >= s));
    const r = Object.assign({ inicio: S.cfg.inicio, fin: S.cfg.fin }, rec);
    openSheet(U.largo(s), `
      <label>Tipo de día</label>
      <select id="dTipo">${Object.entries(U.TIPOS).map(([k, v]) => `<option value="${k}" ${k === t ? 'selected' : ''}>${v.label}</option>`).join('')}</select>
      <div id="dTimes"><div class="grid2"><div><label>Inicio jornada</label><input type="time" id="dIni" value="${E(r.inicio || '')}"></div>
      <div><label>Fin jornada</label><input type="time" id="dFin" value="${E(r.fin || '')}"></div></div></div>
      <div id="dViaje">
        ${openTrips.length ? `<label>Viaje (rendición de gastos)</label><select id="dTrip"><option value="">— ninguno —</option>${openTrips.map(x => `<option value="${x.id}" ${x.id === r.tripId ? 'selected' : ''}>${E(tripName(x))}</option>`).join('')}</select>` : ''}
        <label>Localidad</label><select id="dLoc"><option value="">— elegir —</option>${cl.map(c => `<option value="${E(c.localidad)}" data-km="${c.km ?? ''}" ${c.localidad === r.localidad ? 'selected' : ''}>${E(c.localidad)} (${c.km ?? '?'} km)</option>`).join('')}
          ${r.localidad && !cl.some(c => c.localidad === r.localidad) ? `<option selected>${E(r.localidad)}</option>` : ''}</select>
        <div class="grid2"><div><label>Km</label><input inputmode="numeric" id="dKm" value="${E(r.km ?? '')}"></div>
        <div><label>Cliente / lugar</label><input id="dCli" list="dlCli" value="${E(r.cliente || '')}"></div></div>
        <datalist id="dlCli">${opts.map(o => `<option value="${E(o.nombre)}">${E(o.localidad)}</option>`).join('')}</datalist>
        <div class="grid2"><div><label>Hora de salida</label><input type="time" id="dSal" value="${E(r.salida || r.inicio || S.cfg.inicio)}"></div>
        <div><label>Hora de llegada</label><input type="time" id="dLle" value="${E(r.llegada || '')}"></div></div>
        <div class="mute small" style="margin-top:6px">Si volvés más tarde de las ${S.cfg.fin}, el fin de jornada se ajusta solo a la llegada (podés cambiarlo).</div>
      </div>
      <label>Observaciones (qué hiciste)</label><textarea id="dObs" placeholder="Ej: cambio de jeringas AU700, mapeo de aguja DxH900">${E(rec.obs || '')}</textarea>
      <div id="dCalc" class="mute small" style="margin-top:6px"></div>
      <div class="btns"><button class="btn pri big" data-act="saveDay">Guardar</button></div>`, box => {
      const upd = () => {
        const tp = val('dTipo');
        box.querySelector('#dViaje').style.display = tp === 'viaje' ? '' : 'none';
        box.querySelector('#dTimes').style.display = U.SIN_TRABAJO.includes(tp) ? 'none' : '';
        const c = U.calcDay({ tipo: tp, inicio: val('dIni'), fin: val('dFin') }, s, S.cfg);
        box.querySelector('#dCalc').textContent = c.dur ? `Total ${U.fmtHs(c.dur)} hs${c.extra ? ` · ${U.fmtHs(c.extra)} hs extra` : ''}` : '';
      };
      box.querySelector('#dTipo').onchange = upd;
      box.querySelector('#dIni').oninput = upd; box.querySelector('#dFin').oninput = upd;
      box.querySelector('#dLoc').onchange = e => {
        const o = e.target.selectedOptions[0]; if (o && o.dataset.km) box.querySelector('#dKm').value = o.dataset.km;
        const c = cl.find(x => x.localidad === e.target.value);
        if (c && !val('dCli')) box.querySelector('#dCli').value = String(c.clientes || '').split(',')[0].trim();
      };
      box.querySelector('#dLle').onchange = e => {
        const ll = e.target.value; if (!ll) return;
        if (U.toMin(ll) > U.toMin(S.cfg.fin) || U.toMin(ll) < U.toMin(val('dIni') || S.cfg.inicio)) { box.querySelector('#dFin').value = ll; upd(); }
      };
      const tr = box.querySelector('#dTrip');
      if (tr) tr.onchange = async () => {
        const x = await DB.get('trips', tr.value); if (!x) return;
        if (x.localidad) { box.querySelector('#dLoc').value = x.localidad; box.querySelector('#dLoc').dispatchEvent(new Event('change')); }
        if (x.cliente) box.querySelector('#dCli').value = x.cliente;
      };
      upd();
    });
  }
  async function saveDay() {
    const tipo = val('dTipo');
    const rec = { date: S.day, tipo, updated: Date.now() };
    if (!U.SIN_TRABAJO.includes(tipo)) { rec.inicio = val('dIni'); rec.fin = val('dFin'); }
    if (tipo === 'viaje') {
      Object.assign(rec, { localidad: val('dLoc'), km: U.num(val('dKm')), cliente: val('dCli'), salida: val('dSal'), llegada: val('dLle'), tripId: val('dTrip') || null });
    }
    rec.obs = val('dObs');
    if (!U.SIN_TRABAJO.includes(tipo) && (!rec.inicio || !rec.fin)) return toast('Completá inicio y fin de jornada');
    await DB.put('days', rec); closeSheet(); toast('Día guardado ✓'); render();
  }

  // =====================================================================
  // PESTAÑA VIAJES Y GASTOS
  // =====================================================================
  async function viewViajes() {
    if (S.tripId) return viewTrip(S.tripId);
    const list = await trips();
    const ex = await DB.all('expenses');
    let h = `<h2>Viajes y gastos</h2>
      <div class="banner warn">💳 Antes de salir: <b>anotá el saldo de la VISA</b> al crear el viaje.</div>
      <div class="btns"><button class="btn pri big" data-act="newTrip">➕ Nuevo viaje</button></div>`;
    if (!list.length) return h + '<div class="empty">Todavía no cargaste viajes.</div>';
    const row = t => {
      const tot = ex.filter(e => e.tripId === t.id).reduce((a, e) => a + (Number(e.monto) || 0), 0);
      return `<li data-act="openTrip" data-id="${t.id}"><div class="grow"><div class="t">${E(tripName(t))}</div>
        <div class="s">${U.ddmm(t.desde)}${t.hasta && t.hasta !== t.desde ? ' al ' + U.ddmm(t.hasta) : ''} · ${t.estado === 'cerrado' ? 'terminado' : 'en curso'}${t.saldoIda == null ? ' · ⚠️ falta saldo' : ''}</div></div>
        <span class="amount">${U.money(tot)}</span></li>`;
    };
    const open = list.filter(t => t.estado !== 'cerrado'), done = list.filter(t => t.estado === 'cerrado');
    if (open.length) h += `<h3>En curso</h3><div class="card"><ul class="list">${open.map(row).join('')}</ul></div>`;
    if (done.length) h += `<h3>Terminados</h3><div class="card"><ul class="list">${done.map(row).join('')}</ul></div>`;
    return h;
  }
  async function tripForm(t) {
    const cl = await clients();
    const opts = clientOptions(cl);
    t = t || { desde: U.today(), hasta: U.today() };
    openSheet(t.id ? 'Editar viaje' : 'Nuevo viaje', `
      ${t.id ? '' : '<div class="banner warn">⏰ <b>Recordatorio:</b> antes de salir, fijate el saldo disponible de la VISA de la empresa y anotalo acá abajo.</div>'}
      <label>Motivo (equipo / cliente)</label><input id="tMot" placeholder="Ej: AU700 Cura" value="${E(t.motivo || '')}">
      <label>Lugar de viaje</label><select id="tLoc"><option value="">— elegir —</option>${cl.map(c => `<option ${c.localidad === t.localidad ? 'selected' : ''}>${E(c.localidad)}</option>`).join('')}
        ${t.localidad && !cl.some(c => c.localidad === t.localidad) ? `<option selected>${E(t.localidad)}</option>` : ''}</select>
      <label>Cliente</label><input id="tCli" list="tlCli" value="${E(t.cliente || '')}"><datalist id="tlCli">${opts.map(o => `<option value="${E(o.nombre)}">${E(o.localidad)}</option>`).join('')}</datalist>
      <div class="grid2"><div><label>Desde</label><input type="date" id="tDes" value="${E(t.desde)}"></div><div><label>Hasta</label><input type="date" id="tHas" value="${E(t.hasta || t.desde)}"></div></div>
      <label>💳 Saldo VISA al salir ($)</label><input inputmode="decimal" id="tSal" placeholder="Ej: 333969,17" value="${t.saldoIda ?? ''}">
      <div class="btns"><button class="btn pri big" data-act="saveTrip" data-id="${t.id || ''}">Guardar viaje</button></div>`, box => {
      box.querySelector('#tLoc').onchange = e => { const c = cl.find(x => x.localidad === e.target.value); if (c && !val('tCli')) box.querySelector('#tCli').value = String(c.clientes || '').split(',')[0].trim(); };
    });
  }
  async function saveTrip(id) {
    const t = (id && await DB.get('trips', id)) || { id: DB.id(), estado: 'abierto', created: Date.now() };
    Object.assign(t, { motivo: val('tMot'), localidad: val('tLoc'), cliente: val('tCli'), desde: val('tDes'), hasta: val('tHas') || val('tDes'), saldoIda: U.num(val('tSal')) });
    if (!t.desde) return toast('Poné la fecha del viaje');
    if (t.hasta < t.desde) t.hasta = t.desde;
    await DB.put('trips', t); closeSheet();
    toast(t.saldoIda == null ? '⚠️ Guardado. Acordate de anotar el saldo de la VISA' : 'Viaje guardado ✓', 3500);
    S.tab = 'viajes'; S.tripId = t.id; render();
  }
  async function viewTrip(id) {
    const t = await DB.get('trips', id);
    if (!t) { S.tripId = null; return viewViajes(); }
    const ex = (await DB.byIndex('expenses', 'tripId', id)).sort((a, b) => (a.fecha + a.created).localeCompare(b.fecha + b.created));
    const photos = (await DB.all('photos'));
    const svcs = (await DB.all('services')).filter(v => v.tripId === id);
    const visa = ex.filter(e => e.medio === 'visa').reduce((a, e) => a + (Number(e.monto) || 0), 0);
    const pers = ex.filter(e => e.medio !== 'visa').reduce((a, e) => a + (Number(e.monto) || 0), 0);
    const esperado = t.saldoIda != null ? t.saldoIda - visa : null;
    let h = `<button class="btn ghost" data-act="backTrips" style="width:auto;min-height:36px;padding:6px 12px">‹ Viajes</button>
      <h2>${E(tripName(t))}</h2><div class="mute">${U.ddmmyy(t.desde)}${t.hasta && t.hasta !== t.desde ? ' al ' + U.ddmmyy(t.hasta) : ''}${t.cliente ? ' · ' + E(t.cliente) : ''} · <a href="#" data-act="editTrip" data-id="${id}">editar</a></div>`;
    if (t.saldoIda == null) h += `<div class="banner warn" style="margin-top:10px">💳 <b>Falta el saldo de la VISA al salir.</b> Anotalo antes de gastar. <button class="btn pri" style="margin-top:8px" data-act="editTrip" data-id="${id}">Anotar saldo</button></div>`;
    h += `<div class="card" style="margin-top:10px"><div class="grid2">
      <div><div class="mute small">Saldo al ir</div><b>${t.saldoIda != null ? U.money(t.saldoIda) : '—'}</b></div>
      <div><div class="mute small">Gastos VISA</div><b>${U.money(visa)}</b></div>
      <div><div class="mute small">Saldo esperado</div><b>${esperado != null ? U.money(esperado) : '—'}</b></div>
      <div><div class="mute small">Saldo al volver</div><b>${t.saldoVuelta != null ? U.money(t.saldoVuelta) : '—'}</b></div>
      <div><div class="mute small">Gastos personales</div><b>${U.money(pers)}</b></div></div>
      ${t.saldoVuelta != null && esperado != null && Math.abs(t.saldoVuelta - esperado) > 0.5 ? `<div class="banner bad" style="margin:10px 0 0">No coincide: falta cargar ${U.money(esperado - t.saldoVuelta)} en gastos VISA (o sobra).</div>` : ''}</div>
      <div class="btns"><button class="btn pri big" data-act="newExpense" data-trip="${id}">📸 Agregar gasto / ticket</button></div>`;
    h += `<h3>Gastos (${ex.length})</h3><div class="card">${ex.length ? `<ul class="list">${ex.map(e => {
      const c = (U.CONCEPTOS.find(x => x[0] === e.concepto) || [, 'Otros'])[1];
      const ph = photos.filter(p => p.refId === e.id);
      return `<li data-act="editExpense" data-id="${e.id}"><div class="grow"><div class="t">${E(c)} ${ph.length ? '📷' : ''}</div><div class="s">${U.ddmm(e.fecha)} · ${e.medio === 'visa' ? 'VISA empresa' : 'Personal'}${e.nota ? ' · ' + E(e.nota) : ''}</div></div><span class="amount">${U.money(e.monto)}</span></li>`;
    }).join('')}</ul>` : '<div class="mute">Sin gastos todavía.</div>'}</div>`;
    h += `<h3>Órdenes de service del viaje</h3><div class="card">${svcs.length ? `<ul class="list">${svcs.map(v => `<li data-act="editService" data-id="${v.id}"><div class="grow"><div class="t">${E(v.cliente || '')} · ${E(v.equipo || '')}</div><div class="s">${U.ddmm(v.fecha)} · ${E(v.tipo || '')}</div></div>›</li>`).join('')}</ul>` : '<div class="mute">Ninguna todavía.</div>'}
      <div class="btns"><button class="btn" data-act="newService" data-trip="${id}">🛠️ Agregar hoja de service</button></div></div>`;
    h += `<h3>Al terminar el viaje</h3><div class="btns">
      ${t.estado !== 'cerrado' ? `<button class="btn ok big" data-act="finishTrip" data-id="${id}">🏁 Terminar viaje (anotar saldo al volver)</button>` : ''}
      <button class="btn" data-act="tripDays" data-id="${id}">📅 Cargar los días del viaje en la planilla de horas</button>
      <button class="btn pri" data-act="xlsTrip" data-id="${id}" data-medio="visa">📊 Excel rendición VISA</button>
      ${pers ? `<button class="btn pri" data-act="xlsTrip" data-id="${id}" data-medio="personal">📊 Excel gastos personales</button>` : ''}
      <button class="btn" data-act="zipTrip" data-id="${id}">🗂️ Carpeta del viaje con fotos (.zip)</button>
      ${t.estado === 'cerrado' ? `<button class="btn ghost" data-act="reopenTrip" data-id="${id}">Reabrir viaje</button>` : ''}
      <button class="btn bad" data-act="delTrip" data-id="${id}">Borrar viaje</button></div>`;
    return h;
  }
  async function expenseForm(tripId, e) {
    const t = await DB.get('trips', tripId || e.tripId);
    e = e || { tripId, fecha: U.today() >= t.desde && U.today() <= (t.hasta || t.desde) ? U.today() : t.desde, medio: 'visa', concepto: '' };
    const existing = e.id ? await photosOf(e.id) : [];
    pendingPhotos = [];
    openSheet(e.id ? 'Editar gasto' : 'Nuevo gasto', `
      <div class="btns two"><button class="btn pri" data-act="addPhoto" data-cam="1">📷 Sacar foto del ticket</button><button class="btn" data-act="addPhoto">🖼️ Galería</button></div>
      <div class="thumbs" id="thumbs">${thumbsHTML(existing, [])}</div>
      <label>Concepto</label><div class="chips" id="xConc">${U.CONCEPTOS.map(c => `<button class="chip ${c[0] === e.concepto ? 'on' : ''}" data-act="pickChip" data-v="${c[0]}">${c[1]}</button>`).join('')}</div>
      <label>Monto ($)</label><input inputmode="decimal" id="xMonto" placeholder="0,00" value="${e.monto != null ? String(e.monto).replace('.', ',') : ''}">
      <label>Medio de pago</label><div class="seg" id="xMedio"><button data-act="pickSeg" data-v="visa" class="${e.medio === 'visa' ? 'on' : ''}">💳 VISA empresa</button><button data-act="pickSeg" data-v="personal" class="${e.medio !== 'visa' ? 'on' : ''}">👛 Personal</button></div>
      <label>Fecha</label><input type="date" id="xFecha" value="${E(e.fecha)}">
      <label>Nota (opcional)</label><input id="xNota" value="${E(e.nota || '')}" placeholder="Ej: YPF Azul">
      <div class="btns"><button class="btn pri big" data-act="saveExpense" data-id="${e.id || ''}" data-trip="${t.id}">Guardar gasto</button>
      ${e.id ? `<button class="btn bad" data-act="delExpense" data-id="${e.id}">Borrar gasto</button>` : ''}</div>`);
    expenseForm.existing = existing;
  }
  async function saveExpense(id, tripId) {
    const conc = document.querySelector('#xConc .chip.on');
    const medio = document.querySelector('#xMedio .on');
    const monto = U.num(val('xMonto'));
    if (!conc) return toast('Elegí el concepto');
    if (monto == null) return toast('Poné el monto');
    const e = (id && await DB.get('expenses', id)) || { id: DB.id(), tripId, created: Date.now() };
    Object.assign(e, { concepto: conc.dataset.v, monto, medio: medio ? medio.dataset.v : 'visa', fecha: val('xFecha'), nota: val('xNota') });
    await DB.put('expenses', e);
    await storePhotos('ticket', e.id);
    closeSheet(); toast('Gasto guardado ✓'); render();
  }
  async function finishTrip(id) {
    const t = await DB.get('trips', id);
    const ex = await DB.byIndex('expenses', 'tripId', id);
    const visa = ex.filter(e => e.medio === 'visa').reduce((a, e) => a + (Number(e.monto) || 0), 0);
    openSheet('Terminar viaje', `
      <div class="banner warn">💳 Fijate el saldo de la VISA ahora y anotalo.</div>
      ${t.saldoIda != null ? `<div class="mute">Saldo al ir ${U.money(t.saldoIda)} − gastos VISA ${U.money(visa)} = <b>${U.money(t.saldoIda - visa)}</b> esperado</div>` : ''}
      <label>Saldo VISA al volver ($)</label><input inputmode="decimal" id="fSal" value="${t.saldoVuelta ?? ''}">
      <label>Fecha de vuelta</label><input type="date" id="fHas" value="${E(t.hasta || U.today())}">
      <div class="btns"><button class="btn ok big" data-act="saveFinish" data-id="${id}">Terminar viaje</button></div>`);
  }
  async function tripDays(id) {
    const t = await DB.get('trips', id);
    const cl = (await clients()).find(c => c.localidad === t.localidad);
    let n = 0;
    for (let s = t.desde; s <= (t.hasta || t.desde); s = U.addDays(s, 1)) {
      const old = await DB.get('days', s);
      if (old && old.tipo !== 'normal') continue; // no pisamos lo que ya cargaste a mano
      await DB.put('days', { date: s, tipo: 'viaje', inicio: S.cfg.inicio, fin: S.cfg.fin, salida: s === t.desde ? S.cfg.inicio : '', llegada: '',
        localidad: t.localidad, km: cl ? cl.km : null, cliente: t.cliente || '', tripId: id, obs: old ? old.obs || '' : '', updated: Date.now() });
      n++;
    }
    toast(n ? `${n} día(s) cargados como viaje. Revisá la hora de llegada en la pestaña Día.` : 'Esos días ya estaban cargados', 4000);
  }
  async function tripZip(id) {
    const t = await DB.get('trips', id);
    const ex = await DB.byIndex('expenses', 'tripId', id);
    const zip = new JSZip();
    const base = U.safe(`${t.desde} ${t.localidad || ''} - ${t.motivo || 'viaje'}`);
    const f = zip.folder(base);
    for (const medio of ['visa', 'personal']) {
      if (medio === 'personal' && !ex.some(e => e.medio === 'personal')) continue;
      const r = await XL.buildRendicion({ trip: t, expenses: ex, medio, cfg: S.cfg });
      f.file(r.filename, r.buffer);
    }
    const tk = f.folder('Tickets');
    for (const e of ex) for (const p of await photosOf(e.id)) { const d = await describePhoto(p); tk.file(d ? d.name : p.id + '.jpg', p.blob); }
    const svcs = (await DB.all('services')).filter(v => v.tripId === id);
    if (svcs.length) await addServicesToZip(f.folder('Services'), svcs, `Services del viaje ${tripName(t)}`);
    return { blob: await zip.generateAsync({ type: 'blob' }), filename: base + '.zip' };
  }

  // =====================================================================
  // PESTAÑA SERVICES
  // =====================================================================
  async function viewServices() {
    const list = (await DB.all('services')).sort((a, b) => (b.fecha + b.created).localeCompare(a.fecha + a.created));
    const photos = await DB.all('photos');
    let h = `<h2>Órdenes de service</h2><div class="btns"><button class="btn pri big" data-act="newService">➕ Nueva orden de service</button></div>`;
    if (!list.length) return h + '<div class="empty">Todavía no cargaste órdenes.<br>Cada vez que visites un cliente, cargá la orden y sacale foto a la hoja firmada.</div>';
    const groups = {};
    for (const v of list) (groups[U.periodOf(v.fecha, S.cfg.cierre)] = groups[U.periodOf(v.fecha, S.cfg.cierre)] || []).push(v);
    for (const [k, vs] of Object.entries(groups)) {
      const P = U.periodRange(k, S.cfg.cierre);
      h += `<h3>${P.mesFin} ${P.year} <span class="mute small">(${vs.length})</span></h3><div class="card"><ul class="list">${vs.map(v => {
        const ph = photos.filter(p => p.refId === v.id);
        return `<li data-act="editService" data-id="${v.id}"><div class="grow"><div class="t">${E(v.cliente || 'Sin cliente')} · ${E(v.equipo || '')}</div>
          <div class="s">${U.ddmm(v.fecha)} · ${E(v.tipo || '')} · ${ph.length ? `📷 ${ph.length} ${ph.every(p => p.synced) ? '☁️' : '⏳'}` : '⚠️ sin foto firmada'}</div></div>›</li>`;
      }).join('')}</ul></div>`;
    }
    return h;
  }
  async function serviceForm(v, preset = {}) {
    const cl = await clients();
    const opts = clientOptions(cl);
    const all = await DB.all('services');
    const uniq = k => [...new Set(all.map(x => x[k]).filter(Boolean))];
    const tr = await trips();
    v = v || Object.assign({ fecha: U.today(), solucion: 'SI', tipo: '' }, preset);
    if (!v.id && v.tripId) { const t = tr.find(x => x.id === v.tripId); if (t) { v.cliente = v.cliente || t.cliente || ''; v.localidad = v.localidad || t.localidad || ''; } }
    const existing = v.id ? await photosOf(v.id) : [];
    pendingPhotos = [];
    openSheet(v.id ? 'Orden de service' : 'Nueva orden de service', `
      <div class="grid2"><div><label>Fecha</label><input type="date" id="sFecha" value="${E(v.fecha)}"></div>
      <div><label>Viaje</label><select id="sTrip"><option value="">— ninguno —</option>${tr.slice(0, 30).map(t => `<option value="${t.id}" ${t.id === v.tripId ? 'selected' : ''}>${E(tripName(t))} ${U.ddmm(t.desde)}</option>`).join('')}</select></div></div>
      <label>Cliente</label><input id="sCli" list="slCli" value="${E(v.cliente || '')}"><datalist id="slCli">${opts.map(o => `<option value="${E(o.nombre)}">${E(o.localidad)}</option>`).join('')}</datalist>
      <label>Localidad</label><input id="sLoc" list="slLoc" value="${E(v.localidad || '')}"><datalist id="slLoc">${cl.map(c => `<option value="${E(c.localidad)}">`).join('')}</datalist>
      <div class="grid2"><div><label>Equipo</label><input id="sEq" list="slEq" value="${E(v.equipo || '')}" placeholder="AU700"></div>
      <div><label>Marca</label><input id="sMarca" list="slMarca" value="${E(v.marca || '')}"></div>
      <div><label>Modelo</label><input id="sModelo" list="slModelo" value="${E(v.modelo || '')}"></div>
      <div><label>Nº de Serie</label><input id="sSerie" list="slSerie" value="${E(v.serie || '')}"></div></div>
      <datalist id="slEq">${uniq('equipo').map(x => `<option value="${E(x)}">`).join('')}</datalist><datalist id="slMarca">${uniq('marca').map(x => `<option value="${E(x)}">`).join('')}</datalist>
      <datalist id="slModelo">${uniq('modelo').map(x => `<option value="${E(x)}">`).join('')}</datalist><datalist id="slSerie">${uniq('serie').map(x => `<option value="${E(x)}">`).join('')}</datalist>
      <label>Tipo de servicio</label><select id="sTipo"><option value="">— elegir —</option>${U.TIPOS_SERVICE.map(x => `<option ${x === v.tipo ? 'selected' : ''}>${x}</option>`).join('')}</select>
      <label>Descripción del problema</label><textarea id="sProb">${E(v.problema || '')}</textarea>
      <label>Tareas realizadas</label><textarea id="sTar">${E(v.tareas || '')}</textarea>
      <label>Observaciones</label><textarea id="sObs" style="min-height:50px">${E(v.observaciones || '')}</textarea>
      <label>Solución del problema</label><div class="seg" id="sSol"><button data-act="pickSeg" data-v="SI" class="${v.solucion !== 'NO' ? 'on' : ''}">SI</button><button data-act="pickSeg" data-v="NO" class="${v.solucion === 'NO' ? 'on' : ''}">NO</button></div>
      <label>Piezas reemplazadas</label><textarea id="sPie" style="min-height:50px">${E(v.piezas || '')}</textarea>
      <div class="grid2"><div><label>Tiempo trabajo</label><input id="sTT" value="${E(v.tiempoTrabajo || '')}" placeholder="7,5 hs"></div>
      <div><label>Tiempo viaje</label><input id="sTV" value="${E(v.tiempoViaje || '')}" placeholder="4 hs"></div></div>
      <label>Foto de la hoja firmada</label>
      <div class="btns two"><button class="btn pri" data-act="addPhoto" data-cam="1">📷 Sacar foto</button><button class="btn" data-act="addPhoto">🖼️ Galería</button></div>
      <div class="thumbs" id="thumbs">${thumbsHTML(existing, [])}</div>
      <div class="btns"><button class="btn pri big" data-act="saveService" data-id="${v.id || ''}">Guardar orden</button>
      ${v.id ? `<button class="btn bad" data-act="delService" data-id="${v.id}">Borrar orden</button>` : ''}</div>`, box => {
      box.querySelector('#sSerie').onchange = e => {
        const prev = all.find(x => x.serie && x.serie === e.target.value);
        if (prev) for (const [id, k] of [['sEq', 'equipo'], ['sMarca', 'marca'], ['sModelo', 'modelo'], ['sCli', 'cliente'], ['sLoc', 'localidad']]) { const el = box.querySelector('#' + id); if (!el.value) el.value = prev[k] || ''; }
      };
      box.querySelector('#sCli').onchange = e => { const o = opts.find(x => x.nombre === e.target.value); if (o && !val('sLoc')) box.querySelector('#sLoc').value = o.localidad; };
      box.querySelector('#sTrip').onchange = e => { const t = tr.find(x => x.id === e.target.value); if (t) { if (!val('sCli')) box.querySelector('#sCli').value = t.cliente || ''; if (!val('sLoc')) box.querySelector('#sLoc').value = t.localidad || ''; } };
    });
    serviceForm.existing = existing;
  }
  async function saveService(id) {
    const v = (id && await DB.get('services', id)) || { id: DB.id(), created: Date.now() };
    const sol = document.querySelector('#sSol .on');
    Object.assign(v, { fecha: val('sFecha'), tripId: val('sTrip') || null, cliente: val('sCli'), localidad: val('sLoc'), equipo: val('sEq'), marca: val('sMarca'),
      modelo: val('sModelo'), serie: val('sSerie'), tipo: val('sTipo'), problema: val('sProb'), tareas: val('sTar'), observaciones: val('sObs'),
      solucion: sol ? sol.dataset.v : 'SI', piezas: val('sPie'), tiempoTrabajo: val('sTT'), tiempoViaje: val('sTV') });
    if (!v.fecha) return toast('Poné la fecha');
    const hadPhotos = (serviceForm.existing || []).length + pendingPhotos.length > 0;
    await DB.put('services', v);
    await storePhotos('service', v.id);
    closeSheet(); toast(hadPhotos ? 'Orden guardada ✓' : '⚠️ Guardada sin foto de la hoja firmada', 3500); render();
  }
  async function addServicesToZip(folder, svcs, title) {
    const names = {};
    const trs = await DB.all('trips');
    for (const v of svcs) {
      v.tripName = (trs.find(t => t.id === v.tripId) || {}).motivo || '';
      names[v.id] = [];
      for (const p of await photosOf(v.id)) { const d = await describePhoto(p); const n = d ? d.name : p.id + '.jpg'; folder.file(n, p.blob); names[v.id].push(n); }
    }
    folder.file(U.safe(title) + '.xlsx', await XL.buildServices({ services: svcs, title, photoNames: names }));
  }

  // =====================================================================
  // PESTAÑA MES (cierre mensual)
  // =====================================================================
  async function viewMes() {
    const key = S.period || U.periodOf(U.today(), S.cfg.cierre);
    S.period = key;
    const P = U.periodRange(key, S.cfg.cierre);
    const dm = await daysMap();
    const svcs = (await DB.all('services')).filter(v => v.fecha >= P.start && v.fecha <= P.end);
    const trs = (await trips()).filter(t => t.desde <= P.end && (t.hasta || t.desde) >= P.start);
    const photos = await DB.all('photos');
    let extra = 0, km = 0, trab = 0;
    const svcBy = {}; for (const v of svcs) (svcBy[v.fecha] = svcBy[v.fecha] || []).push(v);
    const rows = P.days.map(s => XL.resolveDay(s, dm[s], S.cfg, S.fillMissing, svcBy[s]));
    for (const r of rows) { extra += r.extra; if (r.inicio) trab++; if (r.viaje === 'SI' && r.km) km += r.km * 2; }
    const sinCargar = P.days.filter(s => !dm[s] && !U.isWeekend(s));
    const pendFotos = photos.filter(p => !p.synced).length;
    const sinFoto = svcs.filter(v => !photos.some(p => p.refId === v.id)).length;
    let h = `<div class="daynav"><button data-act="perPrev">‹</button><div class="d"><b>${P.mesFin} ${P.year}</b><span class="mute small">del ${U.ddmmyy(P.start)} al ${U.ddmmyy(P.end)}</span></div><button data-act="perNext">›</button></div>
      <div class="kpis"><div class="kpi"><b>${trab}</b><span>días trabajados</span></div><div class="kpi"><b>${U.fmtHs(extra)}</b><span>hs extra</span></div><div class="kpi"><b>${trs.length}</b><span>viajes</span></div>
      <div class="kpi"><b>${svcs.length}</b><span>services</span></div><div class="kpi"><b>${km}</b><span>km (ida y vuelta)</span></div><div class="kpi"><b>${pendFotos}</b><span>fotos sin subir</span></div></div>`;
    if (sinFoto) h += `<div class="banner warn">${sinFoto} orden(es) de service sin foto de la hoja firmada.</div>`;
    if (sinCargar.length) h += `<div class="card"><label style="margin:0;display:flex;gap:10px;align-items:center;color:var(--ink);font-size:15px"><input type="checkbox" id="fillMissing" style="width:22px;height:22px" ${S.fillMissing ? 'checked' : ''}>
      Completar los ${sinCargar.length} días hábiles sin cargar como jornada normal (${S.cfg.inicio} a ${S.cfg.fin})</label>
      <div class="mute small" style="margin-top:6px">${sinCargar.map(U.ddmm).join(', ')}</div></div>`;
    h += `<div class="btns"><button class="btn pri big" data-act="genHoras">📊 Generar reporte (planilla de horas)</button>
      <button class="btn" data-act="zipMes">🗂️ Carpeta de cierre del mes (.zip)</button>
      ${Drive.configured(S.cfg) ? '<button class="btn" data-act="driveMes">☁️ Guardar cierre en Drive</button>' : ''}</div>`;
    h += `<h3>Vista previa</h3><div class="card"><ul class="list">${rows.map(r => `<li data-act="goDay" data-d="${r.date}"><div class="grow"><div class="t">${U.DIAS[U.dow(r.date)]} ${U.ddmm(r.date)} ${r.viaje === 'SI' ? '🚗' : ''}</div>
      <div class="s">${r.inicio ? `${r.inicio}–${r.fin}` : ''}${r.localidad ? ' · ' + E(r.localidad) : ''}${r.obs ? ' · ' + E(r.obs) : ''}${r.auto ? ' · (auto)' : ''}</div></div>
      ${r.extra ? `<span class="status extra">+${U.fmtHs(r.extra)}</span>` : (r.missing && !r.auto ? '<span class="status falta">vacío</span>' : '')}</li>`).join('')}</ul></div>`;
    return h;
  }
  async function genHoras() {
    if (!S.cfg.empleado) { toast('Primero poné tu nombre en Más → Ajustes'); S.tab = 'mas'; return render(); }
    const P = U.periodRange(S.period, S.cfg.cierre);
    const svcs = (await DB.all('services')).filter(v => v.fecha >= P.start && v.fecha <= P.end);
    return XL.buildHoras({ periodKey: S.period, days: await daysMap(), services: svcs, clients: await clients(), cfg: S.cfg, fillMissing: S.fillMissing });
  }
  async function zipMes() {
    const P = U.periodRange(S.period, S.cfg.cierre);
    const zip = new JSZip();
    const base = U.safe(`Cierre ${P.key} ${P.mesFin} ${P.year}`);
    const f = zip.folder(base);
    const h = await genHoras(); if (!h) return null;
    f.file(h.filename, h.buffer);
    const svcs = (await DB.all('services')).filter(v => v.fecha >= P.start && v.fecha <= P.end).sort((a, b) => a.fecha.localeCompare(b.fecha));
    await addServicesToZip(f.folder('Services'), svcs, `Resumen services ${P.mesFin} ${P.year}`);
    const trs = (await trips()).filter(t => t.desde >= P.start && t.desde <= P.end);
    for (const t of trs) {
      const tf = f.folder('Viajes').folder(U.safe(`${t.desde} ${t.localidad || ''} - ${t.motivo || ''}`));
      const ex = await DB.byIndex('expenses', 'tripId', t.id);
      for (const medio of ['visa', 'personal']) {
        if (medio === 'personal' && !ex.some(e => e.medio === 'personal')) continue;
        const r = await XL.buildRendicion({ trip: t, expenses: ex, medio, cfg: S.cfg }); tf.file(r.filename, r.buffer);
      }
      for (const e of ex) for (const p of await photosOf(e.id)) { const d = await describePhoto(p); tf.folder('Tickets').file(d ? d.name : p.id + '.jpg', p.blob); }
    }
    return { blob: await zip.generateAsync({ type: 'blob' }), filename: base + '.zip', P, horas: h };
  }

  // =====================================================================
  // PESTAÑA MÁS (clientes, ajustes, Drive, copia de seguridad)
  // =====================================================================
  async function viewMas() {
    const cl = await clients();
    const st = Drive.state();
    const c = S.cfg;
    let h = `<h2>Clientes y localidades</h2><div class="card"><ul class="list">${cl.map(x => `<li data-act="editClient" data-id="${x.id}"><div class="grow"><div class="t">${E(x.localidad)}</div><div class="s">${E(x.clientes || '')}</div></div><span class="amount">${x.km ?? '?'} km</span></li>`).join('')}</ul>
      <div class="btns"><button class="btn" data-act="editClient">➕ Agregar localidad / cliente</button></div></div>`;
    h += `<h2>Ajustes</h2><div class="card">
      <label>Empleado/a (como figura en la planilla)</label><input id="cEmp" value="${E(c.empleado)}" placeholder="Apellido Nombre">
      <div class="grid2"><div><label>Sector</label><input id="cSec" value="${E(c.sector)}"></div><div><label>Sucursal</label><input id="cSuc" value="${E(c.sucursal)}"></div>
      <div><label>Jornada desde</label><input type="time" id="cIni" value="${E(c.inicio)}"></div><div><label>Jornada hasta</label><input type="time" id="cFin" value="${E(c.fin)}"></div>
      <div><label>Día de cierre del período</label><input inputmode="numeric" id="cCie" value="${E(c.cierre)}"></div><div><label>Sigla sucursal (archivo)</label><input id="cSig" value="${E(c.sucursalCorta)}"></div></div>
      <label style="display:flex;gap:10px;align-items:center;color:var(--ink)"><input type="checkbox" id="cPre" style="width:22px;height:22px" ${c.prefijoExtra ? 'checked' : ''}> Anteponer "X hs extra:" en observaciones</label>
      <div class="btns"><button class="btn pri" data-act="saveCfg">Guardar ajustes</button></div></div>`;
    h += `<h2>Respaldo en Google Drive</h2><div class="card">
      ${Drive.configured(c) ? `<div class="banner ${st.lastError ? 'bad' : 'ok'}">${st.lastError ? '⚠️ ' + E(st.lastError) : '✅ Conectado.'} Fotos en cola: ${st.pending}.</div>
        <div class="btns two"><button class="btn pri" data-act="syncNow">Subir ahora</button><button class="btn" data-act="testDrive">Probar conexión</button></div>`
        : `<div class="mute">Las fotos se guardan en el celu y, cuando hay internet, se suben solas a una carpeta "Pharmapp" de tu Drive, ordenadas por mes. Se configura una sola vez.</div>`}
      <div class="btns"><button class="btn" data-act="driveSetup">${Drive.configured(c) ? 'Ver configuración' : '⚙️ Configurar Drive'}</button></div></div>`;
    h += `<h2>Copia de seguridad</h2><div class="card"><div class="mute small">Todo se guarda en este celu. Además de Drive, podés bajar una copia de tus datos (sin fotos) o restaurarla en otro celu.</div>
      <div class="btns two"><button class="btn" data-act="exportJson">⬇️ Exportar datos</button><button class="btn" data-act="importJson">⬆️ Restaurar</button></div></div>
      <div class="mute small" style="text-align:center;margin:16px 0">Pharmapp v2 · funciona sin internet</div>`;
    return h;
  }
  async function clientForm(id) {
    const c = (id && await DB.get('clients', id)) || {};
    openSheet(id ? 'Editar localidad' : 'Nueva localidad', `
      <label>Localidad</label><input id="lLoc" value="${E(c.localidad || '')}" placeholder="Ej: TRES ARROYOS">
      <label>Km desde la base</label><input inputmode="numeric" id="lKm" value="${E(c.km ?? '')}">
      <label>Clientes / lugares (separados por coma)</label><textarea id="lCli" style="min-height:60px">${E(c.clientes || '')}</textarea>
      <div class="btns"><button class="btn pri big" data-act="saveClient" data-id="${id || ''}">Guardar</button>${id ? `<button class="btn bad" data-act="delClient" data-id="${id}">Borrar</button>` : ''}</div>`);
  }
  function driveSetup() {
    if (!S.cfg.driveToken) { S.cfg.driveToken = Array.from(crypto.getRandomValues(new Uint8Array(12))).map(b => b.toString(16).padStart(2, '0')).join(''); saveCfg(); }
    const code = Drive.scriptCode(S.cfg.driveToken, S.cfg.driveRoot);
    openSheet('Configurar Google Drive', `
      <div class="mute">Se hace una sola vez, mejor desde la compu:</div>
      <ol style="padding-left:20px;margin:8px 0">
        <li>Entrá a <b>script.google.com</b> con tu cuenta de Google y tocá <b>Nuevo proyecto</b>.</li>
        <li>Borrá lo que hay y pegá el código de abajo (botón Copiar). Guardá.</li>
        <li>Arriba a la derecha: <b>Implementar → Nueva implementación → Aplicación web</b>. "Ejecutar como": <b>Yo</b>. "Quién tiene acceso": <b>Cualquier usuario</b> (ojo: no "con cuenta de Google"). Implementar y aceptar los permisos (si dice "Google no verificó esta app", tocá Configuración avanzada → Ir a … y Permitir).</li>
        <li>Copiá la <b>URL de la aplicación web</b> y pegala acá abajo.</li></ol>
      <pre class="code">${E(code)}</pre>
      <div class="btns two"><button class="btn" data-act="copyScript">📋 Copiar código</button><button class="btn" data-act="shareScript">📤 Enviarme el código</button></div>
      <label>URL de la aplicación web</label><input id="dUrl" value="${E(S.cfg.driveUrl)}" placeholder="https://script.google.com/macros/s/…/exec">
      <div class="btns"><button class="btn pri big" data-act="saveDrive">Guardar y probar</button></div>
      <div id="dErr" class="banner bad hidden" style="margin-top:10px"></div>
      <div class="mute small">El código lleva una clave secreta propia de este celu para que nadie más pueda subir archivos a tu Drive.</div>`);
  }

  // =====================================================================
  // RENDER Y EVENTOS
  // =====================================================================
  async function render() {
    freeURLs();
    document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === S.tab));
    const v = { dia: viewDia, viajes: viewViajes, services: viewServices, mes: viewMes, mas: viewMas }[S.tab];
    try { $('#view').innerHTML = await v(); } catch (e) { console.error(e); $('#view').innerHTML = `<div class="banner bad">Error: ${E(e.message)}</div>`; }
    const dp = $('#dayPick'); if (dp) dp.onchange = () => { if (dp.value) { S.day = dp.value; render(); } };
    const fm = $('#fillMissing'); if (fm) fm.onchange = () => { S.fillMissing = fm.checked; render(); };
    Drive.sync && paintPill(Drive.state());
  }
  async function refreshThumbs(existingGetter) {
    const box = $('#thumbs'); if (!box) return;
    box.innerHTML = thumbsHTML(existingGetter ? await existingGetter() : [], pendingPhotos);
  }
  let currentRefId = null;

  const actions = {
    closeSheet,
    tab: el => { S.tab = el.dataset.tab; S.tripId = null; render(); },
    install: async () => { if (S.installEvt) { S.installEvt.prompt(); await S.installEvt.userChoice; S.installEvt = null; render(); } },
    dayPrev: () => { S.day = U.addDays(S.day, -1); render(); },
    dayNext: () => { S.day = U.addDays(S.day, 1); render(); },
    goDay: el => { S.day = el.dataset.d; S.tab = 'dia'; render(); window.scrollTo(0, 0); },
    quickNormal: async () => { await DB.put('days', { date: S.day, tipo: 'normal', inicio: S.cfg.inicio, fin: S.cfg.fin, obs: '', updated: Date.now() }); toast('Jornada normal cargada ✓'); render(); },
    quickTipo: async el => { await DB.put('days', { date: S.day, tipo: el.dataset.tipo, updated: Date.now() }); toast(U.TIPOS[el.dataset.tipo].label + ' ✓'); render(); },
    editDay: el => editDay(el.dataset.tipo),
    saveDay,
    delDay: async () => { if (ask('¿Borrar lo cargado en este día?')) { await DB.del('days', S.day); render(); } },
    fillPending: async () => {
      const P = U.periodRange(U.periodOf(S.day, S.cfg.cierre), S.cfg.cierre); const dm = await daysMap(); let n = 0;
      for (const d of P.days) if (d <= U.today() && !dm[d] && !U.isWeekend(d)) { await DB.put('days', { date: d, tipo: 'normal', inicio: S.cfg.inicio, fin: S.cfg.fin, updated: Date.now() }); n++; }
      toast(`${n} días cargados ✓`); render();
    },
    newTrip: () => tripForm(),
    editTrip: async el => tripForm(await DB.get('trips', el.dataset.id)),
    saveTrip: el => saveTrip(el.dataset.id),
    openTrip: el => { S.tab = 'viajes'; S.tripId = el.dataset.id; render(); window.scrollTo(0, 0); },
    backTrips: () => { S.tripId = null; render(); },
    newExpense: el => { currentRefId = null; expenseForm(el.dataset.trip); },
    editExpense: async el => { currentRefId = el.dataset.id; expenseForm(null, await DB.get('expenses', el.dataset.id)); },
    saveExpense: el => saveExpense(el.dataset.id, el.dataset.trip),
    delExpense: async el => { if (!ask('¿Borrar este gasto?')) return; await DB.del('expenses', el.dataset.id); for (const p of await photosOf(el.dataset.id)) if (!p.synced) await DB.del('photos', p.id); closeSheet(); render(); },
    finishTrip: el => finishTrip(el.dataset.id),
    saveFinish: async el => {
      const t = await DB.get('trips', el.dataset.id); t.saldoVuelta = U.num(val('fSal')); t.hasta = val('fHas') || t.hasta; t.estado = 'cerrado'; t.presentacion = U.today();
      if (t.saldoVuelta == null && !ask('No anotaste el saldo al volver. ¿Terminar igual?')) return;
      await DB.put('trips', t); closeSheet(); toast('Viaje terminado ✓ Ya podés bajar la rendición'); render();
    },
    reopenTrip: async el => { const t = await DB.get('trips', el.dataset.id); t.estado = 'abierto'; await DB.put('trips', t); render(); },
    delTrip: async el => {
      if (!ask('¿Borrar el viaje con todos sus gastos? Las fotos ya subidas a Drive quedan allá.')) return;
      for (const e of await DB.byIndex('expenses', 'tripId', el.dataset.id)) { await DB.del('expenses', e.id); for (const p of await photosOf(e.id)) await DB.del('photos', p.id); }
      await DB.del('trips', el.dataset.id); S.tripId = null; render();
    },
    tripDays: el => tripDays(el.dataset.id),
    xlsTrip: async el => {
      const t = await DB.get('trips', el.dataset.id); const ex = await DB.byIndex('expenses', 'tripId', t.id);
      if (!S.cfg.empleado) toast('Tip: poné tu nombre en Más → Ajustes');
      if (el.dataset.medio === 'visa' && t.saldoIda == null) toast('⚠️ Falta el saldo VISA al salir');
      const r = await XL.buildRendicion({ trip: t, expenses: ex, medio: el.dataset.medio, cfg: S.cfg });
      if (r.extraCols) toast('Hubo gastos en más de 8 días: los últimos van juntos en "Otros días"', 4000);
      const P = U.periodRange(U.periodOf(t.desde, S.cfg.cierre), S.cfg.cierre);
      offerFile(xblob(r.buffer), r.filename, `${P.key} ${P.mesFin}/Viajes/${U.safe(t.desde + ' ' + (t.localidad || '') + ' - ' + (t.motivo || ''))}`);
    },
    zipTrip: async el => { toast('Armando carpeta…'); const r = await tripZip(el.dataset.id); offerFile(r.blob, r.filename, null); },
    newService: el => serviceForm(null, { fecha: el.dataset.fecha || U.today(), tripId: el.dataset.trip || null }),
    editService: async el => { currentRefId = el.dataset.id; serviceForm(await DB.get('services', el.dataset.id)); },
    saveService: el => saveService(el.dataset.id),
    delService: async el => { if (!ask('¿Borrar esta orden de service?')) return; await DB.del('services', el.dataset.id); for (const p of await photosOf(el.dataset.id)) if (!p.synced) await DB.del('photos', p.id); closeSheet(); render(); },
    addPhoto: async el => { const got = await pickPhotos(!!el.dataset.cam); pendingPhotos.push(...got); refreshThumbs(currentRefId ? () => photosOf(currentRefId) : null); },
    delPend: el => { pendingPhotos.splice(Number(el.dataset.i), 1); refreshThumbs(currentRefId ? () => photosOf(currentRefId) : null); },
    delPhoto: async el => { if (!ask('¿Quitar esta foto? (si ya se subió, en Drive queda)')) return; await DB.del('photos', el.dataset.id); refreshThumbs(currentRefId ? () => photosOf(currentRefId) : null); },
    pickChip: el => { el.parentElement.querySelectorAll('.chip').forEach(c => c.classList.remove('on')); el.classList.add('on'); },
    pickSeg: el => { el.parentElement.querySelectorAll('button').forEach(c => c.classList.remove('on')); el.classList.add('on'); },
    perPrev: () => { S.period = U.shiftPeriod(S.period, -1); render(); },
    perNext: () => { S.period = U.shiftPeriod(S.period, 1); render(); },
    genHoras: async () => { toast('Generando planilla…'); const r = await genHoras(); if (r) { const P = r.period; offerFile(xblob(r.buffer), r.filename, `${P.key} ${P.mesFin}/Cierre`); } },
    zipMes: async () => { toast('Armando carpeta del mes…', 5000); const r = await zipMes(); if (r) offerFile(r.blob, r.filename, `${r.P.key} ${r.P.mesFin}/Cierre`); },
    driveMes: async () => {
      if (!navigator.onLine) return toast('Sin internet: probá cuando tengas señal');
      toast('Subiendo cierre a Drive…', 8000);
      try {
        const P = U.periodRange(S.period, S.cfg.cierre); const folder = `${P.key} ${P.mesFin}/Cierre`;
        const h = await genHoras(); if (!h) return;
        await Drive.upload(S.cfg, folder, h.filename, xblob(h.buffer), true);
        const svcs = (await DB.all('services')).filter(v => v.fecha >= P.start && v.fecha <= P.end);
        const trs = await DB.all('trips');
        svcs.forEach(v => { v.tripName = (trs.find(t => t.id === v.tripId) || {}).motivo || ''; });
        const names = {}; for (const v of svcs) names[v.id] = (await photosOf(v.id)).map(p => p.driveName || '').filter(Boolean);
        await Drive.upload(S.cfg, folder, U.safe(`Resumen services ${P.mesFin} ${P.year}`) + '.xlsx', xblob(await XL.buildServices({ services: svcs, title: `Resumen services ${P.mesFin} ${P.year}`, photoNames: names })), true);
        await Drive.sync(S.cfg, describePhoto, backupData);
        toast(`✅ Cierre guardado en Drive: Pharmapp/${folder}`, 5000);
      } catch (e) { toast('Error: ' + e.message, 5000); }
    },
    fileDownload: () => { downloadBlob(lastFile.blob, lastFile.filename); },
    fileShare: async () => { try { await navigator.share({ files: [new File([lastFile.blob], lastFile.filename, { type: lastFile.blob.type })], title: lastFile.filename }); } catch (e) { /* cancelado */ } },
    fileDrive: async () => { try { toast('Subiendo…'); await Drive.upload(S.cfg, lastFile.driveFolder, lastFile.filename, lastFile.blob, true); toast('✅ Guardado en Drive: Pharmapp/' + lastFile.driveFolder, 4000); } catch (e) { toast('Error: ' + e.message, 5000); } },
    editClient: el => clientForm(el.dataset.id),
    saveClient: async el => {
      const c = (el.dataset.id && await DB.get('clients', el.dataset.id)) || { id: DB.id() };
      Object.assign(c, { localidad: val('lLoc').toUpperCase(), km: U.num(val('lKm')), clientes: val('lCli').toUpperCase() });
      if (!c.localidad) return toast('Poné la localidad');
      await DB.put('clients', c); closeSheet(); render();
    },
    delClient: async el => { if (ask('¿Borrar esta localidad?')) { await DB.del('clients', el.dataset.id); closeSheet(); render(); } },
    saveCfg: async () => {
      Object.assign(S.cfg, { empleado: val('cEmp'), sector: val('cSec'), sucursal: val('cSuc'), inicio: val('cIni') || '08:00', fin: val('cFin') || '16:00',
        cierre: Math.min(28, Math.max(1, Number(val('cCie')) || 25)), sucursalCorta: val('cSig'), prefijoExtra: $('#cPre').checked });
      await saveCfg(); toast('Ajustes guardados ✓'); render();
    },
    driveSetup,
    copyScript: async () => { try { await navigator.clipboard.writeText(Drive.scriptCode(S.cfg.driveToken, S.cfg.driveRoot)); toast('Código copiado ✓'); } catch (e) { toast('No se pudo copiar; usá "Enviarme el código"'); } },
    shareScript: async () => { const t = Drive.scriptCode(S.cfg.driveToken, S.cfg.driveRoot); if (navigator.share) { try { await navigator.share({ title: 'Código Pharmapp Drive', text: t }); } catch (e) { } } else downloadBlob(new Blob([t], { type: 'text/plain' }), 'pharmapp-drive.gs.txt'); },
    saveDrive: async () => {
      const url = val('dUrl');
      if (!/^https:\/\/script\.google\.com\/.*\/exec$/.test(url)) {
        return toast(/\/dev$/.test(url) ? 'Esa es la URL de prueba (/dev). Usá la de "Nueva implementación", que termina en /exec' : 'La URL tiene que empezar con https://script.google.com/ y terminar en /exec', 6000);
      }
      S.cfg.driveUrl = url; await saveCfg();
      try { await Drive.call(S.cfg, { action: 'ping' }); closeSheet(); toast('✅ Drive conectado'); syncSoon(100); render(); }
      catch (e) {
        const box = document.getElementById('dErr');
        const msg = /fetch|network|json/i.test(e.message)
          ? `<b>Google no dejó conectar.</b> Casi siempre es por la configuración de la implementación:<ol style="padding-left:20px;margin:6px 0">
             <li>En el script: <b>Implementar → Gestionar implementaciones → ✏️ editar</b>.</li>
             <li>"Ejecutar como": <b>Yo</b>. "Quién tiene acceso": <b>Cualquier usuario</b> (no "Cualquier usuario con cuenta de Google").</li>
             <li>En "Versión" elegí <b>Nueva versión</b> e Implementar.</li></ol>
             Para revisarlo, abrí la URL en el navegador: si ves <code>{"ok":true,"app":"pharmapp"}</code> está bien; si te pide iniciar sesión o permisos, falta el paso 2.
             <a class="btn" style="margin-top:8px" href="${E(url)}" target="_blank" rel="noopener">Abrir la URL para revisar</a>`
          : E(e.message);
        if (box) { box.innerHTML = msg; box.classList.remove('hidden'); } else toast('No se pudo conectar: ' + e.message, 6000);
      }
    },
    syncNow: () => { if (!navigator.onLine) return toast('Sin internet'); Drive.sync(S.cfg, describePhoto, backupData).then(render); },
    testDrive: async () => { try { await Drive.call(S.cfg, { action: 'ping' }); toast('✅ Conexión OK'); } catch (e) { toast('Error: ' + e.message, 5000); } },
    exportJson: async () => downloadBlob(new Blob([JSON.stringify(await backupData(), null, 1)], { type: 'application/json' }), `pharmapp-datos-${U.today()}.json`),
    importJson: () => {
      const inp = $('#fileJson'); inp.value = '';
      inp.onchange = async () => {
        try {
          const data = JSON.parse(await inp.files[0].text());
          if (data.app !== 'pharmapp') throw new Error('No es una copia de Pharmapp');
          if (!ask('Se van a agregar/reemplazar los datos con los de la copia. ¿Seguir?')) return;
          for (const s of DB.stores) if (s !== 'photos' && Array.isArray(data[s])) for (const r of data[s]) await DB.put(s, r);
          S.cfg = Object.assign({}, DEF_CFG, (await DB.get('settings', 'cfg') || {}).value);
          toast('Datos restaurados ✓'); render();
        } catch (e) { toast('Error: ' + e.message, 5000); }
      };
      inp.click();
    }
  };

  document.addEventListener('click', ev => {
    const el = ev.target.closest('[data-act]');
    if (el) {
      ev.preventDefault();
      const f = actions[el.dataset.act];
      if (f) Promise.resolve(f(el)).catch(e => { console.error(e); toast('Error: ' + e.message, 5000); });
      return;
    }
    const tb = ev.target.closest('.tabs button');
    if (tb) { S.tab = tb.dataset.tab; S.tripId = null; currentRefId = null; render(); window.scrollTo(0, 0); }
  });
  $('#sheet').addEventListener('click', ev => { if (ev.target.id === 'sheet') closeSheet(); });
  $('#syncPill').addEventListener('click', () => { S.tab = 'mas'; S.tripId = null; render(); });
  // al cerrar el formulario se olvida a qué registro pertenecían las fotos
  const _close = actions.closeSheet; actions.closeSheet = () => { currentRefId = null; _close(); };

  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); S.installEvt = e; if (S.tab === 'dia') render(); });
  window.addEventListener('online', () => { paintPill(Drive.state()); syncSoon(500); });
  window.addEventListener('offline', () => paintPill(Drive.state()));
  document.addEventListener('visibilitychange', () => { if (!document.hidden) syncSoon(800); });

  async function init() {
    const c = await DB.get('settings', 'cfg');
    S.cfg = Object.assign({}, DEF_CFG, c ? c.value : {});
    if (!c) await saveCfg();
    if (!(await DB.all('clients')).length) for (const [localidad, km, clientes] of SEED_CLIENTS) await DB.put('clients', { id: DB.id(), localidad, km, clientes });
    Drive.onChange(paintPill);
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
    if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
    await render();
    syncSoon(2000);
    setInterval(() => syncSoon(0), 120000);
  }
  init();
})();

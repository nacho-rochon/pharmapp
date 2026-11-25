// Respaldo en Google Drive mediante un "Apps Script" propio en la cuenta de Google del usuario.
// Las fotos se guardan primero en el celu y se suben solas cuando hay internet.
const Drive = (() => {
  let busy = false;
  const listeners = new Set();
  let state = { pending: 0, uploading: false, lastError: '', lastOk: null };
  const emit = () => listeners.forEach(f => f(state));

  function scriptCode(token, root) {
    return `// Pharmapp – respaldo en Google Drive. Pegá este código en script.google.com
const TOKEN = '${token}';
const ROOT = '${root || 'Pharmapp'}';

function doGet() { return out({ ok: true, app: 'pharmapp' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) return out({ ok: false, error: 'token' });
    if (req.action === 'ping') return out({ ok: true });
    const folder = getFolder(ROOT + '/' + (req.folder || ''));
    if (req.action === 'upload') {
      const old = folder.getFilesByName(req.name);
      while (old.hasNext()) {
        const f = old.next();
        if (!req.overwrite) return out({ ok: true, id: f.getId(), dup: true });
        f.setTrashed(true);
      }
      const blob = Utilities.newBlob(Utilities.base64Decode(req.data), req.mime, req.name);
      const file = folder.createFile(blob);
      return out({ ok: true, id: file.getId(), url: folder.getUrl() });
    }
    return out({ ok: false, error: 'accion' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getFolder(path) {
  let f = DriveApp.getRootFolder();
  path.split('/').filter(String).forEach(function (p) {
    const it = f.getFoldersByName(p);
    f = it.hasNext() ? it.next() : f.createFolder(p);
  });
  return f;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
`;
  }

  const b64 = blob => new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(',')[1]);
    r.onerror = () => rej(r.error);
    r.readAsDataURL(blob);
  });

  async function call(cfg, body) {
    const res = await fetch(cfg.driveUrl, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ token: cfg.driveToken }, body))
    });
    const j = await res.json();
    if (!j.ok) throw new Error(j.error === 'token' ? 'El código secreto no coincide con el del script' : (j.error || 'Error de Drive'));
    return j;
  }

  async function upload(cfg, folder, name, blob, overwrite = false) {
    return call(cfg, { action: 'upload', folder, name, mime: blob.type || 'application/octet-stream', data: await b64(blob), overwrite });
  }

  const configured = cfg => !!(cfg && cfg.driveUrl && cfg.driveToken);

  // describe(photo) -> { folder, name } lo resuelve la app según a qué pertenece la foto
  async function sync(cfg, describe, backupData) {
    const pending = (await DB.all('photos')).filter(p => !p.synced);
    state.pending = pending.length; emit();
    if (busy || !configured(cfg) || !navigator.onLine) return;
    busy = true; state.uploading = true; state.lastError = ''; emit();
    try {
      for (const p of pending) {
        const d = await describe(p);
        if (!d) continue; // todavía no pertenece a nada
        const r = await upload(cfg, d.folder, d.name, p.blob);
        p.synced = true; p.driveId = r.id; p.driveName = d.name; p.driveFolder = d.folder;
        await DB.put('photos', p);
        state.pending = Math.max(0, state.pending - 1); emit();
      }
      // copia de seguridad diaria de los datos (sin fotos)
      const hoy = U.today();
      if (backupData && cfg.lastBackup !== hoy) {
        const data = await backupData();
        await upload(cfg, 'Copias de seguridad', `pharmapp-datos-${hoy}.json`, new Blob([JSON.stringify(data)], { type: 'application/json' }), true);
        cfg.lastBackup = hoy; await DB.put('settings', { key: 'cfg', value: cfg });
      }
      state.lastOk = new Date();
    } catch (e) {
      state.lastError = e.message || String(e);
    } finally {
      busy = false; state.uploading = false; emit();
    }
  }

  return { scriptCode, upload, call, sync, configured, onChange: f => { listeners.add(f); f(state); }, state: () => state };
})();

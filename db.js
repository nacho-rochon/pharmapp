// Base de datos local (IndexedDB): todo queda guardado en el celu, con o sin internet.
const DB = (() => {
  const NAME = 'pharmapp';
  const STORES = {
    days: 'date',        // un registro por día: { date:'2026-11-24', tipo, inicio, fin, ... }
    clients: 'id',       // localidades y clientes
    trips: 'id',         // viajes (rendición de gastos)
    expenses: 'id',      // gastos de un viaje
    services: 'id',      // órdenes de service
    photos: 'id',        // fotos (tickets y hojas firmadas) + estado de subida a Drive
    settings: 'key'
  };
  let dbp;
  function open() {
    if (dbp) return dbp;
    dbp = new Promise((res, rej) => {
      const r = indexedDB.open(NAME, 1);
      r.onupgradeneeded = () => {
        const db = r.result;
        for (const [s, key] of Object.entries(STORES)) {
          if (!db.objectStoreNames.contains(s)) {
            const os = db.createObjectStore(s, { keyPath: key });
            if (s === 'expenses') os.createIndex('tripId', 'tripId');
            if (s === 'photos') os.createIndex('synced', 'synced');
          }
        }
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return dbp;
  }
  async function tx(store, mode, fn) {
    const db = await open();
    return new Promise((res, rej) => {
      const t = db.transaction(store, mode);
      const os = t.objectStore(store);
      let out;
      Promise.resolve(fn(os)).then(v => { out = v; });
      t.oncomplete = () => res(out);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error);
    });
  }
  const req2p = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  return {
    get: (s, k) => tx(s, 'readonly', os => req2p(os.get(k))),
    all: (s) => tx(s, 'readonly', os => req2p(os.getAll())),
    put: (s, v) => tx(s, 'readwrite', os => req2p(os.put(v))),
    del: (s, k) => tx(s, 'readwrite', os => req2p(os.delete(k))),
    clear: (s) => tx(s, 'readwrite', os => req2p(os.clear())),
    byIndex: (s, idx, v) => tx(s, 'readonly', os => req2p(os.index(idx).getAll(v))),
    stores: Object.keys(STORES),
    id: () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7)
  };
})();

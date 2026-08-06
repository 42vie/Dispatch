// ║  INDEXEDDB — Stockage local haute capacité (remplace JSON/localStorage) ║
// ╚══════════════════════════════════════════════════════════════════════════╝
const IDB_NAME = 'DispatchDB';
const IDB_VERSION = 2;
const IDB_STORE = 'state';
const IDB_CONTACTS = 'contacts';
let _idb = null;

function idbOpen() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
      if (!db.objectStoreNames.contains(IDB_CONTACTS)) db.createObjectStore(IDB_CONTACTS);
    };
    req.onsuccess = e => { _idb = e.target.result; resolve(_idb); };
    req.onerror = e => { console.warn('IndexedDB:', e.target.error); reject(e.target.error); };
  });
}

function idbPut(store, key, data) {
  if (!_idb) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const tx = _idb.transaction(store, 'readwrite');
    tx.objectStore(store).put(data, key);
    tx.oncomplete = () => resolve();
    tx.onerror = e => { console.warn('IDB save:', e.target.error); reject(e.target.error); };
  });
}

function idbGet(store, key) {
  if (!_idb) return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    const tx = _idb.transaction(store, 'readonly');
    const req = tx.objectStore(store).get(key);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = e => reject(e.target.error);
  });
}

// Raccourcis pour compatibilité
function idbSave(data) { return idbPut(IDB_STORE, 'dispatch_state', data); }
function idbLoad() { return idbGet(IDB_STORE, 'dispatch_state'); }
function idbSaveContacts() { return idbPut(IDB_CONTACTS, 'contacts', g_contacts); }
function idbLoadContacts() { return idbGet(IDB_CONTACTS, 'contacts'); }

// ╔══════════════════════════════════════════════════════════════════════════╗

// ║  AUTO-SAVE — dirty flag + debounce (ne sauvegarde que si changement)    ║
// ╚══════════════════════════════════════════════════════════════════════════╝
let _dirty = false;
let _saveTimer = null;
const SAVE_DELAY = 1500; // 1.5s après dernière modification

function markDirty() { _dirty = true; scheduleSave(); }

// Marquer un dossier comme modifié par cet opérateur (pour la fusion réseau)
function stampRecord(r) {
  r._ts = Date.now();
  const myName = localStorage.getItem('dispatch_operator') || '';
  r._by = myName;
  // Assigner l'opérateur si le dossier n'en a pas encore
  if (!r.operator && myName) r.operator = myName;
}

function scheduleSave() {
  if (_saveTimer) clearTimeout(_saveTimer);
  _saveTimer = setTimeout(autoSave, SAVE_DELAY);
}

async function autoSave() {
  if (!_dirty) return;
  _dirty = false;
  // Fichier ÉTAT : statut + opérateur de chaque dossier (léger)
  const statusData = g_master.map(r => ({ file: r.file, statut: r.statut, operator: r.operator || '', cc: r.cc || '', email: r.email || '', contact: r.contact || '', tel: r.tel || '', comment: r.comment || '', transp: r.transp || '', fcDate: r.fcDate || '', fcHoraire: r.fcHoraire || '' }));
  // Fichier DATA : toutes les infos dossiers + rawData + cpData (sans contacts)
  const infoData = { master: g_master, rawData: g_rawData, cpData: g_cpData };
  // Fichier CONTACTS : g_contacts (séparé, persistant)
  const contactsData = g_contacts;

  // Sauvegarder en parallèle : IndexedDB + API AutoIt (3 fichiers) + Réseau
  const path = localStorage.getItem('dispatch_state_path') || DEFAULT_STATE_PATH;
  try { await Promise.all([
    idbSave(infoData),
    idbSaveContacts(),
    apiCall('save-status', statusData),
    apiCall('save-data', infoData),
    saveContactsChunked(),
    path ? smartNetSave(path) : Promise.resolve()
  ]); }
  catch(e) { console.warn('autoSave error:', e); }
}

// Sauvegarde contacts avec auto-split si trop gros (>500 KB par fichier)
const CONTACTS_CHUNK_SIZE = 500 * 1024; // 500 KB
async function saveContactsChunked() {
  const json = JSON.stringify(g_contacts);
  if (json.length <= CONTACTS_CHUNK_SIZE) {
    // Un seul fichier suffit
    await apiCall('save-contacts', { chunk: 0, total: 1, data: g_contacts });
  } else {
    // Découper en morceaux
    const chunkCount = Math.ceil(json.length / CONTACTS_CHUNK_SIZE);
    const perChunk = Math.ceil(g_contacts.length / chunkCount);
    const promises = [];
    for (let i = 0; i < chunkCount; i++) {
      const slice = g_contacts.slice(i * perChunk, (i + 1) * perChunk);
      promises.push(apiCall('save-contacts', { chunk: i, total: chunkCount, data: slice }));
    }
    await Promise.all(promises);
  }
}

// Sauvegarde contacts uniquement (après ajout/modif/suppression)
async function saveContacts() {
  try {
    if (typeof saveContactsTSV === 'function') await saveContactsTSV();
    else if (typeof saveContactsInline === 'function') await saveContactsInline();
    else await Promise.all([idbSaveContacts(), saveContactsChunked()]);
  }
  catch(e) { console.warn('contacts save:', e); }
}

// Écoute intelligente : ne marque dirty que sur vrais changements
document.addEventListener('input', markDirty);
// Les fonctions appellent markDirty() après modification — le save est debounced automatiquement


// ╔══════════════════════════════════════════════════════════════════════════╗

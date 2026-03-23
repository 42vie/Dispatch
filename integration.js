// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  integration.js — Snippet d'intégration DataManager dans Interface.html  ║
// ║                                                                          ║
// ║  Ce fichier montre comment intégrer DataManager.js dans l'interface       ║
// ║  existante. À inclure dans Interface.html via <script> après les autres   ║
// ║  fichiers JS (validate.js, merge.js, migrate.js, DataManager.js).        ║
// ║                                                                          ║
// ║  SÉCURITÉ :                                                               ║
// ║  - Sanitisation des entrées utilisateur (XSS)                             ║
// ║  - Validation des données avant sauvegarde                                ║
// ║  - Protection beforeunload contre la perte de données                     ║
// ╚══════════════════════════════════════════════════════════════════════════╝

// ══════════════════════════════════════════════════════════════════════════
// 1. INSTANCIATION AU DÉMARRAGE
// ══════════════════════════════════════════════════════════════════════════

// Instance globale du DataManager
let dataManager = null;

/**
 * Initialise le DataManager au démarrage de l'application.
 * À appeler dans window.onload AVANT renderMaster()/renderKanban().
 */
async function initDataManager() {
  const operateur = localStorage.getItem('dispatch_operator') || '';
  dataManager = new DataManager(operateur);

  // Brancher les callbacks UI
  dataManager.onSaveStatusChange = updateSaveIndicator;
  dataManager.onSyncComplete = onSyncComplete;
  dataManager.onError = onDataManagerError;

  // Charger les données (non-bloquant pour l'UI)
  const loaded = await dataManager.load();

  if (loaded) {
    // Peupler g_master depuis DataManager pour compatibilité avec l'UI existante
    syncToLegacy();

    // Afficher l'UI
    if (typeof renderMaster === 'function') renderMaster();
    if (typeof renderKanban === 'function') renderKanban();
    if (typeof renderCP === 'function') renderCP();
    if (typeof renderContacts === 'function') renderContacts();
    if (typeof updateStats === 'function') updateStats();
    if (typeof populateOperatorFilter === 'function') populateOperatorFilter();
  }

  // Lancer le polling réseau intelligent
  startReseauPolling();
}


// ══════════════════════════════════════════════════════════════════════════
// 2. CHARGEMENT NON-BLOQUANT
// ══════════════════════════════════════════════════════════════════════════

/**
 * Synchronise les données de DataManager vers les variables legacy (g_master, etc.)
 * pour que l'UI existante continue de fonctionner.
 */
function syncToLegacy() {
  if (!dataManager) return;
  // Convertir les dossiers indexés en array pour g_master
  g_master = dataManager.getAllDossiers().map(d => ({
    file: d.file || d.id,
    client: d.client ? d.client.nom : '',
    rdl: d.transport ? d.transport.rdl : '',
    svct: d.transport ? d.transport.svct : '',
    transp: d.transport ? d.transport.transp : '',
    poids: d.transport ? d.transport.poids : 0,
    vol: d.transport ? d.transport.volume : 0,
    taxable: d.transport ? d.transport.taxable : 0,
    dept: d.transport ? d.transport.destination : '',
    contact: d.client ? d.client.contact : '',
    tel: d.client ? d.client.tel : '',
    email: d.client ? d.client.email : '',
    statut: d.transport ? String(d.transport.statut) : '0',
    operator: d.operator || '',
    cc: d.cc || '',
    comment: d.notes || '',
    _dateCreated: d.createdAt ? d.createdAt.split('T')[0] : '',
    _ts: d.updatedAt ? new Date(d.updatedAt).getTime() : 0,
    _by: d.updatedBy || '',
    fcDate: d.fc ? d.fc.date : '',
    fcHoraire: d.fc ? d.fc.horaire : '',
    fcDly: d.fc ? d.fc.dly : '',
    fcDlyNotes: d.fc ? d.fc.dlyNotes : '',
    // Référence vers l'ID DataManager pour les mises à jour
    _dmId: d.id
  }));
  g_rawData = dataManager.rawData || {};
  // cpData reste compatible
}


// ══════════════════════════════════════════════════════════════════════════
// 3. REMPLACEMENT smartNetSave → DataManager.syncReseau
// ══════════════════════════════════════════════════════════════════════════

/**
 * Remplace l'ancien smartNetSave().
 * Appeler cette fonction partout où smartNetSave() était utilisé.
 */
async function smartNetSaveV2(path) {
  if (!dataManager) return;
  // Syncer les données legacy vers DataManager avant la sauvegarde
  syncFromLegacy();
  await dataManager.syncReseau();
}

/**
 * Synchronise g_master (legacy) vers DataManager.
 * À appeler avant chaque sauvegarde réseau.
 */
function syncFromLegacy() {
  if (!dataManager || !Array.isArray(g_master)) return;
  g_master.forEach(r => {
    const id = r._dmId || r.file;
    const d = dataManager.dossiers[id];
    if (d) {
      // Mettre à jour depuis les données legacy
      if (d.transport) d.transport.statut = parseInt(r.statut) || 0;
      d.operator = r.operator || '';
      d.cc = r.cc || '';
      d.notes = r.comment || '';
      d.updatedAt = r._ts ? new Date(r._ts).toISOString() : d.updatedAt;
      d.updatedBy = r._by || d.updatedBy;
    }
  });
}


// ══════════════════════════════════════════════════════════════════════════
// 4. REMPLACEMENT AUTO-SAVE → DataManager.markDirty
// ══════════════════════════════════════════════════════════════════════════

/**
 * Remplace l'ancien markDirty() global.
 * Quand un dossier est modifié, appeler markDirtyV2(id).
 */
function markDirtyV2(fileOrId) {
  if (!dataManager) return;
  // Syncer les données legacy d'abord
  syncFromLegacy();
  dataManager.markDirty(fileOrId);
}

/**
 * Remplacement de l'ancien autoSave.
 * À brancher sur le debounce existant.
 */
async function autoSaveV2() {
  if (!dataManager) return;
  syncFromLegacy();
  await dataManager.saveAll();
}


// ══════════════════════════════════════════════════════════════════════════
// 5. SAVE BEFORE UNLOAD — Protection contre la perte de données
// ══════════════════════════════════════════════════════════════════════════

window.addEventListener('beforeunload', (e) => {
  if (dataManager) {
    syncFromLegacy();
    dataManager.saveBeforeUnload();
  }
});

// Backup supplémentaire : visibilitychange (mobile / changement d'onglet)
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden' && dataManager) {
    syncFromLegacy();
    dataManager.saveBeforeUnload();
  }
});


// ══════════════════════════════════════════════════════════════════════════
// 6. INDICATEUR UI — Statut de sauvegarde
// ══════════════════════════════════════════════════════════════════════════

/**
 * Met à jour l'indicateur de sauvegarde dans l'UI.
 * Crée l'élément s'il n'existe pas.
 */
function updateSaveIndicator(status, message) {
  let el = document.getElementById('save-status-indicator');
  if (!el) {
    // Créer l'indicateur
    el = document.createElement('div');
    el.id = 'save-status-indicator';
    el.style.cssText = 'position:fixed;bottom:8px;right:8px;padding:6px 12px;' +
      'border-radius:6px;font-size:11px;font-family:var(--mono,monospace);' +
      'z-index:9999;transition:all 0.3s ease;pointer-events:none;';
    document.body.appendChild(el);
  }

  // Sanitiser le message pour éviter les injections XSS
  const safeMessage = sanitizeForDisplay(message || '');

  switch (status) {
    case 'ok':
      el.style.background = 'rgba(34,197,94,0.15)';
      el.style.color = '#22c55e';
      el.style.border = '1px solid rgba(34,197,94,0.3)';
      el.textContent = '\u2705 ' + safeMessage;
      // Masquer après 5s
      clearTimeout(el._hideTimer);
      el._hideTimer = setTimeout(() => { el.style.opacity = '0'; }, 5000);
      el.style.opacity = '1';
      break;

    case 'pending':
      el.style.background = 'rgba(234,179,8,0.15)';
      el.style.color = '#eab308';
      el.style.border = '1px solid rgba(234,179,8,0.3)';
      el.textContent = '\u23F3 ' + safeMessage;
      el.style.opacity = '1';
      break;

    case 'error':
      el.style.background = 'rgba(239,68,68,0.15)';
      el.style.color = '#ef4444';
      el.style.border = '1px solid rgba(239,68,68,0.3)';
      el.textContent = '\u26A0\uFE0F ' + safeMessage;
      el.style.opacity = '1';
      // Ne pas masquer les erreurs automatiquement
      break;
  }
}

function onSyncComplete(stats) {
  if (stats.synced > 0) {
    syncToLegacy();
    if (typeof renderMaster === 'function') renderMaster();
    if (typeof renderKanban === 'function') renderKanban();
    if (typeof updateStats === 'function') updateStats();
    if (typeof toast === 'function') toast('Sync réseau : ' + stats.synced + ' dossier(s) synchronisés.');
  }
}

function onDataManagerError(context, error) {
  console.error('[DM.' + context + ']', error);
  updateSaveIndicator('error', 'Erreur ' + context + ' : ' + (error.message || error));
}


// ══════════════════════════════════════════════════════════════════════════
// 7. POLLING RÉSEAU INTELLIGENT AVEC BACKOFF
// ══════════════════════════════════════════════════════════════════════════

let _pollInterval = 10000; // 10s de base
let _pollTimer = null;
const POLL_MIN = 10000;    // 10s minimum
const POLL_MAX = 60000;    // 60s maximum
const POLL_BACKOFF = 1.5;  // Multiplicateur de backoff

async function pollReseau() {
  if (!dataManager) return;

  const path = localStorage.getItem('dispatch_state_path') || '';
  if (!path) {
    _pollInterval = POLL_MAX;
    scheduleNextPoll();
    return;
  }

  try {
    const resp = await fetch(API_BASE + '/net-check?path=' + encodeURIComponent(path));
    const result = await resp.json();

    if (result.changed) {
      await dataManager.syncReseau();
      syncToLegacy();
      if (typeof renderMaster === 'function') renderMaster();
      if (typeof renderKanban === 'function') renderKanban();
      if (typeof updateStats === 'function') updateStats();
      _pollInterval = POLL_MIN; // Reset après changement
    } else {
      // Rien n'a changé — augmenter l'intervalle
      _pollInterval = Math.min(_pollInterval * POLL_BACKOFF, POLL_MAX);
    }
  } catch (e) {
    // Erreur réseau — backoff agressif
    _pollInterval = Math.min(_pollInterval * 2, POLL_MAX);
    console.warn('Poll réseau erreur:', e);
  }

  scheduleNextPoll();
}

function scheduleNextPoll() {
  if (_pollTimer) clearTimeout(_pollTimer);
  _pollTimer = setTimeout(pollReseau, _pollInterval);
}

function startReseauPolling() {
  // Premier poll après 3s (laisser le temps au chargement initial)
  _pollTimer = setTimeout(pollReseau, 3000);
}

function stopReseauPolling() {
  if (_pollTimer) clearTimeout(_pollTimer);
  _pollTimer = null;
}


// ══════════════════════════════════════════════════════════════════════════
// 8. BATCH E.TMS NON-BLOQUANT AVEC BARRE DE PROGRESSION
// ══════════════════════════════════════════════════════════════════════════

/**
 * Lance un batch E.TMS et affiche une barre de progression non-bloquante.
 *
 * @param {Array} dossiers - Array d'IDs ou de numéros de dossiers à traiter
 * @param {string} type    - "ETMS" ou "COMAT"
 */
async function lancerBatchETMS(dossiers, type) {
  if (!dossiers || !dossiers.length) return;

  type = type || 'ETMS';

  // Envoyer la demande au backend AutoIt
  try {
    const resp = await fetch(API_BASE + '/etms-batch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dossiers, type })
    });
    const result = await resp.json();
    if (!result || result.error) {
      if (typeof toast === 'function') toast('Erreur lancement batch : ' + (result.error || 'inconnu'));
      return;
    }
  } catch (e) {
    if (typeof toast === 'function') toast('Impossible de lancer le batch : ' + e.message);
    return;
  }

  // Afficher la barre de progression
  showProgressBar(type, dossiers.length);

  // Poll le statut sans bloquer l'UI
  let done = false;
  while (!done) {
    try {
      const resp = await fetch(API_BASE + '/job-status');
      const status = await resp.json();

      updateProgressBar(status.progress || 0, status.total || dossiers.length, status.current || '', status.status);

      if (status.done || status.status === 'idle') {
        done = true;
        hideProgressBar(true); // true = succès
        if (typeof toast === 'function') {
          toast(type + ' terminé : ' + (status.progress || 0) + '/' + (status.total || dossiers.length) + ' dossier(s) traités.');
        }
      }
    } catch (e) {
      // Erreur de poll — on continue d'attendre
      console.warn('job-status poll error:', e);
    }

    if (!done) await new Promise(r => setTimeout(r, 500));
  }
}

function showProgressBar(type, total) {
  // Supprimer l'existante si présente
  hideProgressBar();

  const bar = document.createElement('div');
  bar.id = 'batch-progress-bar';
  bar.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:10000;' +
    'background:var(--surface2,#1e1e2e);border-bottom:1px solid var(--border,#333);' +
    'padding:8px 16px;display:flex;align-items:center;gap:12px;font-size:12px;' +
    'font-family:var(--mono,monospace);color:var(--text1,#e0e0e0);';

  bar.innerHTML =
    '<span id="bp-label" style="min-width:100px">' + sanitizeForDisplay(type) + '</span>' +
    '<div style="flex:1;height:6px;background:var(--surface3,#2a2a3e);border-radius:3px;overflow:hidden">' +
    '  <div id="bp-fill" style="height:100%;width:0%;background:var(--accent,#7c3aed);' +
    '       border-radius:3px;transition:width 0.3s ease"></div>' +
    '</div>' +
    '<span id="bp-count" style="min-width:60px;text-align:right">0/' + total + '</span>' +
    '<span id="bp-current" style="color:var(--text3,#888);max-width:150px;overflow:hidden;' +
    '       text-overflow:ellipsis;white-space:nowrap"></span>' +
    '<button onclick="stopBatchETMS()" style="padding:2px 8px;border:1px solid var(--border,#333);' +
    '        border-radius:4px;background:transparent;color:var(--text2,#ccc);cursor:pointer;' +
    '        font-size:11px" title="Arrêter (Échap)">Stop</button>';

  document.body.appendChild(bar);
}

function updateProgressBar(progress, total, current, status) {
  const fill = document.getElementById('bp-fill');
  const count = document.getElementById('bp-count');
  const label = document.getElementById('bp-label');
  const cur = document.getElementById('bp-current');

  if (fill && total > 0) fill.style.width = Math.round((progress / total) * 100) + '%';
  if (count) count.textContent = progress + '/' + total;
  if (cur) cur.textContent = sanitizeForDisplay(current);
  if (label && status === 'paused') label.textContent += ' [PAUSE]';
}

function hideProgressBar(success) {
  const bar = document.getElementById('batch-progress-bar');
  if (!bar) return;

  if (success) {
    const fill = document.getElementById('bp-fill');
    if (fill) {
      fill.style.width = '100%';
      fill.style.background = '#22c55e';
    }
    setTimeout(() => { if (bar.parentNode) bar.parentNode.removeChild(bar); }, 1500);
  } else {
    if (bar.parentNode) bar.parentNode.removeChild(bar);
  }
}

function stopBatchETMS() {
  // Envoyer un signal d'arrêt via l'API
  fetch(API_BASE + '/etms-stop', { method: 'POST' }).catch(() => {});
}


// ══════════════════════════════════════════════════════════════════════════
// SÉCURITÉ — Sanitisation pour l'affichage (anti-XSS)
// ══════════════════════════════════════════════════════════════════════════

/**
 * Sanitise une chaîne pour l'affichage dans le DOM.
 * Échappe les caractères HTML dangereux.
 */
function sanitizeForDisplay(str) {
  if (typeof str !== 'string') return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

/**
 * Sanitise un objet de données utilisateur avant sauvegarde.
 * Supprime les propriétés dangereuses et nettoie les strings.
 */
function sanitizeData(obj) {
  if (typeof obj === 'string') {
    // Supprimer les caractères de contrôle (sauf \t \n \r)
    return obj.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }
  if (Array.isArray(obj)) return obj.map(sanitizeData);
  if (obj && typeof obj === 'object') {
    const clean = {};
    for (const [k, v] of Object.entries(obj)) {
      // Bloquer les clés suspectes
      if (k.startsWith('__') || k === 'constructor' || k === 'prototype') continue;
      clean[k] = sanitizeData(v);
    }
    return clean;
  }
  return obj;
}


// ══════════════════════════════════════════════════════════════════════════
// INTÉGRATION — Ordre de chargement dans Interface.html
// ══════════════════════════════════════════════════════════════════════════
//
// Ajouter ces <script> à la fin du <body> de Interface.html :
//
//   <script src="validate.js"></script>
//   <script src="merge.js"></script>
//   <script src="migrate.js"></script>
//   <script src="DataManager.js"></script>
//   <script src="integration.js"></script>
//
// Puis remplacer le window.onload existant par :
//
//   window.onload = async () => {
//     await initDataManager();
//     optsLoadPJ();
//     optsLoadReseau();
//     optsLoadFuel();
//     cpCfgLoad();
//     showIdentityModal();
//   };
//
// Et remplacer les appels existants :
//   - smartNetSave(path)  →  smartNetSaveV2(path)
//   - markDirty()         →  markDirtyV2(id)
//   - autoSave()          →  autoSaveV2()
//
// ══════════════════════════════════════════════════════════════════════════

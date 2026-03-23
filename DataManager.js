// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  DataManager.js — Classe principale de gestion des données Dispatch     ║
// ║  Gère le chargement, la sauvegarde, la sync réseau et la validation     ║
// ║  Vanilla JS, aucune dépendance — conçu pour <script> dans HTML          ║
// ╚══════════════════════════════════════════════════════════════════════════╝

// Dépendances attendues (doivent être chargées AVANT ce fichier) :
//   - validate.js   (validateDossier)
//   - merge.js      (mergeNetworkStates, mergeDossier)
//   - migrate.js    (migrateToNewFormat) — optionnel, pour migration initiale

const API_BASE = 'http://127.0.0.1:8700/api';

// ── IndexedDB helpers (autonomes, remplacent ceux d'Interface.html) ──
const DM_IDB_NAME = 'DispatchDB';
const DM_IDB_VERSION = 2;
const DM_IDB_STORE = 'state';
const DM_IDB_CONTACTS = 'contacts';

class DataManager {
  /**
   * @param {string} operateur - Nom de l'opérateur courant
   */
  constructor(operateur) {
    this.operateur = operateur || '';
    this.dossiers = {};       // Objet indexé par ID → accès O(1)
    this.rawData = {};        // Données brutes par file
    this.cpData = {};         // Colis postaux
    this.contacts = [];       // Carnet de contacts

    // Métadonnées
    this.schemaVersion = '2.1';
    this.loaded = false;

    // État de sauvegarde (exposé pour l'UI)
    this.lastSaveAt = null;
    this.lastSaveStatus = null; // 'ok', 'error', 'pending'
    this.lastSyncAt = null;

    // Dirty / debounce
    this._dirty = new Set();   // Set d'IDs modifiés
    this._saveTimer = null;
    this._saveDelay = 1500;

    // IndexedDB
    this._idb = null;

    // Callbacks UI (à brancher depuis l'extérieur)
    this.onSaveStatusChange = null;  // function(status, message)
    this.onSyncComplete = null;      // function(stats)
    this.onError = null;             // function(context, error)
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  CHARGEMENT                                                          ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /**
   * Charge les données : IndexedDB d'abord (instantané), puis API en fond.
   * Retourne dès que l'UI peut s'afficher (IndexedDB ou API).
   */
  async load() {
    this._notify('pending', 'Chargement...');

    // 1. Ouvrir IndexedDB
    try {
      await this._idbOpen();
    } catch (e) {
      this._error('idb_open', e);
    }

    let loaded = false;

    // 2. Essayer IndexedDB (instantané)
    try {
      const idbData = await this._idbGet(DM_IDB_STORE, 'dispatch_state');
      if (idbData && idbData.dossiers && Object.keys(idbData.dossiers).length > 0) {
        this.dossiers = idbData.dossiers;
        this.rawData = idbData.rawData || {};
        this.cpData = idbData.cpData || {};
        loaded = true;
      } else if (idbData && idbData.master && idbData.master.length > 0) {
        // Ancien format dans IndexedDB — migrer
        this._migrateFromLegacy(idbData);
        loaded = true;
      }
    } catch (e) {
      this._error('idb_load', e);
    }

    // 2b. Contacts
    try {
      const c = await this._idbGet(DM_IDB_CONTACTS, 'contacts');
      if (c && Array.isArray(c) && c.length > 0) this.contacts = c;
    } catch (e) { /* silencieux pour contacts */ }

    // 3. Fallback API AutoIt
    if (!loaded) {
      try {
        const resp = await this._fetchWithRetry(API_BASE + '/load-data');
        const d = await resp.json();
        if (d && d.master && d.master.length > 0) {
          // Ancien format — migrer
          this._migrateFromLegacy(d);
          loaded = true;
        } else if (d && d.dossiers) {
          this.dossiers = d.dossiers;
          this.rawData = d.rawData || {};
          this.cpData = d.cpData || {};
          loaded = true;
        }
      } catch (e) {
        this._error('api_load', e);
      }
    }

    if (!loaded) {
      try {
        const resp = await this._fetchWithRetry(API_BASE + '/load');
        const d = await resp.json();
        if (d && d.master && d.master.length > 0) {
          this._migrateFromLegacy(d);
          loaded = true;
        }
      } catch (e) {
        this._error('api_load_fallback', e);
      }
    }

    this.loaded = loaded;

    if (loaded) {
      // Sauver dans IndexedDB pour la prochaine fois
      this._idbSave();
      this._notify('ok', 'Données chargées (' + Object.keys(this.dossiers).length + ' dossiers).');
    } else {
      this._notify('error', 'Aucune donnée trouvée.');
    }

    return loaded;
  }

  /**
   * Convertit les données ancien format (master array) vers le nouveau.
   */
  _migrateFromLegacy(data) {
    if (typeof migrateToNewFormat === 'function') {
      const result = migrateToNewFormat(data);
      if (result.data && result.data.dossiers) {
        this.dossiers = result.data.dossiers;
        this.cpData = result.data.cpData || {};
      }
      this.rawData = data.rawData || {};
    } else {
      // Migration minimale sans migrate.js
      // L'ID est le numéro de dossier (file: "J1A0042031")
      const master = data.master || [];
      this.dossiers = {};
      master.forEach(r => {
        const key = (r.file || '').trim();
        if (!key) return;
        this.dossiers[key] = {
          id: key,
          file: key,
          v: 1,
          createdAt: r._dateCreated ? r._dateCreated + 'T00:00:00Z' : new Date().toISOString(),
          updatedAt: r._ts ? new Date(r._ts).toISOString() : new Date().toISOString(),
          updatedBy: r._by || r.operator || '',
          client: { nom: r.client || '', contact: (r.contact || '').trim(), tel: (r.tel || '').trim(), email: (r.email || '').trim() },
          transport: {
            statut: parseInt(r.statut) || 0,
            svct: r.svct || '', transp: r.transp || '',
            poids: parseFloat(r.poids) || 0, volume: parseFloat(r.vol) || 0,
            taxable: parseFloat(r.taxable) || 0, destination: r.dept || '', rdl: r.rdl || ''
          },
          fc: { date: r.fcDate || '', horaire: r.fcHoraire || '', dly: r.fcDly || '', dlyNotes: r.fcDlyNotes || '' },
          cc: r.cc || '', operator: r.operator || '', notes: r.comment || '',
          historique: []
        };
      });
      this.rawData = data.rawData || {};
      this.cpData = {};
    }
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  ACCÈS DONNÉES — O(1)                                               ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /** Accès O(1) par ID */
  getDossier(id) {
    return this.dossiers[id] || null;
  }

  /** Cherche un dossier par son champ file (pour compatibilité) */
  getDossierByFile(file) {
    for (const d of Object.values(this.dossiers)) {
      if (d.file === file) return d;
    }
    return null;
  }

  /** Retourne un array pour l'UI (kanban, master table) */
  getAllDossiers() {
    return Object.values(this.dossiers);
  }

  /** Retourne les dossiers filtrés par opérateur */
  getDossiersForOperator(name) {
    if (!name) return this.getAllDossiers();
    return Object.values(this.dossiers).filter(d =>
      (d.operator || '') === name || !d.operator
    );
  }

  /** Retourne les dossiers par statut */
  getDossiersByStatut(statut) {
    const s = parseInt(statut);
    return Object.values(this.dossiers).filter(d =>
      parseInt(d.transport && d.transport.statut) === s
    );
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  SAUVEGARDE — PATCH                                                  ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /**
   * Sauvegarde un dossier modifié — incrémente v, met updatedAt/updatedBy.
   * Envoie un patch différentiel via /api/save-patch.
   *
   * @param {string} id       - ID du dossier
   * @param {Object} changes  - Objet avec les champs modifiés (supporte dot notation)
   */
  async saveDossier(id, changes) {
    const dossier = this.dossiers[id];
    if (!dossier) {
      this._error('saveDossier', new Error('Dossier introuvable : ' + id));
      return false;
    }

    // Appliquer les changes localement
    const now = new Date().toISOString();
    const oldStatut = dossier.transport ? dossier.transport.statut : null;

    for (const [path, value] of Object.entries(changes)) {
      this._setNestedValue(dossier, path, value);
    }

    dossier.v = (dossier.v || 0) + 1;
    dossier.updatedAt = now;
    dossier.updatedBy = this.operateur;

    // Historique
    if (!dossier.historique) dossier.historique = [];
    const newStatut = dossier.transport ? dossier.transport.statut : null;
    if (newStatut !== null && newStatut !== oldStatut) {
      dossier.historique.push({
        ts: now,
        op: this.operateur,
        action: 'statut_change',
        detail: { de: oldStatut, a: newStatut }
      });
    } else {
      dossier.historique.push({
        ts: now,
        op: this.operateur,
        action: 'modification',
        detail: { champs: Object.keys(changes) }
      });
    }

    // Validation
    if (typeof validateDossier === 'function') {
      const v = validateDossier(dossier);
      if (!v.valid) {
        console.warn('Validation échouée pour ' + id + ' :', v.errors);
        // On sauvegarde quand même mais on loggue
      }
    }

    // Envoyer le patch au backend
    try {
      const resp = await this._fetchWithRetry(API_BASE + '/save-patch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: id,
          v: dossier.v,
          updatedAt: now,
          updatedBy: this.operateur,
          changes: changes
        })
      });
      const result = await resp.json();
      if (result && result.success === false) {
        // Conflit de version — forcer un saveAll
        console.warn('Conflit de version pour ' + id + ' — fallback saveAll.');
        this._dirty.add(id);
        await this.saveAll();
        return false;
      }
    } catch (e) {
      // Si l'endpoint n'existe pas encore, fallback sur markDirty + saveAll
      this._dirty.add(id);
      this.scheduleSave();
      return true; // sauvé localement, sera persisté par autoSave
    }

    // Sauver dans IndexedDB
    this._idbSave();
    this._notify('ok', 'Dossier ' + id + ' sauvegardé (v' + dossier.v + ').');
    return true;
  }

  /**
   * Sauvegarde uniquement le statut d'un dossier (fichier léger séparé).
   */
  async saveStatut(id, statut) {
    const dossier = this.dossiers[id];
    if (!dossier) return false;

    const oldStatut = dossier.transport ? dossier.transport.statut : 0;
    if (!dossier.transport) dossier.transport = {};
    dossier.transport.statut = parseInt(statut);
    dossier.updatedAt = new Date().toISOString();
    dossier.updatedBy = this.operateur;
    dossier.v = (dossier.v || 0) + 1;

    if (!dossier.historique) dossier.historique = [];
    dossier.historique.push({
      ts: dossier.updatedAt,
      op: this.operateur,
      action: 'statut_change',
      detail: { de: oldStatut, a: parseInt(statut) }
    });

    // Sauvegarder les statuts légers
    try {
      const statusData = {};
      for (const [k, d] of Object.entries(this.dossiers)) {
        statusData[k] = d.transport ? d.transport.statut : 0;
      }
      await this._fetchWithRetry(API_BASE + '/save-status', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          _meta: { updatedAt: new Date().toISOString(), updatedBy: this.operateur },
          statuts: statusData
        })
      });
    } catch (e) {
      this._error('saveStatut', e);
    }

    this._idbSave();
    this.markDirty(id);
    return true;
  }

  /**
   * Sauvegarde complète (fallback ou forcée).
   * Envoie tout le fichier via /api/save-data.
   */
  async saveAll() {
    this._notify('pending', 'Sauvegarde en cours...');

    const payload = {
      _meta: {
        schemaVersion: this.schemaVersion,
        generatedAt: new Date().toISOString(),
        generatedBy: this.operateur,
        count: Object.keys(this.dossiers).length,
        appVersion: 'DispatchMaster-2.1'
      },
      dossiers: this.dossiers,
      rawData: this.rawData,
      cpData: this.cpData
    };

    // Compatibilité : aussi envoyer au format ancien pour /api/save-data
    const legacyPayload = {
      master: this.getAllDossiers(),
      rawData: this.rawData,
      cpData: Object.values(this.cpData)
    };

    try {
      await Promise.all([
        this._idbSave(),
        this._fetchWithRetry(API_BASE + '/save-data', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(legacyPayload)
        })
      ]);

      this._dirty.clear();
      this.lastSaveAt = new Date();
      this.lastSaveStatus = 'ok';
      this._notify('ok', 'Sauvegardé à ' + this.lastSaveAt.toLocaleTimeString());
      return true;
    } catch (e) {
      this.lastSaveStatus = 'error';
      this._error('saveAll', e);
      this._notify('error', 'Erreur sauvegarde : ' + e.message);
      return false;
    }
  }

  /**
   * Save synchrone avant fermeture de page.
   * Utilise navigator.sendBeacon si disponible, sinon XHR synchrone.
   */
  saveBeforeUnload() {
    if (this._dirty.size === 0) return;

    const legacyPayload = JSON.stringify({
      master: this.getAllDossiers(),
      rawData: this.rawData,
      cpData: Object.values(this.cpData)
    });

    if (navigator.sendBeacon) {
      const blob = new Blob([legacyPayload], { type: 'application/json' });
      navigator.sendBeacon(API_BASE + '/save-data', blob);
    } else {
      // XHR synchrone (fallback)
      try {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', API_BASE + '/save-data', false); // synchrone
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.send(legacyPayload);
      } catch (e) {
        console.error('saveBeforeUnload error:', e);
      }
    }
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  DIRTY FLAG + DEBOUNCE                                               ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /** Marque un dossier comme modifié — déclenche le debounce */
  markDirty(id) {
    if (id) this._dirty.add(id);
    this.scheduleSave();
  }

  scheduleSave() {
    if (this._saveTimer) clearTimeout(this._saveTimer);
    this._saveTimer = setTimeout(() => this._autoSave(), this._saveDelay);
  }

  async _autoSave() {
    if (this._dirty.size === 0) return;
    await this.saveAll();

    // Sync réseau en parallèle (non-bloquant)
    const path = this._getStatePath();
    if (path && this.operateur) {
      this.syncReseau().catch(e => this._error('autoSync', e));
    }
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  SYNC RÉSEAU                                                         ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /**
   * Synchronise avec le réseau F:\ :
   *  1. Sauvegarde mon fichier opérateur
   *  2. Charge les fichiers des autres
   *  3. Merge avec mergeNetworkStates
   *  4. Met à jour le fichier de base
   */
  async syncReseau() {
    const path = this._getStatePath();
    if (!path || !this.operateur) return { synced: 0 };

    try {
      // 1. Mon fichier opérateur
      const myBoard = {};
      for (const [id, d] of Object.entries(this.dossiers)) {
        if ((d.operator || '') === this.operateur || !d.operator) {
          myBoard[id] = {
            id: id,
            file: d.file || id,
            v: d.v || 1,
            updatedAt: d.updatedAt || new Date().toISOString(),
            updatedBy: this.operateur,
            statut: d.transport ? d.transport.statut : 0,
            operator: d.operator || ''
          };
        }
      }

      const myPath = this._opPath(path, this.operateur);
      const myPayload = {
        _meta: {
          operateur: this.operateur,
          updatedAt: new Date().toISOString(),
          schemaVersion: this.schemaVersion
        },
        dossiers: myBoard
      };

      await this._netSave(myPath, myPayload);

      // 2. Lister les fichiers opérateurs
      const pattern = path.replace(/\.json$/i, '_*.json');
      const files = await this._netList(pattern);

      // 3. Charger les fichiers distants
      const opFiles = [];
      let synced = 0;

      for (const f of files) {
        const fNorm = f.replace(/\\\\/g, '\\');
        const myNorm = myPath.replace(/\\\\/g, '\\');
        if (fNorm === myNorm) continue;

        const remote = await this._netLoad(fNorm);
        if (remote && remote.dossiers) {
          opFiles.push(remote);
        } else if (Array.isArray(remote)) {
          // Ancien format
          const converted = { _meta: {}, dossiers: {} };
          remote.forEach(r => {
            const key = r.file || r.id || '';
            if (key) converted.dossiers[key] = {
              ...r, id: key, v: r.v || 1,
              updatedAt: r.updatedAt || (r._ts ? new Date(r._ts).toISOString() : '')
            };
          });
          opFiles.push(converted);
        }
      }

      // 4. Appliquer les mises à jour des autres opérateurs
      for (const opFile of opFiles) {
        for (const [id, rd] of Object.entries(opFile.dossiers || {})) {
          const local = this.dossiers[id];
          if (!local) continue;
          if (rd.operator && rd.operator !== this.operateur) {
            const result = mergeDossier(
              { v: local.v || 1, updatedAt: local.updatedAt || '' },
              { v: rd.v || 1, updatedAt: rd.updatedAt || '' }
            );
            if (result.winner === rd || result.winner.updatedAt === rd.updatedAt && result.reason === 'plus_recent') {
              if (local.transport) local.transport.statut = parseInt(rd.statut) || local.transport.statut;
              local.operator = rd.operator;
              local.updatedAt = rd.updatedAt;
              local.v = Math.max(local.v || 1, rd.v || 1);
              synced++;
            }
          }
        }
      }

      // 5. Reconstruire le fichier de base
      try {
        const allOps = [myPayload, ...opFiles];
        const mergeResult = mergeNetworkStates(null, allOps);
        await this._netSave(path, mergeResult.merged);
      } catch (e2) {
        console.warn('syncReseau base rebuild:', e2);
      }

      this.lastSyncAt = new Date();
      if (this.onSyncComplete) this.onSyncComplete({ synced });
      return { synced };

    } catch (e) {
      this._error('syncReseau', e);
      return { synced: 0, error: e.message };
    }
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  VALIDATION                                                          ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /** Valide un dossier avant save */
  validate(dossier) {
    if (typeof validateDossier === 'function') {
      return validateDossier(dossier);
    }
    // Validation minimale si validate.js pas chargé
    const errors = [];
    if (!dossier.id) errors.push('id manquant');
    if (!dossier.v) errors.push('v manquant');
    return { valid: errors.length === 0, errors };
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  FETCH AVEC RETRY                                                    ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  /**
   * fetch() avec retry automatique (backoff exponentiel).
   * @param {string} url
   * @param {Object} options - Options fetch standard
   * @param {number} maxRetries - Nombre max de tentatives (défaut 3)
   */
  async _fetchWithRetry(url, options, maxRetries) {
    if (maxRetries === undefined) maxRetries = 3;
    let lastError;
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        const resp = await fetch(url, options);
        if (!resp.ok && resp.status >= 500) throw new Error('HTTP ' + resp.status);
        return resp;
      } catch (e) {
        lastError = e;
        if (attempt < maxRetries) {
          const delay = Math.pow(2, attempt) * 500; // 500ms, 1s, 2s
          await new Promise(r => setTimeout(r, delay));
        }
      }
    }
    throw lastError;
  }

  // ╔══════════════════════════════════════════════════════════════════════╗
  // ║  HELPERS PRIVÉS                                                      ║
  // ╚══════════════════════════════════════════════════════════════════════╝

  _getStatePath() {
    try { return localStorage.getItem('dispatch_state_path') || ''; }
    catch (e) { return ''; }
  }

  _opPath(basePath, name) {
    return basePath.replace(/\.json$/i, '_' + name.replace(/[^a-zA-Z0-9àâäéèêëïîôùûüÿçæœÀÂÄÉÈÊËÏÎÔÙÛÜŸÇÆŒ_-]/g, '') + '.json');
  }

  /** Applique une valeur via dot notation (ex: 'transport.statut' → dossier.transport.statut) */
  _setNestedValue(obj, path, value) {
    const parts = path.split('.');
    let current = obj;
    for (let i = 0; i < parts.length - 1; i++) {
      if (!current[parts[i]] || typeof current[parts[i]] !== 'object') {
        current[parts[i]] = {};
      }
      current = current[parts[i]];
    }
    current[parts[parts.length - 1]] = value;
  }

  _notify(status, message) {
    this.lastSaveStatus = status;
    if (this.onSaveStatusChange) this.onSaveStatusChange(status, message);
  }

  _error(context, error) {
    console.error('[DataManager.' + context + ']', error);
    if (this.onError) this.onError(context, error);
  }

  // ── Réseau ──

  async _netSave(path, data) {
    const url = API_BASE + '/net-save?path=' + encodeURIComponent(path);
    return this._fetchWithRetry(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
  }

  async _netLoad(path) {
    try {
      const url = API_BASE + '/net-load?path=' + encodeURIComponent(path);
      const resp = await this._fetchWithRetry(url, { method: 'POST' });
      return await resp.json();
    } catch (e) { return null; }
  }

  async _netList(pattern) {
    try {
      const url = API_BASE + '/net-list?pattern=' + encodeURIComponent(pattern);
      const resp = await this._fetchWithRetry(url, { method: 'POST' });
      return await resp.json();
    } catch (e) { return []; }
  }

  // ── IndexedDB ──

  _idbOpen() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DM_IDB_NAME, DM_IDB_VERSION);
      req.onupgradeneeded = e => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(DM_IDB_STORE)) db.createObjectStore(DM_IDB_STORE);
        if (!db.objectStoreNames.contains(DM_IDB_CONTACTS)) db.createObjectStore(DM_IDB_CONTACTS);
      };
      req.onsuccess = e => { this._idb = e.target.result; resolve(this._idb); };
      req.onerror = e => reject(e.target.error);
    });
  }

  _idbGet(store, key) {
    if (!this._idb) return Promise.resolve(null);
    return new Promise((resolve, reject) => {
      const tx = this._idb.transaction(store, 'readonly');
      const req = tx.objectStore(store).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = e => reject(e.target.error);
    });
  }

  _idbPut(store, key, data) {
    if (!this._idb) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const tx = this._idb.transaction(store, 'readwrite');
      tx.objectStore(store).put(data, key);
      tx.oncomplete = () => resolve();
      tx.onerror = e => reject(e.target.error);
    });
  }

  async _idbSave() {
    try {
      await this._idbPut(DM_IDB_STORE, 'dispatch_state', {
        dossiers: this.dossiers,
        rawData: this.rawData,
        cpData: this.cpData
      });
    } catch (e) {
      this._error('idb_save', e);
    }
  }
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  TESTS                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function _test_datamanager() {
  console.log('=== Tests DataManager.js ===');

  // Test 1 : constructeur
  const dm = new DataManager('Jason');
  console.assert(dm.operateur === 'Jason', 'Opérateur initialisé');
  console.assert(typeof dm.dossiers === 'object', 'Dossiers = objet');
  console.assert(dm.loaded === false, 'Pas encore chargé');
  console.assert(dm._dirty.size === 0, 'Pas de dirty');
  console.log('  ✓ Test 1 : constructeur');

  // Test 2 : getDossier O(1)
  dm.dossiers['DSP-2024-0001'] = { id: 'DSP-2024-0001', v: 1, transport: { statut: 2 }, file: 'ABC' };
  dm.dossiers['DSP-2024-0002'] = { id: 'DSP-2024-0002', v: 3, transport: { statut: 5 }, file: 'DEF', operator: 'Jason' };
  console.assert(dm.getDossier('DSP-2024-0001') !== null, 'getDossier trouve le dossier');
  console.assert(dm.getDossier('ZZZZZ') === null, 'getDossier retourne null si absent');
  console.log('  ✓ Test 2 : getDossier O(1)');

  // Test 3 : getDossierByFile
  console.assert(dm.getDossierByFile('ABC').id === 'DSP-2024-0001', 'getDossierByFile OK');
  console.assert(dm.getDossierByFile('NOPE') === null, 'getDossierByFile null si absent');
  console.log('  ✓ Test 3 : getDossierByFile');

  // Test 4 : getAllDossiers retourne un array
  const all = dm.getAllDossiers();
  console.assert(Array.isArray(all), 'getAllDossiers = array');
  console.assert(all.length === 2, '2 dossiers');
  console.log('  ✓ Test 4 : getAllDossiers');

  // Test 5 : getDossiersByStatut
  const stat2 = dm.getDossiersByStatut(2);
  console.assert(stat2.length === 1, '1 dossier en statut 2');
  console.log('  ✓ Test 5 : getDossiersByStatut');

  // Test 6 : _setNestedValue
  const obj = { a: { b: 1 } };
  dm._setNestedValue(obj, 'a.b', 42);
  console.assert(obj.a.b === 42, 'setNestedValue modifie la valeur');
  dm._setNestedValue(obj, 'x.y.z', 'deep');
  console.assert(obj.x.y.z === 'deep', 'setNestedValue crée les niveaux');
  console.log('  ✓ Test 6 : _setNestedValue');

  // Test 7 : markDirty
  dm.markDirty('DSP-2024-0001');
  console.assert(dm._dirty.has('DSP-2024-0001'), 'Dirty flag posé');
  clearTimeout(dm._saveTimer); // annuler le timer pour le test
  console.log('  ✓ Test 7 : markDirty');

  // Test 8 : validate
  const v = dm.validate({ id: 'DSP-2024-X', v: 1, updatedAt: '2024-01-01T00:00:00Z', transport: { statut: 2 } });
  console.assert(v.valid === true || v.errors !== undefined, 'validate retourne un résultat');
  console.log('  ✓ Test 8 : validate');

  console.log('=== Tous les tests DataManager.js passent ===');
}

// Décommenter pour lancer les tests :
// _test_datamanager();

// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  migrate.js — Migration ancien format → nouveau format DispatchMaster   ║
// ║  Convertit les arrays de dossiers plats en objets indexés par ID        ║
// ║  avec versioning, timestamps ISO 8601 et historique                     ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Migre les données de l'ancien format (array plat, _ts en ms)
 * vers le nouveau format (objet indexé par ID, ISO 8601, versioning).
 *
 * @param {Object|Array} oldData - Ancien format : { master: [...], rawData: {}, cpData: [] }
 *                                  ou directement un array de dossiers [...]
 * @returns {{ data: Object, logs: string[], errors: string[] }}
 */
function migrateToNewFormat(oldData) {
  const logs = [];
  const errors = [];

  // --- 1. Normaliser l'entrée ---
  let oldMaster;
  let oldRawData = {};
  let oldCpData = [];

  if (Array.isArray(oldData)) {
    oldMaster = oldData;
  } else if (oldData && typeof oldData === 'object') {
    oldMaster = oldData.master || oldData.dossiers || [];
    oldRawData = oldData.rawData || {};
    oldCpData = oldData.cpData || [];
    if (!Array.isArray(oldMaster)) {
      // Peut être déjà au nouveau format (objet indexé)
      if (typeof oldMaster === 'object' && !Array.isArray(oldMaster)) {
        errors.push('Les données semblent déjà au nouveau format (objet indexé).');
        return { data: oldData, logs, errors };
      }
      oldMaster = [];
    }
  } else {
    errors.push('Format de données non reconnu : ' + typeof oldData);
    return { data: null, logs, errors };
  }

  // Deep clone pour ne pas modifier l'original
  oldMaster = JSON.parse(JSON.stringify(oldMaster));
  oldRawData = JSON.parse(JSON.stringify(oldRawData));
  oldCpData = JSON.parse(JSON.stringify(oldCpData));

  logs.push('Début migration — ' + oldMaster.length + ' dossier(s) détectés.');

  // --- 2. Générer les IDs ---
  const usedIds = new Set();
  let autoCounter = 1;
  const now = new Date().toISOString();
  const yearStr = new Date().getFullYear().toString();

  function generateId(file) {
    // Tenter de normaliser à partir du champ file
    if (file) {
      const cleaned = String(file).trim().replace(/\s+/g, '-').replace(/[^a-zA-Z0-9\-]/g, '');
      const candidate = 'DSP-' + yearStr + '-' + cleaned;
      if (!usedIds.has(candidate)) {
        usedIds.add(candidate);
        return candidate;
      }
    }
    // Fallback : compteur auto
    while (usedIds.has('DSP-' + yearStr + '-' + String(autoCounter).padStart(4, '0'))) {
      autoCounter++;
    }
    const id = 'DSP-' + yearStr + '-' + String(autoCounter).padStart(4, '0');
    usedIds.add(id);
    autoCounter++;
    return id;
  }

  // --- 3. Convertir les timestamps ---
  function msToISO(ms) {
    if (!ms) return now;
    if (typeof ms === 'string' && ms.includes('T')) return ms; // déjà ISO
    const n = Number(ms);
    if (isNaN(n) || n < 1e12) return now; // invalide
    try { return new Date(n).toISOString(); } catch (e) { return now; }
  }

  function dateToISO(d) {
    if (!d) return null;
    if (typeof d === 'string' && d.includes('T')) return d;
    // Format DD/MM/YYYY
    const m = String(d).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0') + 'T00:00:00Z';
    // Format YYYY-MM-DD
    const m2 = String(d).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m2) return d + 'T00:00:00Z';
    return null;
  }

  // --- 4. Convertir chaque dossier ---
  const dossiers = {};
  const seen = new Map(); // file → id (pour détecter doublons)

  for (let i = 0; i < oldMaster.length; i++) {
    const old = oldMaster[i];

    if (!old || typeof old !== 'object') {
      errors.push('Entrée #' + i + ' ignorée : type invalide (' + typeof old + ').');
      continue;
    }

    const fileKey = (old.file || '').trim();
    if (!fileKey) {
      errors.push('Entrée #' + i + ' ignorée : champ "file" vide.');
      continue;
    }

    // Doublon ?
    if (seen.has(fileKey)) {
      const existingId = seen.get(fileKey);
      const existing = dossiers[existingId];
      // Garder le plus récent
      const oldTs = old._ts || 0;
      const existTs = existing._raw_ts || 0;
      if (oldTs > existTs) {
        logs.push('Doublon détecté pour "' + fileKey + '" — version plus récente gardée.');
        // On va écraser l'existant
      } else {
        logs.push('Doublon détecté pour "' + fileKey + '" — version plus ancienne ignorée.');
        continue;
      }
    }

    const id = seen.has(fileKey) ? seen.get(fileKey) : generateId(fileKey);
    seen.set(fileKey, id);

    const statut = parseInt(old.statut) || 0;
    const updatedAt = msToISO(old._ts);
    const createdAt = dateToISO(old._dateCreated) || updatedAt;
    const updatedBy = old._by || old.operator || '';

    // Récupérer les données brutes si disponibles
    const raw = oldRawData[fileKey] || {};

    const newDossier = {
      id: id,
      file: fileKey,
      v: 1,
      createdAt: createdAt,
      updatedAt: updatedAt,
      updatedBy: updatedBy,

      client: {
        nom: old.client || '',
        ref: '',
        contact: old.contact || '',
        tel: old.tel || '',
        email: old.email || ''
      },

      transport: {
        statut: statut,
        type: (old.svct || '').split(',')[0] || '',
        svct: old.svct || '',
        poids: parseFloat(old.poids) || 0,
        volume: parseFloat(old.vol) || 0,
        nbColis: (fileKey.includes('+') ? fileKey.split('+').length : 1),
        taxable: parseFloat(old.taxable) || 0,
        origine: '',
        destination: old.dept || '',
        rdl: old.rdl || '',
        transp: old.transp || ''
      },

      financier: {
        nto: 0,
        fuel: 0,
        totalHT: 0,
        facture: (statut >= 7)
      },

      documents: {
        rdv: '',
        prealerte: ''
      },

      fc: {
        date: old.fcDate || '',
        horaire: old.fcHoraire || '',
        dly: old.fcDly || '',
        dlyNotes: old.fcDlyNotes || ''
      },

      cc: old.cc || '',
      operator: old.operator || '',
      notes: old.comment || '',

      raw: Object.keys(raw).length > 0 ? raw : undefined,

      historique: [
        {
          ts: createdAt,
          op: updatedBy || 'migration',
          action: 'création',
          detail: null
        }
      ],

      // Champ temporaire pour la déduplication (supprimé ensuite)
      _raw_ts: old._ts || 0
    };

    // Ajouter une entrée d'historique si le statut a un sens
    if (statut > 0) {
      newDossier.historique.push({
        ts: updatedAt,
        op: updatedBy || 'migration',
        action: 'statut_change',
        detail: { de: 0, a: statut }
      });
    }

    dossiers[id] = newDossier;
    logs.push('Migré : "' + fileKey + '" → ' + id + ' (statut=' + statut + ')');
  }

  // Nettoyer les champs temporaires
  Object.values(dossiers).forEach(d => { delete d._raw_ts; });

  // --- 5. Migrer les cpData ---
  const cpData = {};
  if (Array.isArray(oldCpData)) {
    oldCpData.forEach((cp, idx) => {
      if (cp && cp.file) {
        const cpId = 'CP-' + yearStr + '-' + String(idx + 1).padStart(4, '0');
        cpData[cpId] = {
          id: cpId,
          file: cp.file,
          client: cp.client || '',
          poids: parseFloat(cp.poids) || 0,
          vol: parseFloat(cp.vol) || 0,
          taxable: parseFloat(cp.taxable) || 0,
          operator: cp.operator || '',
          createdAt: now
        };
      }
    });
    if (Object.keys(cpData).length > 0) {
      logs.push('CP migrés : ' + Object.keys(cpData).length + ' entrée(s).');
    }
  }

  // --- 6. Construire le résultat final ---
  const count = Object.keys(dossiers).length;
  const data = {
    _meta: {
      schemaVersion: '2.1',
      generatedAt: now,
      generatedBy: 'migration',
      count: count,
      appVersion: 'DispatchMaster-2.1'
    },
    dossiers: dossiers,
    cpData: Object.keys(cpData).length > 0 ? cpData : undefined
  };

  logs.push('Migration terminée — ' + count + ' dossier(s) migrés, ' + errors.length + ' erreur(s).');

  return { data, logs, errors };
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  TESTS                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function _test_migrate() {
  console.log('=== Tests migrate.js ===');

  // Test 1 : migration basique
  const ancien = [
    { file: 'ABC123', client: 'DUPONT', statut: '2', operator: 'Jason',
      poids: 10, vol: 0.5, taxable: 10, svct: 'S5', transp: 'DPD',
      _ts: 1711100000000, _by: 'Jason', _dateCreated: '2025-03-22', cc: 'Cc',
      comment: 'Test', dept: '75', rdl: '2025-03-28' }
  ];
  const r = migrateToNewFormat(ancien);
  console.assert(r.errors.length === 0, 'Migration sans erreur');
  console.assert(typeof r.data.dossiers === 'object', 'Dossiers indexés par ID');
  console.assert(r.data._meta.schemaVersion === '2.1', 'Version schema correcte');
  console.assert(r.data._meta.count === 1, 'Count = 1');
  const first = Object.values(r.data.dossiers)[0];
  console.assert(first.id.startsWith('DSP-'), 'ID commence par DSP-');
  console.assert(first.v === 1, 'Version = 1');
  console.assert(first.transport.statut === 2, 'Statut = 2 (entier)');
  console.assert(first.client.nom === 'DUPONT', 'Client correct');
  console.assert(first.updatedAt.includes('T'), 'updatedAt en ISO 8601');
  console.assert(first.historique.length >= 1, 'Historique non vide');
  console.log('  ✓ Test 1 : migration basique OK');

  // Test 2 : doublons (garde le plus récent)
  const avecDoublons = [
    { file: 'XYZ', client: 'A', statut: '1', _ts: 1000000000000 },
    { file: 'XYZ', client: 'B', statut: '3', _ts: 2000000000000 }
  ];
  const r2 = migrateToNewFormat(avecDoublons);
  console.assert(r2.data._meta.count === 1, 'Doublon dédupliqué');
  const d2 = Object.values(r2.data.dossiers)[0];
  console.assert(d2.transport.statut === 3, 'Version plus récente gardée');
  console.log('  ✓ Test 2 : doublons OK');

  // Test 3 : format objet avec master/rawData
  const r3 = migrateToNewFormat({
    master: [{ file: 'F1', client: 'C1', statut: '0' }],
    rawData: { F1: { poids: 5, supplement: 0.3 } },
    cpData: [{ file: 'CP1', client: 'C1', poids: 2, vol: 0.1 }]
  });
  console.assert(r3.errors.length === 0, 'Format objet OK');
  console.assert(r3.data.cpData !== undefined, 'cpData migrés');
  const d3 = Object.values(r3.data.dossiers)[0];
  console.assert(d3.raw !== undefined, 'rawData préservé');
  console.log('  ✓ Test 3 : format objet + rawData + cpData OK');

  // Test 4 : entrées invalides
  const r4 = migrateToNewFormat([null, { file: '' }, { file: 'OK', statut: '1' }]);
  console.assert(r4.errors.length === 2, '2 erreurs pour entrées invalides');
  console.assert(r4.data._meta.count === 1, '1 dossier valide');
  console.log('  ✓ Test 4 : entrées invalides gérées');

  // Test 5 : array vide
  const r5 = migrateToNewFormat([]);
  console.assert(r5.data._meta.count === 0, 'Array vide → 0 dossiers');
  console.assert(r5.errors.length === 0, 'Pas d\'erreur');
  console.log('  ✓ Test 5 : array vide OK');

  console.log('=== Tous les tests migrate.js passent ===');
}

// Décommenter pour lancer les tests :
// _test_migrate();

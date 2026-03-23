// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  migrate.js — Migration ancien format → nouveau format DispatchMaster   ║
// ║  Convertit les arrays de dossiers plats en objets indexés par file      ║
// ║  avec versioning, timestamps ISO 8601 et historique                     ║
// ║                                                                          ║
// ║  Format réel des données existantes :                                     ║
// ║    file: "J1A0042031"         (ID unique = numéro de dossier)            ║
// ║    rdl: "24.03.26"            (date DD.MM.YY)                            ║
// ║    svct: "BX I2 S5 ST ZX"    (codes service séparés par espaces)        ║
// ║    transp: "Flex (7)"         (transporteur + numéro)                    ║
// ║    _ts: 1774266500445         (timestamp ms)                             ║
// ║    statut: "2"                (string "0"-"8")                           ║
// ║    _dateCreated: "2026-03-20" (ISO date)                                 ║
// ║    fcDate: "24.03.26"         (date DD.MM.YY)                            ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Migre les données de l'ancien format (array plat, _ts en ms)
 * vers le nouveau format (objet indexé par file, ISO 8601, versioning).
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

  // --- 2. Fonctions de conversion ---

  /** Convertit un timestamp ms en ISO 8601 */
  function msToISO(ms) {
    if (!ms) return null;
    if (typeof ms === 'string' && ms.includes('T')) return ms; // déjà ISO
    const n = Number(ms);
    if (isNaN(n) || n < 1e12) return null;
    try { return new Date(n).toISOString(); } catch (e) { return null; }
  }

  /** Convertit une date DD.MM.YY en ISO 8601 (ex: "24.03.26" → "2026-03-24T00:00:00Z") */
  function ddmmyyToISO(d) {
    if (!d || typeof d !== 'string') return null;
    // Format DD.MM.YY (ex: "24.03.26")
    const m1 = d.match(/^(\d{1,2})\.(\d{2})\.(\d{2})$/);
    if (m1) {
      const yy = parseInt(m1[3]);
      const year = yy >= 50 ? 1900 + yy : 2000 + yy; // 26 → 2026, 99 → 1999
      return year + '-' + m1[2] + '-' + m1[1].padStart(2, '0') + 'T00:00:00Z';
    }
    // Format DD/MM/YYYY
    const m2 = d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m2) return m2[3] + '-' + m2[2].padStart(2, '0') + '-' + m2[1].padStart(2, '0') + 'T00:00:00Z';
    // Format YYYY-MM-DD (déjà quasi-ISO)
    const m3 = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m3) return d + 'T00:00:00Z';
    return null;
  }

  /** Tente de corriger l'encodage UTF-8 cassé (mojibake) */
  function fixEncoding(s) {
    if (!s || typeof s !== 'string') return s || '';
    // Patterns courants de mojibake UTF-8
    return s
      .replace(/Ã©/g, 'é').replace(/Ã¨/g, 'è').replace(/Ãª/g, 'ê').replace(/Ã«/g, 'ë')
      .replace(/Ã /g, 'à').replace(/Ã¢/g, 'â').replace(/Ã¤/g, 'ä')
      .replace(/Ã¯/g, 'ï').replace(/Ã®/g, 'î')
      .replace(/Ã´/g, 'ô').replace(/Ã¶/g, 'ö')
      .replace(/Ã¹/g, 'ù').replace(/Ã»/g, 'û').replace(/Ã¼/g, 'ü')
      .replace(/Ã§/g, 'ç').replace(/Ã±/g, 'ñ')
      .replace(/Ã‰/g, 'É').replace(/Ã€/g, 'À')
      .replace(/ÃƒÂ«/g, 'ë').replace(/ÃƒÂ©/g, 'é')
      .replace(/ÃƒÂ¨/g, 'è').replace(/ÃƒÂ /g, 'à')
      .replace(/ÃƒÂ®/g, 'î').replace(/ÃƒÂ´/g, 'ô')
      .trim();
  }

  // --- 3. Convertir chaque dossier ---
  const now = new Date().toISOString();
  const dossiers = {};
  const seen = new Map(); // file → entrée la plus récente

  for (let i = 0; i < oldMaster.length; i++) {
    const old = oldMaster[i];

    if (!old || typeof old !== 'object') {
      errors.push('Entrée #' + i + ' ignorée : type invalide.');
      continue;
    }

    const fileKey = (old.file || '').trim();
    if (!fileKey) {
      errors.push('Entrée #' + i + ' ignorée : champ "file" vide.');
      continue;
    }

    // Doublon ? Garder le plus récent par _ts
    if (seen.has(fileKey)) {
      const prev = seen.get(fileKey);
      const oldTs = old._ts || 0;
      const prevTs = prev._ts || 0;
      if (oldTs > prevTs) {
        logs.push('Doublon "' + fileKey + '" — version plus récente gardée.');
        // On continue et on écrasera
      } else {
        logs.push('Doublon "' + fileKey + '" — version plus ancienne ignorée.');
        continue;
      }
    }
    seen.set(fileKey, old);

    const statut = parseInt(old.statut) || 0;
    const updatedAt = msToISO(old._ts) || now;
    const createdAt = ddmmyyToISO(old._dateCreated) || (old._dateCreated ? old._dateCreated + 'T00:00:00Z' : updatedAt);
    const updatedBy = old._by || old.operator || '';

    // Récupérer les données brutes si disponibles
    const raw = oldRawData[fileKey] || {};

    const newDossier = {
      id: fileKey,   // L'ID c'est le numéro de dossier (J1A0042031)
      file: fileKey,
      v: 1,
      createdAt: createdAt,
      updatedAt: updatedAt,
      updatedBy: updatedBy,

      client: {
        nom: old.client || '',
        ref: '',
        contact: fixEncoding(old.contact),
        tel: (old.tel || '').trim(),
        email: (old.email || '').trim()
      },

      transport: {
        statut: statut,
        svct: old.svct || '',
        transp: old.transp || '',
        poids: parseFloat(old.poids) || 0,
        volume: parseFloat(old.vol) || 0,
        taxable: parseFloat(old.taxable) || 0,
        nbColis: (fileKey.includes('+') ? fileKey.split('+').length : 1),
        destination: old.dept || '',
        rdl: old.rdl || '',
        rdlISO: ddmmyyToISO(old.rdl)
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
        dateISO: ddmmyyToISO(old.fcDate),
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
      ]
    };

    // Ajouter une entrée d'historique si le statut > 0
    if (statut > 0) {
      newDossier.historique.push({
        ts: updatedAt,
        op: updatedBy || 'migration',
        action: 'statut_change',
        detail: { de: 0, a: statut }
      });
    }

    dossiers[fileKey] = newDossier;
    logs.push('Migré : ' + fileKey + ' → ' + (old.client || '?') + ' (statut=' + statut + ', op=' + (old.operator || '—') + ')');
  }

  // --- 4. Migrer les cpData ---
  const cpData = {};
  if (Array.isArray(oldCpData)) {
    oldCpData.forEach((cp, idx) => {
      if (cp && cp.file) {
        cpData[cp.file] = {
          id: cp.file,
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

  // --- 5. Construire le résultat final ---
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
    cpData: Object.keys(cpData).length > 0 ? cpData : undefined,
    rawData: Object.keys(oldRawData).length > 0 ? oldRawData : undefined
  };

  logs.push('Migration terminée — ' + count + ' dossier(s) migrés, ' + errors.length + ' erreur(s).');

  return { data, logs, errors };
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  TESTS — basés sur les vraies données                                    ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function _test_migrate() {
  console.log('=== Tests migrate.js ===');

  // Test 1 : migration avec données réelles
  const ancien = {
    master: [
      { file: 'J1A0042031', client: 'THALES LAS SAS', rdl: '24.03.26',
        svct: 'BX I2 S5 ST ZX', transp: 'Flex (7)', poids: 73.5, vol: 0.358,
        taxable: 119.21, dept: '91', contact: '', tel: '06.12.92.87.13',
        email: '', cc: 'Cc', comment: '', statut: '8', operator: 'Abderrahamn',
        _dateCreated: '2026-03-20', _ts: 1774266500445, _by: 'Abderrahamn',
        fcDate: '24.03.26', fcHoraire: '09h et 12h', fcDly: '', fcDlyNotes: '' },
      { file: 'J1A0042096', client: 'ALSO FRANCE', rdl: '25.03.26',
        svct: 'S5', transp: 'Flex (7)', poids: 38.5, vol: 0.211,
        taxable: 70.26, dept: '95', cc: 'Cc', statut: '2', operator: 'Jason',
        _dateCreated: '2026-03-23' }
    ],
    rawData: { 'J1A0042031': { poids: 73.5, supplement: 0.2 } }
  };

  const r = migrateToNewFormat(ancien);
  console.assert(r.errors.length === 0, 'Migration sans erreur');
  console.assert(r.data._meta.schemaVersion === '2.1', 'Schema 2.1');
  console.assert(r.data._meta.count === 2, 'Count = 2');

  // Vérifier que file = ID (pas de DSP-YYYY-XXXX)
  const d1 = r.data.dossiers['J1A0042031'];
  console.assert(d1 !== undefined, 'Dossier trouvé par file key');
  console.assert(d1.id === 'J1A0042031', 'ID = file (J1A...)');
  console.assert(d1.transport.statut === 8, 'Statut converti en entier');
  console.assert(d1.transport.svct === 'BX I2 S5 ST ZX', 'SVCT préservé');
  console.assert(d1.transport.transp === 'Flex (7)', 'Transp préservé');
  console.assert(d1.client.tel === '06.12.92.87.13', 'Tel trimé');
  console.assert(d1.fc.date === '24.03.26', 'fcDate préservée');
  console.assert(d1.fc.dateISO === '2026-03-24T00:00:00Z', 'fcDate → ISO');
  console.assert(d1.transport.rdlISO === '2026-03-24T00:00:00Z', 'rdl → ISO');
  console.assert(d1.raw !== undefined, 'rawData préservé');
  console.assert(d1.operator === 'Abderrahamn', 'Opérateur correct');
  console.assert(d1.updatedAt.includes('T'), 'updatedAt en ISO');
  console.log('  ✓ Test 1 : données réelles migrées correctement');

  // Test 2 : dossier sans _ts (Jason, pas de modif réseau)
  const d2 = r.data.dossiers['J1A0042096'];
  console.assert(d2 !== undefined, 'Dossier J1A0042096 trouvé');
  console.assert(d2.transport.statut === 2, 'Statut 2');
  console.assert(d2.operator === 'Jason', 'Opérateur Jason');
  console.assert(d2.createdAt === '2026-03-23T00:00:00Z', '_dateCreated → ISO');
  console.log('  ✓ Test 2 : dossier sans _ts OK');

  // Test 3 : doublons
  const avecDoublons = [
    { file: 'J1A0042031', client: 'A', statut: '1', _ts: 1000000000000 },
    { file: 'J1A0042031', client: 'B', statut: '3', _ts: 2000000000000 }
  ];
  const r3 = migrateToNewFormat(avecDoublons);
  console.assert(r3.data._meta.count === 1, 'Doublon dédupliqué');
  console.assert(r3.data.dossiers['J1A0042031'].transport.statut === 3, 'Version récente gardée');
  console.log('  ✓ Test 3 : doublons OK');

  // Test 4 : fix encodage mojibake
  const r4 = migrateToNewFormat([
    { file: 'TEST', client: 'X', statut: '1', contact: 'RaphaÃƒÂ«l MAURICE' }
  ]);
  console.assert(r4.data.dossiers['TEST'].client.contact === 'Raphaël MAURICE', 'Mojibake corrigé');
  console.log('  ✓ Test 4 : fix encodage mojibake');

  // Test 5 : conversion dates DD.MM.YY
  const r5 = migrateToNewFormat([
    { file: 'D1', client: 'X', statut: '1', rdl: '25.03.26', fcDate: '24.03.26' }
  ]);
  const d5 = r5.data.dossiers['D1'];
  console.assert(d5.transport.rdlISO === '2026-03-25T00:00:00Z', 'rdl DD.MM.YY → ISO');
  console.assert(d5.fc.dateISO === '2026-03-24T00:00:00Z', 'fcDate DD.MM.YY → ISO');
  console.log('  ✓ Test 5 : dates DD.MM.YY → ISO');

  // Test 6 : tel avec espaces trailing
  const r6 = migrateToNewFormat([
    { file: 'T1', client: 'X', statut: '1', tel: '0612345678   ' }
  ]);
  console.assert(r6.data.dossiers['T1'].client.tel === '0612345678', 'Tel trimé');
  console.log('  ✓ Test 6 : trim tel/email');

  // Test 7 : entrées invalides
  const r7 = migrateToNewFormat([null, { file: '' }, { file: 'OK', statut: '1' }]);
  console.assert(r7.errors.length === 2, '2 erreurs');
  console.assert(r7.data._meta.count === 1, '1 valide');
  console.log('  ✓ Test 7 : entrées invalides gérées');

  console.log('=== Tous les tests migrate.js passent ===');
}

// Décommenter pour lancer les tests :
// _test_migrate();

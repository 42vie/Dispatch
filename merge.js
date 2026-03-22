// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  merge.js — Fusion réseau multi-opérateurs DispatchMaster               ║
// ║  Merge les fichiers state_{Nom}.json en un fichier de base unique       ║
// ║  Résolution de conflits par version (v) puis par date (updatedAt)       ║
// ╚══════════════════════════════════════════════════════════════════════════╝

/**
 * Compare deux dossiers et détermine le gagnant.
 *
 * Règles :
 *  1. Version supérieure gagne
 *  2. À version égale, le plus récent (updatedAt) gagne
 *  3. Même version + même timestamp = conflit réel → on garde le local + on loggue
 *
 * @param {Object} local   - Dossier existant dans la base
 * @param {Object} distant - Dossier provenant d'un fichier opérateur
 * @returns {{ winner: Object, reason: string, conflict?: { local: Object, distant: Object } }}
 */
function mergeDossier(local, distant) {
  const lv = parseInt(local.v) || 0;
  const dv = parseInt(distant.v) || 0;

  if (dv > lv) return { winner: distant, reason: 'version_superieure' };
  if (lv > dv) return { winner: local, reason: 'version_superieure' };

  // Même version — comparer updatedAt
  const lDate = local.updatedAt || '';
  const dDate = distant.updatedAt || '';

  if (dDate > lDate) return { winner: distant, reason: 'plus_recent' };
  if (lDate > dDate) return { winner: local, reason: 'plus_recent' };

  // Conflit réel (même version, même timestamp)
  return { winner: local, reason: 'conflit_reel', conflict: { local, distant } };
}


/**
 * Fusionne un fichier de base avec plusieurs fichiers opérateurs.
 *
 * @param {Object} base            - Fichier fusionné existant (nouveau format avec _meta + dossiers)
 *                                    Peut être null pour un premier merge.
 * @param {Array<Object>} operateurFiles - Array de fichiers opérateurs,
 *                                          chacun au format { _meta: {...}, dossiers: {...} }
 * @returns {{
 *   merged: Object,
 *   conflicts: Array<{ id: string, local: Object, distant: Object, operateur: string }>,
 *   stats: { total: number, updated: number, conflicts: number, added: number }
 * }}
 */
function mergeNetworkStates(base, operateurFiles) {
  const conflicts = [];
  const stats = { total: 0, updated: 0, conflicts: 0, added: 0 };

  // Normaliser la base
  let baseDossiers = {};
  if (base && base.dossiers && typeof base.dossiers === 'object') {
    baseDossiers = JSON.parse(JSON.stringify(base.dossiers)); // deep clone
  }

  // Support ancien format (array) pour compatibilité
  if (Array.isArray(base)) {
    base.forEach(r => {
      const key = r.id || r.file || '';
      if (key) baseDossiers[key] = r;
    });
  }

  // Itérer sur chaque fichier opérateur
  if (!Array.isArray(operateurFiles)) operateurFiles = [];

  for (const opFile of operateurFiles) {
    if (!opFile || typeof opFile !== 'object') continue;

    const opName = (opFile._meta && opFile._meta.operateur) || 'inconnu';
    let opDossiers = {};

    // Support nouveau format (objet indexé) et ancien format (array)
    if (opFile.dossiers && typeof opFile.dossiers === 'object' && !Array.isArray(opFile.dossiers)) {
      opDossiers = opFile.dossiers;
    } else if (Array.isArray(opFile.dossiers || opFile)) {
      const arr = opFile.dossiers || opFile;
      arr.forEach(r => {
        const key = r.id || r.file || '';
        if (key) opDossiers[key] = r;
      });
    }

    for (const [id, distant] of Object.entries(opDossiers)) {
      const local = baseDossiers[id];

      if (!local) {
        // Nouveau dossier — ajouter directement
        baseDossiers[id] = JSON.parse(JSON.stringify(distant));
        stats.added++;
        continue;
      }

      // Merge
      const result = mergeDossier(local, distant);

      if (result.reason === 'conflit_reel') {
        // Conflit réel — garder le local + stocker le conflit
        conflicts.push({
          id: id,
          local: result.conflict.local,
          distant: result.conflict.distant,
          operateur: opName
        });

        // Sauvegarder la version conflictuelle sous un ID suffixé
        const conflictId = id + '_conflict_' + opName;
        baseDossiers[conflictId] = JSON.parse(JSON.stringify(distant));
        baseDossiers[conflictId].id = conflictId;
        baseDossiers[conflictId]._conflictSource = {
          originalId: id,
          operateur: opName,
          detectedAt: new Date().toISOString()
        };

        stats.conflicts++;
      } else if (result.winner === distant) {
        // Le distant gagne — remplacer
        baseDossiers[id] = JSON.parse(JSON.stringify(distant));
        stats.updated++;
      }
      // Si local gagne, on ne fait rien (il est déjà dans baseDossiers)
    }
  }

  stats.total = Object.keys(baseDossiers).length;

  // Construire le résultat
  const merged = {
    _meta: {
      schemaVersion: (base && base._meta && base._meta.schemaVersion) || '2.1',
      generatedAt: new Date().toISOString(),
      generatedBy: 'merge',
      count: stats.total,
      appVersion: 'DispatchMaster-2.1',
      mergeStats: stats
    },
    dossiers: baseDossiers
  };

  return { merged, conflicts, stats };
}


/**
 * Merge simplifié pour l'ancien format (arrays avec file + _ts).
 * Wrapper de compatibilité pour l'interface existante.
 *
 * @param {Array} baseArray       - Array de dossiers (ancien format)
 * @param {Array<Array>} opArrays - Array d'arrays d'opérateurs
 * @returns {{ merged: Array, conflicts: Array, stats: Object }}
 */
function mergeNetworkStatesLegacy(baseArray, opArrays) {
  // Convertir en nouveau format temporaire
  function arrayToNewFormat(arr, opName) {
    const dossiers = {};
    (arr || []).forEach(r => {
      const key = r.file || r.id || '';
      if (!key) return;
      dossiers[key] = {
        id: key,
        v: r.v || 1,
        updatedAt: r.updatedAt || (r._ts ? new Date(r._ts).toISOString() : ''),
        updatedBy: r._by || r.operator || opName || '',
        statut: r.statut,
        operator: r.operator || '',
        file: r.file || ''
      };
    });
    return { _meta: { operateur: opName || '' }, dossiers };
  }

  const base = { dossiers: {} };
  (baseArray || []).forEach(r => {
    const key = r.file || r.id || '';
    if (key) base.dossiers[key] = {
      ...r, id: key, v: r.v || 1,
      updatedAt: r.updatedAt || (r._ts ? new Date(r._ts).toISOString() : '')
    };
  });

  const opFiles = (opArrays || []).map((arr, i) => arrayToNewFormat(arr, 'op_' + i));

  const result = mergeNetworkStates(base, opFiles);

  // Reconvertir en array pour compatibilité
  const mergedArray = Object.values(result.merged.dossiers);
  return { merged: mergedArray, conflicts: result.conflicts, stats: result.stats };
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  TESTS                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function _test_merge() {
  console.log('=== Tests merge.js ===');

  // Test 1 : version supérieure gagne
  const r1 = mergeDossier(
    { id: 'A', v: 2, updatedAt: '2024-03-22T10:00:00Z' },
    { id: 'A', v: 5, updatedAt: '2024-03-21T10:00:00Z' }
  );
  console.assert(r1.winner.v === 5, 'Version supérieure gagne');
  console.assert(r1.reason === 'version_superieure', 'Raison correcte');
  console.log('  ✓ Test 1 : version supérieure gagne');

  // Test 2 : même version, plus récent gagne
  const r2 = mergeDossier(
    { id: 'A', v: 3, updatedAt: '2024-03-22T10:00:00Z' },
    { id: 'A', v: 3, updatedAt: '2024-03-22T14:00:00Z' }
  );
  console.assert(r2.winner.updatedAt === '2024-03-22T14:00:00Z', 'Plus récent gagne');
  console.assert(r2.reason === 'plus_recent', 'Raison = plus_recent');
  console.log('  ✓ Test 2 : même version, plus récent gagne');

  // Test 3 : conflit réel
  const r3 = mergeDossier(
    { id: 'A', v: 3, updatedAt: '2024-03-22T10:00:00Z', data: 'local' },
    { id: 'A', v: 3, updatedAt: '2024-03-22T10:00:00Z', data: 'distant' }
  );
  console.assert(r3.reason === 'conflit_reel', 'Conflit réel détecté');
  console.assert(r3.conflict !== undefined, 'Détails du conflit présents');
  console.assert(r3.winner.data === 'local', 'Local gardé en cas de conflit');
  console.log('  ✓ Test 3 : conflit réel');

  // Test 4 : mergeNetworkStates complet
  const base = {
    _meta: { schemaVersion: '2.1' },
    dossiers: {
      'D1': { id: 'D1', v: 2, updatedAt: '2024-03-22T10:00:00Z', updatedBy: 'Jason' },
      'D2': { id: 'D2', v: 1, updatedAt: '2024-03-22T08:00:00Z', updatedBy: 'Jason' }
    }
  };
  const opFiles = [
    {
      _meta: { operateur: 'Marie' },
      dossiers: {
        'D1': { id: 'D1', v: 3, updatedAt: '2024-03-22T12:00:00Z', updatedBy: 'Marie' },
        'D3': { id: 'D3', v: 1, updatedAt: '2024-03-22T11:00:00Z', updatedBy: 'Marie' }
      }
    }
  ];
  const r4 = mergeNetworkStates(base, opFiles);
  console.assert(r4.merged.dossiers['D1'].v === 3, 'D1 mis à jour (v3 > v2)');
  console.assert(r4.merged.dossiers['D1'].updatedBy === 'Marie', 'D1 updatedBy Marie');
  console.assert(r4.merged.dossiers['D2'].v === 1, 'D2 inchangé');
  console.assert(r4.merged.dossiers['D3'] !== undefined, 'D3 ajouté');
  console.assert(r4.stats.updated === 1, '1 mis à jour');
  console.assert(r4.stats.added === 1, '1 ajouté');
  console.assert(r4.stats.conflicts === 0, '0 conflits');
  console.assert(r4.stats.total === 3, '3 total');
  console.log('  ✓ Test 4 : merge complet OK');

  // Test 5 : conflit réel dans merge
  const base5 = {
    dossiers: {
      'X': { id: 'X', v: 2, updatedAt: '2024-03-22T10:00:00Z' }
    }
  };
  const op5 = [{
    _meta: { operateur: 'Bob' },
    dossiers: {
      'X': { id: 'X', v: 2, updatedAt: '2024-03-22T10:00:00Z' }
    }
  }];
  const r5 = mergeNetworkStates(base5, op5);
  console.assert(r5.stats.conflicts === 1, '1 conflit');
  console.assert(r5.conflicts.length === 1, 'Conflit logué');
  console.assert(r5.conflicts[0].operateur === 'Bob', 'Opérateur du conflit');
  console.assert(r5.merged.dossiers['X_conflict_Bob'] !== undefined, 'Version conflit sauvegardée');
  console.log('  ✓ Test 5 : conflit réel dans merge');

  // Test 6 : base null (premier merge)
  const r6 = mergeNetworkStates(null, [{
    _meta: { operateur: 'First' },
    dossiers: { 'A': { id: 'A', v: 1, updatedAt: '2024-01-01T00:00:00Z' } }
  }]);
  console.assert(r6.stats.added === 1, '1 ajouté sur base vide');
  console.assert(r6.stats.total === 1, '1 total');
  console.log('  ✓ Test 6 : base null OK');

  // Test 7 : mergeNetworkStatesLegacy (ancien format arrays)
  const r7 = mergeNetworkStatesLegacy(
    [{ file: 'F1', statut: '1', _ts: 1711100000000 }],
    [[{ file: 'F1', statut: '3', _ts: 1711200000000 }, { file: 'F2', statut: '1', _ts: 1711200000000 }]]
  );
  console.assert(Array.isArray(r7.merged), 'Legacy retourne un array');
  console.assert(r7.merged.length === 2, '2 dossiers mergés');
  console.log('  ✓ Test 7 : legacy merge OK');

  console.log('=== Tous les tests merge.js passent ===');
}

// Décommenter pour lancer les tests :
// _test_merge();

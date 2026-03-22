// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  validate.js — Validation JSON pour dossiers DispatchMaster             ║
// ║  Vérifie la conformité des dossiers au schéma v2.1                      ║
// ╚══════════════════════════════════════════════════════════════════════════╝

const ID_REGEX = /^DSP-\d{4}-.+$/;
const ISO_REGEX = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?Z$/;

/**
 * Valide un dossier individuel au format v2.1.
 *
 * @param {Object} dossier - Dossier à valider
 * @returns {{ valid: boolean, errors: string[], warnings: string[] }}
 */
function validateDossier(dossier) {
  const errors = [];
  const warnings = [];

  if (!dossier || typeof dossier !== 'object') {
    return { valid: false, errors: ['Le dossier n\'est pas un objet.'], warnings };
  }

  // --- Champs obligatoires ---
  if (!dossier.id || typeof dossier.id !== 'string') {
    errors.push('Champ "id" manquant ou invalide.');
  } else if (!ID_REGEX.test(dossier.id)) {
    errors.push('Format ID invalide : "' + dossier.id + '" (attendu : DSP-YYYY-XXXX).');
  }

  if (dossier.v === undefined || dossier.v === null) {
    errors.push('Champ "v" (version) manquant.');
  } else if (!Number.isInteger(dossier.v) || dossier.v < 1) {
    errors.push('"v" doit être un entier >= 1 (reçu : ' + dossier.v + ').');
  }

  if (!dossier.updatedAt) {
    errors.push('Champ "updatedAt" manquant.');
  } else if (!ISO_REGEX.test(dossier.updatedAt)) {
    // Accepter aussi le format partiel YYYY-MM-DDTHH:MM:SSZ (sans ms)
    warnings.push('"updatedAt" n\'est pas au format ISO 8601 strict : "' + dossier.updatedAt + '".');
  }

  if (!dossier.updatedBy) {
    warnings.push('Champ "updatedBy" manquant ou vide.');
  }

  // --- Transport ---
  if (!dossier.transport || typeof dossier.transport !== 'object') {
    errors.push('Section "transport" manquante.');
  } else {
    const st = dossier.transport.statut;
    if (st === undefined || st === null) {
      errors.push('Champ "transport.statut" manquant.');
    } else {
      const statInt = parseInt(st);
      if (isNaN(statInt) || statInt < 0 || statInt > 8) {
        errors.push('"transport.statut" doit être entre 0 et 8 (reçu : ' + st + ').');
      }
    }

    if (dossier.transport.poids !== undefined && typeof dossier.transport.poids !== 'number') {
      warnings.push('"transport.poids" devrait être un nombre.');
    }
    if (dossier.transport.volume !== undefined && typeof dossier.transport.volume !== 'number') {
      warnings.push('"transport.volume" devrait être un nombre.');
    }
  }

  // --- Client ---
  if (dossier.client && typeof dossier.client === 'object') {
    if (!dossier.client.nom) {
      warnings.push('"client.nom" est vide.');
    }
  }

  // --- Historique ---
  if (dossier.historique !== undefined) {
    if (!Array.isArray(dossier.historique)) {
      errors.push('"historique" doit être un array.');
    }
  }

  // --- Dates ---
  if (dossier.createdAt && !ISO_REGEX.test(dossier.createdAt)) {
    warnings.push('"createdAt" n\'est pas au format ISO 8601 strict.');
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings
  };
}


/**
 * Valide un fichier complet (avec _meta et dossiers).
 *
 * @param {Object} file - Fichier au format { _meta: {...}, dossiers: {...} }
 * @returns {{ valid: boolean, errors: string[], warnings: string[], details: Object }}
 */
function validateFile(file) {
  const errors = [];
  const warnings = [];
  const details = {};

  if (!file || typeof file !== 'object') {
    return { valid: false, errors: ['Le fichier n\'est pas un objet.'], warnings, details };
  }

  // _meta
  if (!file._meta) {
    errors.push('Section "_meta" manquante.');
  } else {
    if (file._meta.schemaVersion !== '2.1') {
      warnings.push('schemaVersion attendu "2.1", reçu "' + file._meta.schemaVersion + '".');
    }
    if (!file._meta.generatedAt) {
      warnings.push('"_meta.generatedAt" manquant.');
    }
  }

  // dossiers
  if (!file.dossiers || typeof file.dossiers !== 'object') {
    errors.push('Section "dossiers" manquante ou invalide.');
    return { valid: false, errors, warnings, details };
  }

  let validCount = 0;
  let invalidCount = 0;
  const dossierErrors = {};

  for (const [id, dossier] of Object.entries(file.dossiers)) {
    const result = validateDossier(dossier);
    if (result.valid) {
      validCount++;
    } else {
      invalidCount++;
      dossierErrors[id] = result.errors;
    }
    if (result.warnings.length > 0) {
      if (!dossierErrors[id]) dossierErrors[id] = [];
      dossierErrors[id].push(...result.warnings.map(w => '[warn] ' + w));
    }

    // Vérifier cohérence ID clé ↔ dossier.id
    if (dossier.id && dossier.id !== id) {
      warnings.push('ID incohérent pour "' + id + '" : dossier.id = "' + dossier.id + '".');
    }
  }

  // Vérifier count
  const actualCount = Object.keys(file.dossiers).length;
  if (file._meta && file._meta.count !== undefined && file._meta.count !== actualCount) {
    warnings.push('_meta.count (' + file._meta.count + ') ne correspond pas au nombre réel (' + actualCount + ').');
  }

  details.validCount = validCount;
  details.invalidCount = invalidCount;
  details.totalCount = actualCount;
  details.dossierErrors = Object.keys(dossierErrors).length > 0 ? dossierErrors : undefined;

  return {
    valid: errors.length === 0 && invalidCount === 0,
    errors,
    warnings,
    details
  };
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  TESTS                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function _test_validate() {
  console.log('=== Tests validate.js ===');

  // Test 1 : dossier valide
  const r1 = validateDossier({
    id: 'DSP-2024-0001',
    v: 3,
    createdAt: '2024-03-15T08:30:00Z',
    updatedAt: '2024-03-22T14:12:00Z',
    updatedBy: 'Jason',
    transport: { statut: 2, poids: 10, volume: 0.5 },
    client: { nom: 'DUPONT' },
    historique: []
  });
  console.assert(r1.valid === true, 'Dossier valide reconnu');
  console.assert(r1.errors.length === 0, 'Aucune erreur');
  console.log('  ✓ Test 1 : dossier valide');

  // Test 2 : ID invalide
  const r2 = validateDossier({ id: 'BLABLA', v: 1, updatedAt: '2024-03-22T14:12:00Z', transport: { statut: 0 } });
  console.assert(r2.valid === false, 'ID invalide détecté');
  console.assert(r2.errors.some(e => e.includes('Format ID')), 'Erreur ID');
  console.log('  ✓ Test 2 : ID invalide');

  // Test 3 : champs manquants
  const r3 = validateDossier({});
  console.assert(r3.valid === false, 'Champs manquants détectés');
  console.assert(r3.errors.length >= 3, 'Au moins 3 erreurs (id, v, transport)');
  console.log('  ✓ Test 3 : champs manquants');

  // Test 4 : statut hors bornes
  const r4 = validateDossier({ id: 'DSP-2024-X', v: 1, updatedAt: '2024-03-22T14:12:00Z', transport: { statut: 15 } });
  console.assert(r4.errors.some(e => e.includes('entre 0 et 8')), 'Statut hors bornes');
  console.log('  ✓ Test 4 : statut hors bornes');

  // Test 5 : version invalide
  const r5 = validateDossier({ id: 'DSP-2024-X', v: 0, updatedAt: '2024-01-01T00:00:00Z', transport: { statut: 1 } });
  console.assert(r5.errors.some(e => e.includes('>= 1')), 'Version < 1 détectée');
  console.log('  ✓ Test 5 : version invalide');

  // Test 6 : validateFile complet
  const r6 = validateFile({
    _meta: { schemaVersion: '2.1', count: 1 },
    dossiers: {
      'DSP-2024-0001': {
        id: 'DSP-2024-0001', v: 1, updatedAt: '2024-01-01T00:00:00Z',
        updatedBy: 'test', transport: { statut: 0 }, historique: []
      }
    }
  });
  console.assert(r6.valid === true, 'Fichier valide');
  console.assert(r6.details.validCount === 1, '1 dossier valide');
  console.log('  ✓ Test 6 : validateFile complet');

  // Test 7 : null
  const r7 = validateDossier(null);
  console.assert(r7.valid === false, 'null détecté');
  console.log('  ✓ Test 7 : null');

  console.log('=== Tous les tests validate.js passent ===');
}

// Décommenter pour lancer les tests :
// _test_validate();

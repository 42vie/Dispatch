// ║  RÉSEAU — SAVE / LOAD SUR F:\                                            ║
// ╚══════════════════════════════════════════════════════════════════════════╝
async function reseauForceSave() {
  const path = localStorage.getItem('dispatch_state_path') || DEFAULT_STATE_PATH;
  const myName = localStorage.getItem('dispatch_operator') || '';
  if (!path) return toast('Chemin réseau non configuré. Aller dans Options → Réseau partagé.');
  if (!myName) return toast('Nom opérateur non configuré. Aller dans Options → Réseau partagé.');
  try {
    await smartNetSave(path);
    const st = document.getElementById('opts-reseau-status');
    if (st) st.textContent = '✓ Sauvegardé → ' + _opPath(path, myName).split('\\').pop() + ' — ' + new Date().toLocaleTimeString();
    toast('État sauvegardé : ' + _opPath(path, myName).split('\\').pop());
  } catch(e) { toast('Erreur sauvegarde réseau : ' + e.message); }
}

async function reseauForceLoad() {
  const path = localStorage.getItem('dispatch_state_path') || DEFAULT_STATE_PATH;
  if (!path) return toast('Chemin réseau non configuré.');
  try {
    // 1. Lister tous les fichiers opérateur
    const pattern = _opPattern(path);
    const files = await netListFiles(pattern);

    if (!files.length) {
      // Fallback : charger le fichier de base (somme des opérateurs)
      const oldData = await netLoad(path);
      if (oldData) {
        applyLoadedState(Array.isArray(oldData) ? oldData : (oldData.board || oldData.master || []));
        toast('Réseau : chargé depuis le fichier de base.');
        return;
      }
      return toast('Aucun fichier opérateur trouvé.');
    }

    // 2. Charger tous les fichiers et fusionner
    const allDossiers = [];
    const operatorFiles = [];
    for (const f of files) {
      const fNorm = f.replace(/\\\\/g, '\\');
      const data = await netLoad(fNorm);
      if (Array.isArray(data)) {
        allDossiers.push(...data);
        // Extraire le nom d'opérateur du fichier
        const match = fNorm.match(/dispatch_state_(.+)\.json$/i);
        if (match) operatorFiles.push(match[1]);
      }
    }

    // 3. Dédupliquer par file (garder le plus récent _ts)
    const merged = new Map();
    allDossiers.forEach(r => {
      const existing = merged.get(r.file);
      if (!existing || (r._ts || 0) >= (existing._ts || 0)) {
        merged.set(r.file, r);
      }
    });

    applyLoadedState([...merged.values()]);
    toast(`Réseau : ${merged.size} dossier(s) chargés depuis ${files.length} opérateur(s) (${operatorFiles.join(', ')}).`);
  } catch(e) { toast('Erreur chargement réseau : ' + e.message); }
}

// Sauvegarde par opérateur : chaque opérateur écrit SON fichier uniquement
// dispatch_state_Jason.json, dispatch_state_Abderrahamn.json, etc.
async function smartNetSave(path) {
  const myName = localStorage.getItem('dispatch_operator') || '';
  if (!myName) return; // pas de nom = pas de sauvegarde réseau

  try {
    // 1. Assigner l'opérateur aux dossiers sans propriétaire (en mémoire, pas juste la copie)
    g_master.forEach(r => { if (!r.operator && myName) r.operator = myName; });

    // 2. Construire l'état local (seulement MES dossiers)
    const myBoard = g_master
      .filter(r => (r.operator || '') === myName)
      .map(r => ({
        file: r.file, client: r.client || '', rdl: r.rdl || '',
        svct: r.svct || '', transp: r.transp || '',
        poids: r.poids || 0, vol: r.vol || 0, taxable: r.taxable || 0,
        dept: r.dept || '', contact: r.contact || '', tel: r.tel || '', email: r.email || '',
        cc: r.cc || '', comment: r.comment || '',
        statut: r.statut, operator: r.operator || myName,
        _dateCreated: r._dateCreated || '',
        _ts: r._ts || Date.now(), _by: myName,
        fcDate: r.fcDate || '', fcHoraire: r.fcHoraire || '',
        fcDly: r.fcDly || '', fcDlyNotes: r.fcDlyNotes || ''
      }));

    // 2. Sauvegarder dans mon fichier opérateur
    const myPath = _opPath(path, myName);
    await netSave(myPath, myBoard);

    // 3. Charger les fichiers des AUTRES opérateurs pour syncer l'interface
    const pattern = _opPattern(path);
    const files = await netListFiles(pattern);
    let synced = 0;
    for (const f of files) {
      // Normaliser les backslashes pour comparer
      const fNorm = f.replace(/\\\\/g, '\\');
      const myNorm = myPath.replace(/\\\\/g, '\\');
      if (fNorm === myNorm) continue; // c'est mon propre fichier
      const remote = await netLoad(fNorm);
      if (!Array.isArray(remote)) continue;
      remote.forEach(rd => {
        const local = g_master.find(r => r.file === rd.file);
        if (local && rd.operator && rd.operator !== myName) {
          // Dossier d'un autre opérateur — syncer si plus récent
          if ((rd._ts || 0) > (local._ts || 0)) {
            local.statut = rd.statut;
            local.operator = rd.operator;
            local._ts = rd._ts;
            ['client','rdl','svct','transp','poids','vol','taxable','dept',
             'contact','tel','email','cc','comment','_dateCreated',
             'fcDate','fcHoraire','fcDly','fcDlyNotes'].forEach(k => {
              if (rd[k] !== undefined && rd[k] !== '') local[k] = rd[k];
            });
            synced++;
          }
        }
      });
    }

    if (synced > 0) {
      markDirty();  // → déclenche autoSave → écrit dispatch_status + dispatch_data
      renderMaster(); renderKanban(); updateStats();
      toast(`Sync réseau : ${synced} dossier(s) mis à jour depuis un autre opérateur.`);
    }

    // 4. Reconstruire le fichier de base = somme de tous les opérateurs
    try {
      const allOps = [...myBoard]; // commencer avec mes propres données (toujours disponibles)
      for (const f of files) {
        // Normaliser : retirer les doubles backslash éventuels
        const fNorm = f.replace(/\\\\/g, '\\').toLowerCase();
        const myNorm = myPath.replace(/\\\\/g, '\\').toLowerCase();
        if (fNorm === myNorm) continue; // déjà inclus via myBoard
        const data = await netLoad(f); // utiliser le chemin original (pas normalisé)
        if (Array.isArray(data)) allOps.push(...data);
      }
      // Dédupliquer par file (garder le plus récent _ts)
      const merged = new Map();
      allOps.forEach(r => {
        const existing = merged.get(r.file);
        if (!existing || (r._ts || 0) >= (existing._ts || 0)) {
          merged.set(r.file, r);
        }
      });
      console.log(`smartNetSave base rebuild: ${files.length} fichiers, ${allOps.length} dossiers, ${merged.size} uniques`);
      await netSave(path, [...merged.values()]);
    } catch(e2) { console.warn('smartNetSave base rebuild:', e2); }

  } catch(e) {
    console.warn('smartNetSave error:', e);
  }
}

// Au load : on écrase le statut local avec celui du réseau (si le dossier existe)
function applyLoadedState(state) {
  const board = Array.isArray(state) ? state : (state.board || state.master || []);
  let synced = 0;
  board.forEach(remote => {
    const local = g_master.find(r => r.file === remote.file);
    if (!local) return;
    // Syncer UNIQUEMENT si le remote est STRICTEMENT plus récent (pas >=, sinon on régresse)
    const remoteNewer = (remote._ts || 0) > (local._ts || 0);
    const changed = local.statut !== remote.statut || (remote.operator && local.operator !== remote.operator);
    if (changed && remoteNewer) {
      local.statut = remote.statut;
      if (remote.operator) local.operator = remote.operator;
      local._ts = remote._ts;
      ['client','rdl','svct','transp','poids','vol','taxable','dept',
       'contact','tel','email','cc','comment','_dateCreated',
       'fcDate','fcHoraire','fcDly','fcDlyNotes'].forEach(k => {
        if (remote[k] !== undefined && remote[k] !== '') local[k] = remote[k];
      });
      synced++;
    }
  });
  if (synced > 0) {
    markDirty();  // → déclenche autoSave → écrit dispatch_status + dispatch_data
    renderMaster(); renderKanban(); updateStats();
    console.log(`Réseau : ${synced} dossier(s) synchronisés.`);
  }
}

// Init au chargement
window.addEventListener('load', () => {
  setTimeout(() => {
    optsLoadPJ();
    optsLoadReseau();
    optsLoadFuel();
    cpCfgLoad();
    populateOperatorFilter();
    showIdentityModal();
    // Charger l'état réseau au démarrage (contacts inclus dans le même fichier)
    const path = localStorage.getItem('dispatch_state_path') || DEFAULT_STATE_PATH;
    if (path) reseauForceLoad();
  }, 800); // après que g_master soit peuplé depuis AutoIt
});


// ╔══════════════════════════════════════════════════════════════════════════╗

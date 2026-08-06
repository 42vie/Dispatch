// ║  REFRESH / SYNC DEPUIS SERVEUR (sans recharger la page)                  ║
// ╚══════════════════════════════════════════════════════════════════════════╝
let _refreshing = false;
async function refreshData(silent) {
  if (_refreshing) return;
  _refreshing = true;
  if (!silent) toast('Rafraîchissement en cours...');
  try {
    // Sauvegarder les modifications locales non sauvées d'abord
    await autoSave();

    // Charger depuis l'API
    let freshData = null;
    try {
      const resp = await fetch(`${API_URL}/load-data`);
      freshData = await resp.json();
    } catch(e) {
      try {
        const resp = await fetch(`${API_URL}/load`);
        freshData = await resp.json();
      } catch(e2) { /* ignore */ }
    }

    if (!freshData || !freshData.master || !freshData.master.length) {
      if (!silent) toast('Aucune donnée reçue du serveur.');
      _refreshing = false;
      return;
    }

    const myName = localStorage.getItem('dispatch_operator') || '';

    // L'API est la source de vérité : on réconcilie g_master avec le serveur
    const serverMap = new Map();
    freshData.master.forEach(r => serverMap.set(r.file, r));

    // Retirer les dossiers locaux qui n'existent plus sur le serveur
    const beforeCount = g_master.length;
    g_master = g_master.filter(r => serverMap.has(r.file));
    const removed = beforeCount - g_master.length;

    // Mettre à jour les dossiers existants avec les données serveur si plus récentes
    let updated = 0;
    g_master.forEach(r => {
      const remote = serverMap.get(r.file);
      if (!remote) return;
      serverMap.delete(r.file); // traité
      if ((remote._ts || 0) > (r._ts || 0)) {
        Object.keys(remote).forEach(k => { r[k] = remote[k]; });
        updated++;
      }
    });

    // Ajouter les vrais nouveaux dossiers (ajoutés par un autre opérateur/import)
    let added = 0;
    serverMap.forEach(r => {
      g_master.push(r);
      added++;
    });

    // Réconcilier rawData et cpData avec le serveur
    if (freshData.rawData) {
      g_rawData = freshData.rawData;
    }
    if (freshData.cpData && Array.isArray(freshData.cpData)) {
      g_cpData = freshData.cpData;
    }

    // Nettoyer les CP obsolètes
    cleanupCpData();

    // Sauvegarder en local
    idbSave({ master: g_master, rawData: g_rawData, cpData: g_cpData });

    // Rafraîchir l'interface
    renderMaster(); renderKanban(); renderCP(); updateStats();
    populateOperatorFilter();

    const changes = [updated && `${updated} mis à jour`, added && `${added} ajouté(s)`, removed && `${removed} retiré(s)`].filter(Boolean).join(', ');
    if (!silent) {
      toast(changes ? `Sync : ${changes}.` : 'Sync : données à jour.');
    } else if (changes) {
      toast(`Sync auto : ${changes}.`);
    }
  } catch(e) {
    console.warn('refreshData error:', e);
    if (!silent) toast('Erreur lors du rafraîchissement.');
  }
  _refreshing = false;
}

// Auto-sync toutes les 2 minutes
setInterval(() => refreshData(true), 120000);

// Chargement automatique au démarrage — IndexedDB prioritaire, API fallback
window.onload = async () => {
  // 1. Ouvrir IndexedDB
  try { await idbOpen(); } catch(e) { console.warn('IndexedDB indisponible, fallback API seul.'); }

  let loaded = false;

  // 2. Charger contacts depuis leur store dédié (indépendant)
  try {
    const idbContacts = await idbLoadContacts();
    if (idbContacts && idbContacts.length > 0) {
      g_contacts = idbContacts;
      console.log(`Contacts chargés depuis IndexedDB : ${g_contacts.length}`);
    }
  } catch(e) { console.warn('IDB contacts load:', e); }

  // 3. Essayer IndexedDB d'abord pour les dossiers (instantané, pas de réseau)
  try {
    const idbData = await idbLoad();
    if (idbData && idbData.master && idbData.master.length > 0) {
      g_master = idbData.master || [];
      g_rawData = idbData.rawData || {};
      g_cpData = idbData.cpData || [];
      if (!g_contacts.length && idbData.contacts) g_contacts = idbData.contacts;
      loaded = true;
      console.log('Chargé depuis IndexedDB (local).');
    }
  } catch(e) { console.warn('IDB load error:', e); }

  // 4. Toujours essayer l'API AutoIt pour réconcilier (données fraîches)
  // Si IDB a chargé, l'API sert de mise à jour ; sinon, elle est la source principale
  try {
    const respData = await fetch(`${API_URL}/load-data`);
    const d = await respData.json();
    if (d && d.master && d.master.length > 0) {
      if (loaded) {
        // Réconcilier : l'API est la source de vérité, IDB est le cache
        // Garder seulement les dossiers qui existent aussi côté serveur
        // OU qui ont été modifiés localement plus récemment
        const serverFiles = new Set(d.master.map(r => r.file));
        const serverMap = new Map(d.master.map(r => [r.file, r]));

        // Retirer de g_master les dossiers absents du serveur (sauf si modifiés localement après dernier save)
        g_master = g_master.filter(r => serverFiles.has(r.file));

        // Mettre à jour avec les données serveur si plus récentes
        g_master.forEach(r => {
          const srv = serverMap.get(r.file);
          if (srv && (srv._ts || 0) > (r._ts || 0)) {
            Object.keys(srv).forEach(k => { r[k] = srv[k]; });
          }
          serverMap.delete(r.file);
        });

        // Ajouter les dossiers présents sur le serveur mais pas en local
        serverMap.forEach(r => g_master.push(r));

        // Réconcilier rawData et cpData
        if (d.rawData) Object.assign(g_rawData, d.rawData);
        if (d.cpData) g_cpData = d.cpData;
      } else {
        g_master = d.master || [];
        g_rawData = d.rawData || {};
        g_cpData = d.cpData || [];
      }
      loaded = true;
      // Syncer vers IndexedDB avec les données réconciliées
      idbSave({ master: g_master, rawData: g_rawData, cpData: g_cpData });
    }
  } catch(e) { console.log('load-data:', e); }

  // Charger contacts séparément (multi-chunks)
  if (!g_contacts.length) {
    try {
      const respC = await fetch(`${API_URL}/load-contacts`);
      const c = await respC.json();
      if (Array.isArray(c) && c.length > 0) g_contacts = c;
    } catch(e) { console.log('load-contacts:', e); }
  }

  // Fallback ultime : ancien format monolithique (seulement si rien n'a chargé)
  if (!loaded) {
    try {
      const response = await fetch(`${API_URL}/load`);
      const d = await response.json();
      if (d && d.master && d.master.length > 0) {
        g_master = d.master || [];
        g_rawData = d.rawData || {};
        g_cpData = d.cpData || [];
        if (!g_contacts.length && d.contacts) g_contacts = d.contacts;
        loaded = true;
      }
    } catch(e) { console.log("Serveur hors ligne ou premier démarrage."); }

    if (loaded) {
      idbSave({ master: g_master, rawData: g_rawData, cpData: g_cpData });
      idbSaveContacts();
    }
  }

  if (loaded) {
    // Assigner l'opérateur courant aux dossiers orphelins (sans opérateur)
    const myName = localStorage.getItem('dispatch_operator') || '';
    if (myName) g_master.forEach(r => { if (!r.operator) r.operator = myName; });

    // Sync : s'assurer que tous les dossiers CP en statut 2 (CC) sont dans g_cpData
    g_master.forEach(r => {
      if (isCP(r.client) && r.cc === 'Cc' && ['2'].includes(String(r.statut))) {
        (r.file||'').split(/\s*\+\s*/).map(s=>s.trim()).filter(Boolean).forEach(sf => {
          addOrUpdateCP(r.client, sf, r.poids, r.vol, r.taxable, r.operator);
        });
      }
    });
    // Nettoyer les CP dont les dossiers ne sont plus en statut 2
    cleanupCpData();
    renderMaster(); renderKanban(); renderCP(); renderContacts(); updateStats();
    populateOperatorFilter();
    toast(loaded ? '✓ Données restaurées.' : 'Premier démarrage.');
  }
};
// ╔══════════════════════════════════════════════════════════════════════════╗

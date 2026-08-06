// ║  2. ACTION BATCH POUR LES CHANNEL PARTNERS (Onglet CP)                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝
document.getElementById('btn-rdv-cp').addEventListener('click', async () => {
  // Récupérer les indices cochés dans le tableau CP (basés sur g_cpData visible)
  const checkedCps = g_cpData
    .map((cp,i) => ({cp,i}))
    .filter(({cp}) => !g_operatorFilter || (cp.operator||'') === g_operatorFilter)
    .filter(({cp,i}) => {
      const cb = document.getElementById('chk-cp-' + i);
      return cb && cb.checked;
    });

  if (!checkedCps.length) return toast('Cochez au moins un CP.');
  if (!confirm(`Envoyer demandes RDV pour ${checkedCps.length} partenaire(s) ?`)) return;

  // Construire les données pour AutoIt depuis g_cpData + g_cpConfig
  const dataArr = checkedCps.map(({cp}) => {
    // Matching partiel : "Arrow" trouve "Arrow ECS SAS"
    const cl = (cp.client||'').toLowerCase();
    const cfg = g_cpConfig.find(c => cl.includes((c.nom||'').toLowerCase()) || (c.nom||'').toLowerCase().includes(cl)) || {};
    const cmds = cp.files.map(f => f.replace(/[^\x20-\x7E]/g, '').trim()).join(' + ');
    const doc  = (cp.doc||'').replace(/[^\x20-\x7E]/g, '').trim();
    // Poids = somme brute (pas taxable)
    return `${cp.client};${cmds};${cp.palettes||''};${cp.colis||''};${fmt(cp.poids)};${doc};${cfg.emailTo||''};${cfg.emailCc||''}`;
  });

  const strData = dataArr.join('|');
  toast(`Préparation de ${checkedCps.length} e-mail(s) CP en cours...`);
  const cpResult = await apiCall('action', { action: 'BATCH_CP', data: strData });
  if (!cpResult || cpResult.error) {
    toast('Erreur lors de l\'envoi des mails CP. Les dossiers restent à leur statut actuel.');
    return;
  }

  // Passer les dossiers CP au statut 3 + nettoyer g_cpData
  const idxToRemove = checkedCps.map(({i}) => i).sort((a,b) => b-a);
  checkedCps.forEach(({cp}) => {
    cp.files.forEach(sf => {
      const mi = g_master.findIndex(r => r.file === sf);
      if (mi >= 0) g_master[mi].statut = '3';
    });
  });
  idxToRemove.forEach(i => g_cpData.splice(i, 1));
  renderCP(); renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(`${checkedCps.length} CP traité(s) — dossiers passés au statut 3.`);
});

// ╔══════════════════════════════════════════════════════════════════════════╗

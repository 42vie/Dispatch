// ║  CC PROCESSING                                                           ║
// ╚══════════════════════════════════════════════════════════════════════════╝
function processCC(idx) {
  pushUndo('CC ' + (g_master[idx]?.file||''));
  const r = g_master[idx];
  r.cc = 'Cc';

  // Cas 1 : c'est un groupe (file contient ' + ') → éclater en solos CC
  if ((r.file||'').includes(' + ')) {
    const subs = r.file.split(' + ').map(s=>s.trim()).filter(Boolean);
    g_master.splice(idx, 1);
    subs.forEach(sf => {
      const raw = g_rawData[sf]||{};
      const p=raw.poids||r.poids/subs.length, v=raw.vol||r.vol/subs.length;
      const sv=raw.svct||r.svct, d=raw.dept||r.dept, t=calcTaxable(p,v);
      const cp=isCP(r.client), tr=calcTransp(sv,t,d,cp,p);
      const st=transpTarget(sv,r.client);
      const dup=g_master.find(x=>x.file===sf);
      if (dup) {
        Object.assign(dup,{client:r.client,rdl:r.rdl,svct:sv,transp:tr,
          poids:p,vol:v,taxable:t,dept:d,cc:'Cc',statut:st,operator:r.operator||''});
      } else {
        g_master.push({file:sf,client:r.client,rdl:r.rdl,svct:sv,transp:tr,
          poids:p,vol:v,taxable:t,dept:d,contact:r.contact||'',tel:r.tel||'',
          email:r.email||'',cc:'Cc',comment:'[Solo depuis groupe]',statut:st,
          operator:r.operator||'',
          _dateCreated:r._dateCreated||new Date().toISOString().slice(0,10)});
      }
      storeRaw(sf,r.client,sv,p,v,t,d,r.rdl);
      autoFillContact(sf,r.client);
      if (cp) addOrUpdateCP(r.client,sf,p,v,t,r.operator);
    });
    renderMaster(); renderCP(); updateStats();
    toast('CC validée — groupe éclaté en solos.');
    return;
  }

  // Cas 2 : solo → chercher s'il existe déjà un groupe CC du même client dans g_master
  const existing = g_master.find((x, i) => i !== idx
    && x.client === r.client
    && x.cc === 'Cc'
    && !x.file.includes(' + ')  // on cherche aussi les solos CC pour les fusionner
    || (i !== idx && x.client === r.client && x.cc === 'Cc' && x.file.includes(' + '))
  );

  // Cherche plus proprement : un groupe CC existant (peut être solo ou multi)
  const ccGroup = g_master.find((x, i) => i !== idx
    && x.client === r.client
    && x.cc === 'Cc'
    && parseInt(statutNum(x.statut)) < 6  // pas déjà traité
  );

  if (ccGroup) {
    // Rejoindre le groupe existant : fusionner les fichiers
    const allFiles = [
      ...ccGroup.file.split(' + ').map(s=>s.trim()).filter(Boolean),
      ...r.file.split(' + ').map(s=>s.trim()).filter(Boolean)
    ];
    ccGroup.file    = allFiles.join(' + ');
    ccGroup.poids   = +(ccGroup.poids + r.poids).toFixed(2);
    ccGroup.vol     = +(ccGroup.vol   + r.vol  ).toFixed(3);
    ccGroup.taxable = +calcTaxable(ccGroup.poids, ccGroup.vol).toFixed(2);
    ccGroup.svct    = mergeSVCT(ccGroup.svct, r.svct);
    ccGroup.transp  = calcTransp(ccGroup.svct, ccGroup.taxable, ccGroup.dept, isCP(ccGroup.client), ccGroup.poids);
    // Supprimer le solo qui vient de rejoindre le groupe
    g_master.splice(g_master.indexOf(r), 1);
    toast(`Fichier ${r.file} rejoint le groupe CC de ${ccGroup.client} (${allFiles.length} dossiers)`);
  } else {
    // Pas de groupe existant → solo CC, déterminer statut
    r.statut = transpTarget(r.svct, r.client);
    if (isCP(r.client)) addOrUpdateCP(r.client, r.file, r.poids, r.vol, r.taxable, r.operator);
    toast('CC validée.');
  }

  renderMaster(); renderCP(); updateStats();
  toast('CC validée — dossiers mis à jour.');
}

// ╔══════════════════════════════════════════════════════════════════════════╗

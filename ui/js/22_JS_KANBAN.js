// ║  KANBAN                                                                  ║
// ╚══════════════════════════════════════════════════════════════════════════╝
const K_CFG = [
  {n:1, title:'Pas encore CC',   c:'--k1-c', bg:'--k1-bg'},
  {n:2, title:'RDV à envoyer',   c:'--k2-c', bg:'--k2-bg', action:'ENVOYER →3'},
  {n:3, title:'Attente réponse', c:'--k3-c', bg:'--k3-bg', action:'VALIDER →5'},
  {n:4, title:'Pré-alerte S5',   c:'--k4-c', bg:'--k4-bg', action:'ALERTER →5'},
  {n:5, title:'Prêt pour FC',    c:'--k5-c', bg:'--k5-bg', action:'FC →6'},
  {n:6, title:'COMAT',           c:'--k6-c', bg:'--k6-bg', action:'COMAT →7'},
  {n:7, title:'NTO',             c:'--k7-c', bg:'--k7-bg', action:'NTO →8'},
  {n:9, title:'En attente',       c:'--k9-c', bg:'--k9-bg'},
];

function buildKanban() {
  const wrap = document.getElementById('kanban-wrap');
  wrap.innerHTML = K_CFG.map(k => `
    <div class="k-col" id="k-col-${k.n}"
      ondragover="event.preventDefault();document.getElementById('k-col-${k.n}').classList.add('drag-over')"
      ondragleave="document.getElementById('k-col-${k.n}').classList.remove('drag-over')"
      ondrop="onKDrop(event,${k.n})">
      <div class="k-header" style="border-top-color:var(${k.c})">
        <span class="k-title" style="color:var(${k.c})">${k.title}</span>
        <span class="k-count" style="background:var(${k.bg});color:var(${k.c})" id="k-cnt-${k.n}">0</span>
        ${k.action ? `<span class="k-action-btn" style="border-color:var(${k.c});color:var(${k.c});background:var(${k.bg})" onclick="kanbanAction(${k.n})">${k.action}</span>` : ''}
      </div>
      <div class="k-body" id="k-body-${k.n}"></div>
    </div>`).join('');
}

let g_kabSelected = new Set(); // fichiers sélectionnés

function renderKanban() {
  const kSort = document.getElementById('kanban-sort')?.value || '';
  const siblingsMap = buildSiblingsMap();
  K_CFG.forEach(k => {
    let items = g_master.filter(r => parseInt(statutNum(r.statut))===k.n && (!g_operatorFilter || (r.operator||'')===g_operatorFilter));
    for (const [col, allowed] of Object.entries(g_colFilters)) {
      if (!allowed || !allowed.size) continue;
      items = items.filter(r => {
        const v = col === 'statut' ? statutLabel(statutNum(r.statut)) : String(r[col]||'');
        return allowed.has(v);
      });
    }
    // Tri Kanban
    if (kSort) {
      items.sort((a, b) => {
        const va = (a[kSort]||'').toString().toLowerCase();
        const vb = (b[kSort]||'').toString().toLowerCase();
        return va < vb ? -1 : va > vb ? 1 : 0;
      });
    }
    document.getElementById(`k-cnt-${k.n}`).textContent = items.length;
    document.getElementById(`k-body-${k.n}`).innerHTML = items.map(r => {
      const idx = g_master.indexOf(r);
      const days = daysSince(r._dateCreated);
      const sel   = g_kabSelected.has(r.file);
      const fileE = r.file.replace(/'/g,"\\'");
      const isGrp = (r.file||'').includes(' + ');
      const subFiles = isGrp ? r.file.split(' + ') : null;
      const fileDisp = isGrp
        ? `<span style="font-size:10px;color:var(--purple);font-weight:700">${subFiles.length} dossiers</span>`
        : `<span style="font-family:var(--mono)">${trunc(r.file,14)}</span>`;
      // Colonne 5 : afficher date livraison + transporteur éditables
      const fcRow = (k.n === 5) ? (() => {
        fcEnsureAutoDate(r, false);
        const safeFile = (r.file||'').replace(/[^a-zA-Z0-9]/g,'_');
        return `<div class="k-fc-row" onclick="event.stopPropagation()">
          <span class="k-fc-lbl">📅</span>
          <input class="k-fc-inp" style="width:70px" value="${r.fcDate}"
            id="kfc-date-${safeFile}"
            onchange="kabFcUpdate('${fileE}','date',this.value)"
            onclick="event.stopPropagation()" onfocus="event.stopPropagation()">
          <span class="k-fc-lbl">🚚</span>
          <input class="k-fc-inp" style="width:90px" value="${trunc(r.transp||'?',12)}"
            id="kfc-transp-${safeFile}"
            onchange="kabFcUpdate('${fileE}','transp',this.value)"
            onclick="event.stopPropagation()" onfocus="event.stopPropagation()">
        </div>`;
      })() : '';
      return `<div class="k-item${sel?' k-selected':''}${isGrp?' k-group':''}" style="border-left-color:var(${k.c})"
        data-file="${(r.file||'').replace(/"/g,'&quot;')}" data-client="${(r.client||'').replace(/"/g,'&quot;')}" data-transp="${(r.transp||'').replace(/"/g,'&quot;')}" data-email="${(r.email||'').replace(/"/g,'&quot;')}" data-rdl="${(r.rdl||'').replace(/"/g,'&quot;')}" data-cp="${(r.cp||r.cc||'').replace(/"/g,'&quot;')}"
        draggable="true"
        ondragstart="kanbanDragStartBulk(event,'${fileE}')"
        ondragover="event.preventDefault();event.stopPropagation();this.classList.add('k-drop-target')"
        ondragleave="this.classList.remove('k-drop-target')"
        ondrop="event.stopPropagation();this.classList.remove('k-drop-target');onCardDrop(event,'${fileE}')"
        ondblclick="openEdit(${idx})">
        <input type="checkbox" class="k-chk" ${sel?'checked':''} onclick="event.stopPropagation();kabToggle(event,'${fileE}')">
        <div class="k-file">${fileDisp}${isGrp?`<span class="tag-group" style="font-size:9px;margin-left:4px">GRP</span>`:''}</div>
        <div class="k-client">${trunc(r.client,20)}${getSiblingBadge(r, siblingsMap)}</div>
        ${(k.n>=2&&k.n<=7||k.n===9) ? `<div class="k-meta">${daysBadge(days)}</div>` : ''}
        ${fcRow}
      </div>`;
    }).join('');
  });
  kabUpdateBar();
}

function kabToggle(e, file) {
  // Ignorer si double-clic (handled by ondblclick)
  if (e && e.detail >= 2) return;
  if (g_kabSelected.has(file)) g_kabSelected.delete(file);
  else g_kabSelected.add(file);
  renderKanban();
}

function kabUpdateBar() {
  const bar = document.getElementById('kanban-action-bar');
  const cnt = g_kabSelected.size;
  document.getElementById('kab-count').textContent = cnt;
  if (cnt > 0) bar.classList.add('visible');
  else bar.classList.remove('visible');
}

function kabClearSel() {
  g_kabSelected.clear();
  renderKanban();
}

function kabSelectAll() {
  // Sélectionner tous les éléments visibles dans le kanban
  const items = document.querySelectorAll('.k-item');
  const allVisible = [];
  items.forEach(el => {
    const mono = el.querySelector('.k-file');
    if (mono) allVisible.push(mono.textContent.trim());
  });
  // Trouver le fichier complet dans g_master (le trunc peut couper)
  allVisible.forEach(short => {
    const r = g_master.find(r => r.file.startsWith(short.replace('…','')) || r.file === short);
    if (r) g_kabSelected.add(r.file);
  });
  renderKanban();
}

function kabGetSelected() {
  return g_master.filter(r => g_kabSelected.has(r.file));
}

function kabChangeStatus() {
  const sel = kabGetSelected();
  if (!sel.length) return;
  document.getElementById('kab-status-files').textContent =
    sel.length + ' dossier(s) : ' + sel.map(r => r.file).join(', ');
  openModal('modal-kab-status');
}

function kabApplyStatus() {
  const newStat = document.getElementById('kab-new-status').value;
  kabGetSelected().forEach(r => { r.statut = newStat; stampRecord(r); });
  closeModal('modal-kab-status');
  kabClearSel();
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(`Statut ${newStat} appliqué.`);
}

function kabOpenEdit() {
  const sel = kabGetSelected();
  if (!sel.length) return;
  document.getElementById('kab-edit-title').textContent = 'Éditer ' + sel.length + ' dossier(s)';
  document.getElementById('kab-edit-files').textContent =
    sel.map(r => r.file).join(' · ');
  ['kab-contact','kab-tel','kab-email','kab-comment'].forEach(id =>
    document.getElementById(id).value = '');
  openModal('modal-kab-edit');
}

function kabSaveEdit() {
  const contact = document.getElementById('kab-contact').value.trim();
  const tel     = document.getElementById('kab-tel').value.trim();
  const email   = document.getElementById('kab-email').value.trim();
  const comment = document.getElementById('kab-comment').value.trim();
  kabGetSelected().forEach(r => {
    if (contact) r.contact = contact;
    if (tel)     r.tel     = tel;
    if (email)   r.email   = email;
    if (comment) r.comment = comment;
  });
  closeModal('modal-kab-edit');
  kabClearSel();
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast('Modifications appliquées à la sélection.');
}

async function kabSendMail() {
  const sel = kabGetSelected();
  if (!sel.length) return;
  const missing = sel.filter(r => !r.email);
  if (missing.length) {
    if (!confirm(`${missing.length} dossier(s) sans email (${missing.map(r=>r.file).join(', ')}).
Continuer pour les autres ?`)) return;
  }
  const withEmail = sel.filter(r => r.email);
  if (!withEmail.length) return toast('Aucun dossier avec email dans la sélection.');
  const strData = withEmail.map(r => {
    const contactStr = `${r.contact||''} ${r.tel||''}`.trim();
    return `${r.file};${r.email};${r.client};${r.transp||''};${contactStr}`;
  }).join('|');
  toast(`📧 Envoi mail RDV pour ${withEmail.length} dossier(s)...`);
  const mailResult = await apiCall('action', { action: 'KANBAN_2', data: strData });
  if (!mailResult || mailResult.error) {
    toast('Erreur lors de l\'envoi des mails. Les dossiers restent à leur statut actuel.');
    return;
  }
  withEmail.forEach(r => { if (parseInt(statutNum(r.statut)) < 3) r.statut = '3'; });
  kabClearSel();
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(`Mail RDV envoyé pour ${withEmail.length} dossier(s).`);
}

// ══ Actions batch depuis la barre Kanban ══

async function kabSendAlerte() {
  const sel = kabGetSelected();
  if (!sel.length) return toast('Sélectionnez au moins un dossier.');
  const withEmail = sel.filter(r => r.email);
  if (!withEmail.length) return toast('Aucun dossier avec email dans la sélection.');
  if (!confirm(`Envoyer une pré-alerte pour ${withEmail.length} dossier(s) ?`)) return;
  const strData = withEmail.map(r => {
    const contactStr = `${r.contact||''} ${r.tel||''}`.trim();
    return `${r.file};${r.email};${r.client};${r.transp||''};${contactStr}`;
  }).join('|');
  showSpinner(`Pré-alerte : ${withEmail.length} dossier(s)...`);
  try {
    await apiCall('action', { action: 'KANBAN_4', data: strData });
    withEmail.forEach(r => { if (parseInt(statutNum(r.statut)) < 5) r.statut = '5'; stampRecord(r); });
    kabClearSel();
    renderMaster(); renderKanban(); updateStats(); markDirty();
    toast(`Pré-alerte envoyée pour ${withEmail.length} dossier(s).`);
  } catch(e) { toast('Erreur pré-alerte : ' + (e.message||'')); }
  finally { hideSpinner(); }
}

function kabLaunchFC() {
  const sel = kabGetSelected();
  if (!sel.length) return toast('Sélectionnez au moins un dossier.');
  // Ouvrir la modale FC avec les dossiers sélectionnés
  fcOpenModal(sel);
  kabClearSel();
}

async function kabLaunchCOMAT() {
  const sel = kabGetSelected();
  if (!sel.length) return toast('Sélectionnez au moins un dossier.');
  if (!confirm(`Lancer COMAT pour ${sel.length} dossier(s) ?`)) return;
  // Déterminer le statut cible (7 ou 8 selon d'où ils viennent)
  const toStatus = 7;
  comatStartBatch(sel, toStatus);
  kabClearSel();
}

function onKDrop(event, status) {
  event.preventDefault();
  document.getElementById(`k-col-${status}`).classList.remove('drag-over');
  const files = (window._kanbanBulkDragFiles && window._kanbanBulkDragFiles.length) ? window._kanbanBulkDragFiles : (g_dragFile ? [g_dragFile] : []);
  if (!files.length) return;
  pushUndo('Déplacement groupé → statut ' + status);
  files.forEach(f => {
    const idx = g_master.findIndex(r => r.file===f);
    if (idx < 0) return;
    g_master[idx].statut = String(status);
    stampRecord(g_master[idx]);
    if (status>=2 && g_master[idx].cc!=='Cc') g_master[idx].cc='Cc';
  });
  cleanupCpData();
  window._kanbanBulkDragFiles = [];
  g_dragFile = null;
  g_kabSelected.clear();
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(files.length + ' dossier(s) déplacé(s).');
}

function onCardDrop(event, targetFile) {
  event.preventDefault();
  if (!g_dragFile || g_dragFile === targetFile) { g_dragFile = null; return; }
  const srcIdx = g_master.findIndex(r => r.file === g_dragFile);
  const tgtIdx = g_master.findIndex(r => r.file === targetFile);
  if (srcIdx < 0 || tgtIdx < 0) { g_dragFile = null; return; }
  pushUndo('Fusion ' + g_dragFile + ' + ' + targetFile);
  const src = g_master[srcIdx], tgt = g_master[tgtIdx];
  // Collecter tous les fichiers individuels
  const srcFiles = (src.file||'').split(' + ').map(s=>s.trim()).filter(Boolean);
  const tgtFiles = (tgt.file||'').split(' + ').map(s=>s.trim()).filter(Boolean);
  const allFiles = [...tgtFiles, ...srcFiles];
  // Vérifier pas de doublons
  const uniq = [...new Set(allFiles)];
  if (uniq.length < 2) { g_dragFile = null; return; }
  // Agréger poids, vol
  const newPoids = +(tgt.poids + src.poids).toFixed(2);
  const newVol   = +(tgt.vol + src.vol).toFixed(3);
  const newTax   = +calcTaxable(newPoids, newVol).toFixed(2);
  const newSvct  = mergeSVCT(tgt.svct, src.svct);
  const cpFlag   = isCP(tgt.client);
  const newTransp = calcTransp(newSvct, newTax, tgt.dept, cpFlag);
  // Créer la nouvelle ligne groupée
  const merged = {
    file: uniq.join(' + '), client: tgt.client, rdl: tgt.rdl, svct: newSvct,
    transp: newTransp, poids: newPoids, vol: newVol, taxable: newTax,
    dept: tgt.dept, contact: tgt.contact||'', tel: tgt.tel||'', email: tgt.email||'',
    cc: tgt.cc, comment: tgt.comment||'', statut: tgt.statut,
    operator: tgt.operator||'',
    _dateCreated: tgt._dateCreated || new Date().toISOString().slice(0,10)
  };
  // Supprimer les deux lignes originales (plus grand index en premier)
  const idxs = [srcIdx, tgtIdx].sort((a,b) => b-a);
  idxs.forEach(i => g_master.splice(i, 1));
  g_master.push(merged);
  g_dragFile = null;
  g_kabSelected.clear();
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(`Groupe créé : ${uniq.length} dossiers fusionnés.`);
}

async function kanbanAction(from) {
  const toMap = {2:3, 4:5, 5:6, 6:7, 7:8};
  const affected = g_master.filter(r => parseInt(statutNum(r.statut)) === from && (!g_operatorFilter || (r.operator||'') === g_operatorFilter));

  if (!affected.length) return toast('Aucun dossier dans cette colonne.');

  // Colonne 3 : les S5 vont en 4 (pré-alerte), les autres en 5 (FC prêt)
  let to;
  if (from === 3) {
    const hasS5 = affected.some(r => (r.svct||'').includes('S5'));
    const hasNonS5 = affected.some(r => !(r.svct||'').includes('S5'));
    if (hasS5 && hasNonS5) to = 'mixed';
    else if (hasS5) to = 4;
    else to = 5;
  } else {
    to = toMap[from];
  }

  if (!confirm(`Traiter ${affected.length} dossier(s) et avancer ?`)) return;

  const strData = affected.map(r => {
    const contactStr = `${r.contact || ''} ${r.tel || ''}`.trim();
    const cleanFile = (r.file||'').replace(/[^\x20-\x7E+]/g, '').trim();
    return `${cleanFile};${r.email};${r.client};${r.transp || ''};${contactStr}`;
  }).join('|');

  if (from === 7) {
    // NTO batch : calculer chaque dossier, afficher dans onglet NTO, PAS d'avance auto
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tabpane').forEach(p => p.classList.remove('active'));
    document.querySelector('[data-tab="nto"]').classList.add('active');
    document.getElementById('tab-nto').classList.add('active');
    ntoBatchLancer(affected); // stocke affected globalement pour valider après
    return; // statut 8 = bouton "Passer statut 8" dans le batch
  }

  if (from === 6) {
    // CAS COMAT : traitement séquentiel avec pause/reprise
    comatStartBatch(affected, to);
    return;
  } else if (from === 5) {
    // FC : ouvrir la modale de paramétrage dates/horaires
    fcOpenModal(affected);
    return; // fcLancer() gère la suite
  } else {
    toast(`Ordre envoyé à AutoIt pour ${affected.length} dossier(s)...`);
    const kanResult = await apiCall('action', { action: 'KANBAN_' + from, data: strData });
    if (!kanResult || kanResult.error) {
      toast('Erreur lors de l\'envoi. Les dossiers restent à leur statut actuel.');
      return;
    }
    if (to === 'mixed' || from === 3) {
      // Colonne 3 : S5 → 4 (pré-alerte), non-S5 → 5 (FC prêt)
      affected.forEach(r => {
        r.statut = (r.svct||'').includes('S5') ? '4' : '5';
      });
    } else {
      affected.forEach(r => { r.statut = String(to); });
    }
  }

  cleanupCpData();
  renderMaster(); renderKanban(); updateStats();
  markDirty();
}
//╔══════════════════════════════════════════════════════════════════════════╗

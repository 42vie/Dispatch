// ║  OPÉRATEUR / FILTRE PERSONNE                                             ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function populateOperatorFilter() {
  const sel = document.getElementById('operator-filter');
  if (!sel) return;
  const names = [...new Set(g_master.map(r => r.operator || '').filter(Boolean))].sort();
  const saved = localStorage.getItem('dispatch_operator') || '';
  sel.innerHTML = '<option value="">Tous</option>'
    + names.map(n => '<option value="' + n + '"' + (n===saved?' selected':'') + '>' + n + '</option>').join('');
  // Toujours restaurer le filtre depuis localStorage — même si le nom n'est pas dans g_master
  if (saved) {
    sel.value = saved;
    g_operatorFilter = saved;
  }
}

function applyOperatorFilter() {
  const sel = document.getElementById('operator-filter');
  g_operatorFilter = sel ? sel.value : '';
  localStorage.setItem('dispatch_operator', g_operatorFilter);
  const nameInput = document.getElementById('opts-my-name');
  if (nameInput) nameInput.value = g_operatorFilter;
  renderMaster(); renderKanban(); updateStats();
}

function optsUpdateMyName() {
  const v = (document.getElementById('opts-my-name').value || '').trim();
  g_operatorFilter = v;
  localStorage.setItem('dispatch_operator', v);
  const sel = document.getElementById('operator-filter');
  if (sel) sel.value = v;
  renderMaster(); renderKanban(); updateStats();
}

// Démarrage : afficher popup si pas de nom mémorisé
function showIdentityModal() {
  const saved = localStorage.getItem('dispatch_operator') || '';
  const names = [...new Set(g_master.map(r => r.operator || '').filter(Boolean))].sort();
  if (!names.length) return; // pas de noms dans le CSV
  const sel = document.getElementById('identity-select');
  if (sel) {
    sel.innerHTML = '<option value="">— Voir tous les dossiers —</option>'
      + names.map(n => '<option value="' + n + '"' + (n===saved?' selected':'') + '>' + n + '</option>').join('');
  }
  const inp = document.getElementById('identity-input');
  if (inp && saved) inp.value = saved;
  // Afficher seulement si aucun nom encore choisi
  if (!saved) {
    document.getElementById('modal-identity').style.display = 'flex';
  } else {
    g_operatorFilter = saved;
    const filterSel = document.getElementById('operator-filter');
    if (filterSel) filterSel.value = saved;
  }
}

function identityConfirm() {
  const sel = document.getElementById('identity-select').value;
  const inp = (document.getElementById('identity-input').value || '').trim();
  const name = inp || sel || '';
  localStorage.setItem('dispatch_operator', name);
  g_operatorFilter = name;
  const filterSel = document.getElementById('operator-filter');
  if (filterSel) filterSel.value = name;
  const nameOpts = document.getElementById('opts-my-name');
  if (nameOpts) nameOpts.value = name;
  document.getElementById('modal-identity').style.display = 'none';
  renderMaster(); renderKanban(); updateStats();
}

function identitySkip() {
  localStorage.setItem('dispatch_operator', '');
  g_operatorFilter = '';
  document.getElementById('modal-identity').style.display = 'none';
}

// ╔══════════════════════════════════════════════════════════════════════════╗

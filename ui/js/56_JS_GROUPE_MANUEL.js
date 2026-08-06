// ║  GROUPE MANUEL                                                           ║
// ╚══════════════════════════════════════════════════════════════════════════╝
function openGroupManual() {
  // Pré-remplir avec les dossiers sélectionnés dans le dispatch si des lignes sont cochées
  const checked = [...document.querySelectorAll('#master-tbody tr')]
    .filter(tr => tr.querySelector('input[type=checkbox]')?.checked)
    .map(tr => {
      const fileCell = tr.querySelector('td:nth-child(2)');
      return fileCell ? fileCell.title || fileCell.textContent.trim() : '';
    }).filter(Boolean);

  document.getElementById('grpm-files').value = checked.join('\n');
  document.getElementById('grpm-preview').style.display = 'none';

  // Si un seul client dans la sélection, pré-remplir
  if (checked.length) {
    const rows = g_master.filter(r => checked.includes(r.file));
    if (rows.length) {
      document.getElementById('grpm-client').value = rows[0].client || '';
      document.getElementById('grpm-rdl').value    = rows[0].rdl    || '';
      document.getElementById('grpm-svct').value   = rows[0].svct   || '';
      document.getElementById('grpm-dept').value   = rows[0].dept   || '';
      document.getElementById('grpm-transp').value = rows[0].transp || '';
    }
  } else {
    ['grpm-client','grpm-rdl','grpm-svct','grpm-dept','grpm-transp'].forEach(id =>
      document.getElementById(id).value = '');
  }
  openModal('modal-group-manual');
}

function grpmParseFiles() {
  const raw = document.getElementById('grpm-files').value;
  return raw.split(/[\n\r,;\t ]+/).map(s => s.trim()).filter(Boolean);
}

function grpmPreview() {
  const files = grpmParseFiles();
  if (!files.length) return toast('Saisir au moins un N° de dossier.');
  const prev = document.getElementById('grpm-preview');
  let html = '<div style="font-weight:600;margin-bottom:6px;color:var(--text)">'
    + files.length + ' dossier(s) dans le groupe :</div>';
  let totP = 0, totV = 0;
  files.forEach(f => {
    const r = g_master.find(x => x.file === f);
    if (r) {
      totP += r.poids || 0; totV += r.vol || 0;
      html += '<div style="display:flex;gap:12px;padding:3px 0;border-bottom:1px solid var(--border)">'
        + '<span style="font-family:var(--mono);color:var(--blue);min-width:130px">' + f + '</span>'
        + '<span style="color:var(--text2)">' + (r.client||'—') + '</span>'
        + '<span style="color:var(--orange);font-family:var(--mono);margin-left:auto">' + fmt(r.taxable) + ' kg tx</span>'
        + '</div>';
    } else {
      html += '<div style="display:flex;gap:12px;padding:3px 0;border-bottom:1px solid var(--border)">'
        + '<span style="font-family:var(--mono);color:var(--text3);min-width:130px">' + f + '</span>'
        + '<span style="color:var(--text3);font-style:italic">→ sera créé vide</span>'
        + '</div>';
    }
  });
  const tax = calcTaxable(totP, totV);
  html += '<div style="margin-top:8px;font-size:12px;font-weight:600;color:var(--purple)">'
    + 'Total groupe : ' + fmt(totP) + ' kg brut · ' + fmt(totV,3) + ' m³ · '
    + '<span style="color:var(--orange)">' + fmt(tax) + ' kg taxable</span></div>';
  prev.innerHTML = html;
  prev.style.display = 'block';
  // Auto-remplir poids/transp si vides
  const svct   = document.getElementById('grpm-svct').value;
  const dept   = document.getElementById('grpm-dept').value;
  const client = document.getElementById('grpm-client').value;
  if (!document.getElementById('grpm-transp').value && dept) {
    document.getElementById('grpm-transp').value = calcTransp(svct, tax, dept, isCP(client));
  }
}

function grpmAutoTransp() {
  const files  = grpmParseFiles();
  const svct   = document.getElementById('grpm-svct').value;
  const dept   = document.getElementById('grpm-dept').value;
  const client = document.getElementById('grpm-client').value;
  let totP = 0, totV = 0;
  files.forEach(f => {
    const r = g_rawData[f] || g_master.find(x => x.file === f) || {};
    totP += r.poids || 0; totV += r.vol || 0;
  });
  const tax = calcTaxable(totP, totV);
  document.getElementById('grpm-transp').value = calcTransp(svct, tax, dept, isCP(client));
}

function grpmSave() {
  const files  = grpmParseFiles();
  if (files.length < 2) return toast('Un groupe nécessite au moins 2 dossiers.');
  const client = document.getElementById('grpm-client').value.trim();
  const rdl    = document.getElementById('grpm-rdl').value.trim();
  const svct   = document.getElementById('grpm-svct').value.trim();
  const dept   = document.getElementById('grpm-dept').value.trim();
  const statut = document.getElementById('grpm-statut').value;
  const today  = new Date().toISOString().slice(0,10);

  // Stocker les raws et supprimer les existants
  let totP = 0, totV = 0;
  files.forEach(f => {
    const existing = g_master.find(r => r.file === f);
    if (existing) {
      totP += existing.poids || 0;
      totV += existing.vol   || 0;
      storeRaw(f, existing.client||client, existing.svct||svct,
        existing.poids||0, existing.vol||0, existing.taxable||0,
        existing.dept||dept, existing.rdl||rdl);
      // Supprimer le dossier solo
      const idx = g_master.indexOf(existing);
      if (idx >= 0) g_master.splice(idx, 1);
    }
  });

  const tax    = calcTaxable(totP, totV);
  let   transp = document.getElementById('grpm-transp').value.trim()
              || calcTransp(svct, tax, dept, isCP(client));
  const fileStr = files.join(' + ');

  // Vérifier que le groupe n'existe pas déjà
  if (g_master.find(r => r.file === fileStr)) return toast('Ce groupe existe déjà.');

  g_master.push({
    file: fileStr, client, rdl, svct, transp,
    poids: +totP.toFixed(2), vol: +totV.toFixed(3), taxable: +tax.toFixed(2),
    dept, contact: '', tel: '', email: '', cc: 'Non',
    comment: '[Groupe Manuel]', statut, _dateCreated: today
  });

  closeModal('modal-group-manual');
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(`Groupe créé : ${files.length} dossiers → ${fileStr}`);
}


// ╔══════════════════════════════════════════════════════════════════════════╗

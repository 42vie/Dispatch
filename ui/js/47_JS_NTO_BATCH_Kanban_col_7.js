// ║  NTO BATCH — Kanban col 7                                                ║
// ╚══════════════════════════════════════════════════════════════════════════╝
let _ntoBatchAffected = []; // dossiers en attente de validation statut 8

function ntoBatchLancer(affected) {
  _ntoBatchAffected = affected;

  const body  = document.getElementById('nto-batch-body');
  const IDF   = new Set(['60','75','77','78','91','92','93','94','95']);
  const EFDS  = s => /efds|groussard/i.test(s||'');
  const isUPS = s => /ups/i.test(s||'');

  let rows = '', totalCalc = 0, skipped = 0, counted = 0;
  let lines = []; // pour copier

  affected.forEach(r => {
    // Extraire les dossiers individuels si groupe
    const files = r.file.includes(' + ')
      ? r.file.split(' + ').map(s => s.trim())
      : [r.file.trim()];

    files.forEach((f, fIdx) => {
      const transp  = (r.transp || '').trim();
      const dept    = String(r.dept || '').replace(/^0/,'');
      const taxable = parseFloat(r.taxable) || 0;
      const isGroup = files.length > 1;
      const rawF    = g_rawData[f] || {};
      const suppl   = +(rawF.supplement || 0);

      // EFDS / Groussard → skip
      if (EFDS(transp)) {
        skipped++;
        rows += _ntoBatchRow(f, transp, '<span style="color:var(--text3);font-style:italic">Sur cotation — ignoré</span>', '—', '#FAFBFC');
        return;
      }

      // Groupe : seul le 1er fichier porte la NTO, le reste = frais camion individuels
      if (isGroup && fIdx > 0) {
        const supplStr = suppl ? suppl.toFixed(2) : '0';
        const supplDetail = suppl
          ? 'Groupe — frais camion import ' + suppl.toFixed(2) + ' €'
          : 'Groupe — inclus dans ' + files[0];
        totalCalc += suppl;
        if (suppl) counted++;
        rows += _ntoBatchRow(f, 'NTO', supplDetail, supplStr, '#F8F9FA');
        lines.push({ file: f, prix: supplStr });
        return;
      }

      let prix, detail, color = '';

      if (isUPS(transp)) {
        // UPS = forfait 25 fixe
        prix   = 25;
        detail = 'UPS — forfait fixe';
        color  = 'var(--green-bg)';
      } else if (IDF.has(dept)) {
        // Cartage IDF
        const res = ntoCalcUTE(dept, taxable);
        if (!res) {
          skipped++;
          rows += _ntoBatchRow(f, transp, '<span style="color:var(--red)">Dept IDF inconnu : ' + (dept||'').replace(/[<>&"]/g, c=>({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'})[c]) + '</span>', '—', 'var(--red-bg)');
          return;
        }
        prix   = Math.round(res.total * 100) / 100;
        detail = 'Cartage Dept ' + dept + ' · ' + taxable + ' kg · fuel ' + (res.fuelRate*100).toFixed(1) + '% · min=' + res.minimum.toFixed(2);
        color  = 'var(--blue-bg)';
      } else {
        // DGS national
        const res = ntoCalcDGS(dept, taxable);
        if (!res) {
          skipped++;
          rows += _ntoBatchRow(f, transp, '<span style="color:var(--red)">Dept inconnu : ' + (dept||'').replace(/[<>&"]/g, c=>({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'})[c]) + '</span>', '—', 'var(--red-bg)');
          return;
        }
        prix   = Math.round(res.total * 100) / 100;
        detail = 'DGS Dept ' + dept + ' (' + res.nom + ') · ' + taxable + ' kg · base ' + res.base.toFixed(2) + ' + ' + NTO_DGS_FRAIS + '€ frais · fuel ' + (res.fuelRate*100).toFixed(1) + '%';
        color  = 'var(--purple-bg)';
      }

      // Ajouter le supplément camion import au prix NTO
      if (suppl) {
        detail += ' + camion ' + suppl.toFixed(2) + ' €';
        prix += suppl;
        prix = Math.round(prix * 100) / 100;
      }

      totalCalc += prix;
      counted++;
      const prixStr = isUPS(transp) ? (25 + suppl).toFixed(2) : prix.toFixed(2);
      rows += _ntoBatchRow(f, 'NTO', detail, prixStr, color);
      lines.push({ file: f, prix: prixStr });
    });
  });

  body.innerHTML = rows;
  document.getElementById('nto-batch-title').textContent = 'NTO Batch — ' + affected.length + ' dossier(s) colonne 7';
  document.getElementById('nto-batch-count').textContent = counted;
  document.getElementById('nto-batch-skip').textContent  = skipped ? skipped + ' ignoré(s) EFDS' : '';
  document.getElementById('nto-batch-total').textContent = counted ? totalCalc.toFixed(2) + ' €' : '—';

  // Stocker les lignes pour copie
  window._ntoBatchLines = lines;

  // Afficher
  document.getElementById('nto-empty').style.display  = 'none';
  document.getElementById('nto-result').style.display = 'none';
  document.getElementById('nto-batch').style.display  = 'flex';
}

function _ntoBatchRow(file, event, detail, prix, bg) {
  return '<div style="display:grid;grid-template-columns:150px 55px 1fr 90px;border-bottom:1px solid var(--border);background:' + (bg||'') + '">'
    + '<div style="padding:6px 8px;font-family:var(--mono);font-size:11px;font-weight:500;border-right:1px solid var(--border)">' + file + '</div>'
    + '<div style="padding:6px 8px;font-size:11px;border-right:1px solid var(--border);color:var(--k7-c);font-weight:600">' + event + '</div>'
    + '<div style="padding:6px 8px;font-size:11px;color:var(--text2);border-right:1px solid var(--border)">' + detail + '</div>'
    + '<div style="padding:6px 8px;font-family:var(--mono);font-size:12px;font-weight:700;color:var(--purple);text-align:right">' + prix + '</div>'
    + '</div>';
}

// Copier format TSV : File TAB NTO TAB TAB TAB TAB TAB TAB Prix
// Colonnes : B=File, C=Event, D=Date, E=Time, F=From, G=To, H=Voy/FLT, I=Remark
function ntoBatchCopier() {
  const lines = window._ntoBatchLines || [];
  if (!lines.length) return toast('Aucune ligne à copier.');
  const tsv = lines.map(l => l.file + '\tNTO\t\t\t\t\t\t' + l.prix).join('\n');
  navigator.clipboard.writeText(tsv)
    .then(() => toast('✓ ' + lines.length + ' ligne(s) copiées — colle en B dans ton outil.'))
    .catch(() => {
      const t = document.createElement('textarea');
      t.value = tsv;
      document.body.appendChild(t); t.select(); document.execCommand('copy'); document.body.removeChild(t);
      toast('✓ ' + lines.length + ' ligne(s) copiées — colle en B dans ton outil.');
    });
}

// Passer les dossiers au statut 8
function ntoBatchValider() {
  if (!_ntoBatchAffected.length) return;
  if (!confirm('Passer ' + _ntoBatchAffected.length + ' dossier(s) au statut 8 (Terminé) ?')) return;
  _ntoBatchAffected.forEach(r => { r.statut = '8'; });
  renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(_ntoBatchAffected.length + ' dossier(s) passés au statut 8.');
  ntoBatchFermer();
}

function ntoBatchFermer() {
  document.getElementById('nto-batch').style.display = 'none';
  document.getElementById('nto-empty').style.display = 'flex';
  _ntoBatchAffected = [];
  window._ntoBatchLines = [];
}


// ╔══════════════════════════════════════════════════════════════════════════╗

// ║  EXPORT EXCEL — Tableau Dispatch complet                                ║
// ╚══════════════════════════════════════════════════════════════════════════╝
document.getElementById('btn-export-excel').onclick = () => {
  if (!g_master.length) return toast('Aucun dossier à exporter.');
  const headers = ['N° Dossier','Client','RDL','Service','Transporteur','Poids (kg)','Volume (m³)','Taxable','Dept','Contact','Téléphone','Email','CC','Statut','Opérateur','Commentaire','Date création'];
  const rows = g_master.map(r => [
    r.file || '',
    r.client || '',
    r.rdl || '',
    r.svct || '',
    r.transp || '',
    r.poids || '',
    r.vol || '',
    r.taxable || '',
    r.dept || '',
    r.contact || '',
    r.tel || '',
    r.email || '',
    r.cc || '',
    STATUT_LABELS[String(r.statut||'')] || r.statut || '',
    r.operator || '',
    r.comment || '',
    r._dateCreated || ''
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  // Auto-largeur colonnes
  ws['!cols'] = headers.map((h,i) => ({ wch: Math.max(h.length, ...rows.map(r => String(r[i]).length).slice(0,50)) + 2 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Dispatch');
  XLSX.writeFile(wb, 'Dispatch_Export_' + new Date().toISOString().slice(0,10) + '.xlsx');
  toast('Export Excel téléchargé.');
};

document.getElementById('btn-load-json').onclick = () => document.getElementById('json-input').click();
document.getElementById('json-input').addEventListener('change', e => {
  const f=e.target.files[0]; if (!f) return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try {
      const d=JSON.parse(ev.target.result);
      g_master  = d.master   || [];
      g_rawData = d.rawData  || {};
      g_cpData  = d.cpData   || [];
      g_contacts= d.contacts || [];
      renderMaster(); renderKanban(); renderCP(); renderContacts(); updateStats();
      if (d.contacts && typeof saveContacts === 'function') saveContacts();
      toast('Sauvegarde chargée — tous vos dossiers et contacts sont restaurés.');
    } catch(err){ alert('Erreur lecture JSON : '+err.message); }
  };
  reader.readAsText(f);
  e.target.value='';
});

function downloadText(filename, text) {
  const a=document.createElement('a');
  a.href='data:text/plain;charset=utf-8,'+encodeURIComponent(text);
  a.download=filename; a.click();
}

// ╔══════════════════════════════════════════════════════════════════════════╗

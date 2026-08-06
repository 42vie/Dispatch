// ║  IMPORT EXCEL                                                            ║
// ╚══════════════════════════════════════════════════════════════════════════╝
document.getElementById('btn-import').onclick = () => document.getElementById('file-input').click();

// Retirer fichiers cochés du tableau CP vers g_master solo
document.getElementById('btn-retirer-cp').onclick = () => {
  const checked = g_cpData
    .map((cp,i) => ({cp,i}))
    .filter(({cp}) => !g_operatorFilter || (cp.operator||'') === g_operatorFilter)
    .filter(({cp,i}) => { const cb = document.getElementById('chk-cp-'+i); return cb && cb.checked; });

  if (!checked.length) return toast('Cochez au moins un groupe CP à retirer.');
  if (!confirm(`Retirer ${checked.length} groupe(s) CP et remettre les dossiers en Solo ?`)) return;

  const today = new Date().toISOString().slice(0,10);
  checked.forEach(({cp}) => {
    cp.files.forEach(sf => {
      if (g_master.find(r => r.file === sf)) return; // déjà présent
      const raw = g_rawData[sf] || {};
      g_master.push({
        file: sf, client: cp.client, rdl: '', svct: '',
        transp: calcTransp('', raw.taxable||0, raw.dept||'', false),
        poids: raw.poids||0, vol: raw.vol||0, taxable: raw.taxable||0,
        dept: raw.dept||'', contact:'', tel:'', email:'',
        cc:'Non', comment:'[Retiré CP]', statut:'1',
        operator: cp.operator||'', _dateCreated: today
      });
    });
  });

  // Supprimer les groupes CP retirés
  const idxToRemove = checked.map(({i}) => i).sort((a,b) => b-a);
  idxToRemove.forEach(i => g_cpData.splice(i, 1));
  renderCP(); renderMaster(); renderKanban(); updateStats(); markDirty();
  toast(`${checked.length} groupe(s) CP retirés — dossiers remis en Solo.`);
};
document.getElementById('file-input').addEventListener('change', e => {
  const f=e.target.files[0]; if (!f) return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try {
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'});
      importRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''}));
    } catch(err){ alert('Erreur : '+err.message); }
  };
  reader.readAsArrayBuffer(f);
  e.target.value='';
});

function importRows(rows) {
  const grouped={}, order=[];
  let skipped=0;
  const today=new Date().toISOString().slice(0,10);

  for (let i=1;i<rows.length;i++) {
    const row=rows[i];
    const file=String(row[0]||'').replace(/[^\x20-\x7E]/g, '').trim();
    if (!file) continue;
    // Déjà présent → on met à jour les champs vides seulement
    const existing = g_master.find(r => r.file===file || (r.file||'').includes(file));
    if (existing) { skipped++; continue; }

    const rdl   =String(row[1]||'').trim();
    const svct  =String(row[22]||'').trim().replace(/;/g,'');
    const client=String(row[23]||'').trim();
    const city  =String(row[24]||'').trim();
    const poids =parseFloat(String(row[25]||'0').replace(',','.'))||0;
    const vol   =parseFloat(String(row[26]||'0').replace(',','.'))||0;
    const operator=String(row[27]||'').trim();
    const supplement=parseFloat(String(row[28]||'0').replace(',','.'))||0;
    const tax   =calcTaxable(poids,vol);
    let isCons=false;
    for (let c=8;c<=14;c++){
      if ((String(row[c]||'')).includes('Consolidat')||(String(row[c]||'')).includes('Complete')){isCons=true;break;}
    }
    const dept=/^\d{2}/.test(city)?city.substring(0,2):'??';
    storeRaw(file,client,svct,poids,vol,tax,dept,rdl,supplement);
    const key=(client+'|'+city).toLowerCase();
    if (!grouped[key]){
      // ccFiles = fichiers déjà CC (isCons), nonFiles = pas encore CC
      grouped[key]={ccFiles:[],nonFiles:[],client,rdl,svct,city,
        poids:0,vol:0,taxable:0,dept,operator};
      order.push(key);
    }
    const g=grouped[key];
    if (isCons) g.ccFiles.push(file); else g.nonFiles.push(file);
    g.poids+=poids; g.vol+=vol;
    g.taxable=calcTaxable(g.poids,g.vol);
    g.svct=mergeSVCT(g.svct,svct);
  }

  let added=0;
  for (const key of order) {
    const g=grouped[key];
    const cpFlag=isCP(g.client);
    const operator=g.operator||'';

    // ── Fichiers CC (isCons) ─────────────────────────────────────────────
    if (g.ccFiles.length > 0) {
      if (cpFlag || ['S8','SZ','SY'].some(c=>(g.svct||'').includes(c))) {
        // CP ou SVCT forbid → chaque file solo, statut 2
        g.ccFiles.forEach(sf => {
          const raw=g_rawData[sf]||{};
          const p=raw.poids||0, v=raw.vol||0, t=calcTaxable(p,v);
          const tr=calcTransp(g.svct,t,g.dept,cpFlag,p);
          const dup=g_master.find(x=>x.file===sf);
          if (dup) {
            Object.assign(dup,{client:g.client,rdl:g.rdl,svct:g.svct,transp:tr,
              poids:p,vol:v,taxable:t,dept:g.dept,cc:'Cc',statut:'2',operator});
          } else {
            g_master.push({file:sf,client:g.client,rdl:g.rdl,svct:g.svct,transp:tr,
              poids:p,vol:v,taxable:t,dept:g.dept,contact:'',tel:'',email:'',
              cc:'Cc',comment:'[Solo import]',statut:'2',operator,_dateCreated:today});
          }
          autoFillContact(sf,g.client);
          if (cpFlag) addOrUpdateCP(g.client,sf,p,v,t,operator);
        });
      } else {
        // Groupe CC → une ligne groupée
        let ccP=0, ccV=0;
        g.ccFiles.forEach(sf=>{const r=g_rawData[sf]||{};ccP+=r.poids||0;ccV+=r.vol||0;});
        const ccT=calcTaxable(ccP,ccV);
        const ccTr=calcTransp(g.svct,ccT,g.dept,false,ccP);
        const st=g.svct.includes('S5')?'4':'5';
        g_master.push({file:g.ccFiles.join(' + '),client:g.client,rdl:g.rdl,svct:g.svct,
          transp:ccTr,poids:+ccP.toFixed(2),vol:+ccV.toFixed(3),taxable:+ccT.toFixed(2),
          dept:g.dept,contact:'',tel:'',email:'',cc:'Cc',
          comment:g.ccFiles.length>1?'[Groupe Auto]':'',statut:st,operator,_dateCreated:today});
      }
      added++;
    }

    // ── Fichiers non-CC → chacun en statut 1, attend la CC
    if (g.nonFiles.length > 0) {
      g.nonFiles.forEach(sf => {
        const raw=g_rawData[sf]||{};
        const p=raw.poids||0, v=raw.vol||0, t=calcTaxable(p,v);
        const tr=calcTransp(g.svct,t,g.dept,cpFlag,p);
        const dup=g_master.find(x=>x.file===sf);
        if (dup) {
          Object.assign(dup,{client:g.client,rdl:g.rdl,svct:g.svct,transp:tr,
            poids:p,vol:v,taxable:t,dept:g.dept,operator});
        } else {
          g_master.push({file:sf,client:g.client,rdl:g.rdl,svct:g.svct,transp:tr,
            poids:p,vol:v,taxable:t,dept:g.dept,contact:'',tel:'',email:'',
            cc:'Non',comment:'',statut:'1',operator,_dateCreated:today});
        }
        autoFillContact(sf,g.client);
      });
      added++;
    }
  }

  // Pré-remplir contacts depuis g_contacts
  g_master.forEach(r => autoFillContact(r.file,r.client));

  renderMaster(); renderKanban(); renderCP(); updateStats();
  const infoEl=document.getElementById('import-info');
  document.getElementById('import-info-text').textContent=
    `${added} dossier(s) ajouté(s) · ${skipped} déjà présent(s) (données conservées)`;
  infoEl.style.display='flex';
  toast(skipped > 0 && added === 0
    ? `Import : aucun nouveau dossier (${skipped} déjà présents).`
    : `Import terminé : +${added} dossiers${skipped ? `, ${skipped} déjà présents (ignorés)` : ''}.`);
}

// ╔══════════════════════════════════════════════════════════════════════════╗

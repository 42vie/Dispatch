// ║  MÉTIER                                                                  ║
// ╚══════════════════════════════════════════════════════════════════════════╝
const IDF = new Set(['75','77','78','91','92','93','94','95','60']);

function calcTaxable(p, v) {
  return Math.round(Math.max(+p||0, (+v||0)*333) * 100) / 100;
}
function calcTransp(svct, tax, dept, cpFlag, poidsBrut) {
  // poidsBrut = poids brut réel (pas taxable) — c'est lui qui détermine le transporteur
  const brut = poidsBrut !== undefined ? +poidsBrut : +tax; // fallback sur taxable si brut absent
  const isIDF = IDF.has(String(dept||'').trim());
  // UPS uniquement si SVCT = S5, ST, ou vide
  const sv = (svct||'').toUpperCase().trim();
  const upsOK = (sv === '' || sv === 'S5' || sv === 'ST' || sv.includes('S5') || sv.includes('ST'));

  // Règle 1 : 0–10 kg BRUT + SVCT compatible (S5/ST/vide) → UPS partout
  if (brut <= 10 && upsOK) return 'UPS';

  // Règle 2 : IDF (75,77,78,91,92,93,94,95,60) + > 10 kg brut (ou noUPS) → Flex
  if (isIDF) return 'Flex (7)';

  // Règle 3 : Hors IDF
  // CP hors IDF → EFDS / Groussard
  if (cpFlag) return 'EFDS (1) / Groussard (12)';

  // 0–90 kg brut hors IDF (inclut noUPS <= 10kg) → DGS
  if (brut <= 90) return 'DGS (13)';

  // > 90 kg brut hors IDF → EFDS / Groussard
  return 'EFDS (1) / Groussard (12)';
}
function isCP(client) {
  // Matching partiel : le mot-clé CP doit être contenu dans le nom client
  // ex: "Arrow" reconnaît "Arrow ECS SAS", "MC3 LOGISTIQUE" reconnaît "MC3 LOGISTIQUE SAS"
  const cl = (client||'').toLowerCase().trim();
  const names = (typeof _cpNames !== 'undefined' && _cpNames.length)
    ? _cpNames
    : ['arrow','scc','computacenter','also','mc3 logistique','dexxon'];
  return names.some(c => c && cl.includes(c));
}
function mergeSVCT(cur, nw) {
  if (['SY','SZ','S8'].some(c => (nw||'').includes(c))) return nw;
  if ((nw||'').includes('S5') && !['SY','SZ','S8'].some(c => (cur||'').includes(c))) return nw;
  return cur;
}
function transpTarget(svct, client) {
  // CP : toujours RDV à envoyer (statut 2), quel que soit le SVCT
  if (isCP(client)) return '2';
  if (['S8','SZ','SY'].some(c=>(svct||'').includes(c))) return '2';
  if ((svct||'').includes('S5')) return '4';
  return '5';
}
function statutLabel(n) {
  const L = {'1':'1. Pas encore CC','2':'2. RDV à envoyer','3':'3. Attente réponse',
    '4':'4. Pré-alerte & FC','5':'5. Prêt pour FC','6':'6. COMAT','7':'7. NTO','8':'8. Terminé'};
  return L[String(n)] || String(n);
}
function statutNum(s) { const m = String(s||'').match(/^(\d)/); return m?m[1]:'1'; }
function statutClass(n) {
  return ['b1','b1','b2','b3','b4','b5','b6','b7','b8'][+n]||'b1';
}
function daysSince(dateStr) {
  if (!dateStr) return 0;
  const d = new Date(dateStr); d.setHours(0,0,0,0);
  return Math.floor((TODAY - d) / 86400000);
}
function daysBadge(n) {
  if (n === 0) return `<span class="days-badge days-0">Aujourd'hui</span>`;
  if (n === 1) return `<span class="days-badge days-1">J+1</span>`;
  if (n <= 2)  return `<span class="days-badge days-2">J+${n}</span>`;
  return `<span class="days-badge days-3">J+${n} ⚠</span>`;
}
function storeRaw(file, client, svct, poids, vol, taxable, dept, rdl, supplement) {
  if (!g_rawData[file]) {
    g_rawData[file] = {client,svct,poids:+poids,vol:+vol,taxable:+taxable,dept,rdl,supplement:+(supplement||0)};
  }
}
function autoFillContact(file, client) {
  const c = g_contacts.find(c => (c.societe||'').toLowerCase() === (client||'').toLowerCase());
  if (!c) return;
  const idx = g_master.findIndex(r => r.file === file);
  if (idx < 0) return;
  if (!g_master[idx].contact && c.nom)   g_master[idx].contact = c.nom;
  if (!g_master[idx].tel     && c.tel)   g_master[idx].tel     = c.tel;
  if (!g_master[idx].email   && c.email) g_master[idx].email   = c.email;
}

// ╔══════════════════════════════════════════════════════════════════════════╗

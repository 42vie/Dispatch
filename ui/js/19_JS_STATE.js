// ║  STATE                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝
// ── Chemins réseau par défaut (hardcodés) ──
const DEFAULT_STATE_PATH = 'F:\\CDG\\PRODUCT\\TRANSCON\\Shared\\Clients\\HPE\\Workboard\\dispatch_state.json';

let g_master   = [];  // dossiers actifs
let g_operatorFilter = localStorage.getItem('dispatch_operator') || ''; // restauré immédiatement
let g_rawData  = {};  // {fileId -> {client,svct,poids,vol,taxable,dept,rdl}}
let g_cpData   = [];  // [{client,files[],palettes,colis,poids,vol,taxable,doc}]
let g_contacts = [];  // [{societe,nom,tel,email,cp,notes}]
let g_editIdx  = -1;
let g_groupIdx = -1;
let g_conIdx   = -1;
let g_dragFile = null;
let g_sortCol  = null;
let g_sortAsc  = true;
let g_colFilters = {}; // {colName: Set of allowed values}

const TODAY = new Date(); TODAY.setHours(0,0,0,0);

// ── Utilitaires perf ──
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }

// ╔══════════════════════════════════════════════════════════════════════════╗

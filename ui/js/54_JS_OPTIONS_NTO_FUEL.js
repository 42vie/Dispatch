// ║  OPTIONS — NTO FUEL                                                       ║
// ╚══════════════════════════════════════════════════════════════════════════╝
function optsSaveFuel() {
  const vGlobal = parseFloat(document.getElementById('opts-fuel-val').value || '0');
  const vFlex   = parseFloat(document.getElementById('opts-fuel-flex').value || '0');
  const vDgs    = parseFloat(document.getElementById('opts-fuel-dgs').value || '0');
  // Si global est renseigné, il écrase les deux
  if (vGlobal > 0) {
    localStorage.setItem('nto_fuel_override', String(vGlobal));
    localStorage.setItem('nto_fuel_flex', String(vGlobal));
    localStorage.setItem('nto_fuel_dgs', String(vGlobal));
  } else {
    localStorage.setItem('nto_fuel_override', '0');
    if (vFlex > 0) localStorage.setItem('nto_fuel_flex', String(vFlex));
    else localStorage.removeItem('nto_fuel_flex');
    if (vDgs > 0) localStorage.setItem('nto_fuel_dgs', String(vDgs));
    else localStorage.removeItem('nto_fuel_dgs');
  }
  const st = document.getElementById('opts-fuel-status');
  if (st) {
    if (vGlobal > 0) st.textContent = '✓ Fuel global ' + vGlobal.toFixed(1) + '%';
    else if (vFlex > 0 || vDgs > 0) st.textContent = '✓ Flex=' + (vFlex||'auto') + '% DGS=' + (vDgs||'auto') + '%';
    else st.textContent = '✓ Mode auto activé';
    setTimeout(() => st.textContent = '', 3000);
  }
  if (typeof ntoInitFuel === 'function') ntoInitFuel();
}

function optsResetFuel() {
  localStorage.removeItem('nto_fuel_override');
  localStorage.removeItem('nto_fuel_flex');
  localStorage.removeItem('nto_fuel_dgs');
  document.getElementById('opts-fuel-val').value = '0';
  document.getElementById('opts-fuel-flex').value = '0';
  document.getElementById('opts-fuel-dgs').value = '0';
  const st = document.getElementById('opts-fuel-status');
  if (st) { st.textContent = '✓ Mode auto — tableau mensuel'; setTimeout(() => st.textContent = '', 2500); }
  if (typeof ntoInitFuel === 'function') ntoInitFuel();
}

function optsLoadFuel() {
  const v = localStorage.getItem('nto_fuel_override') || '';
  const el = document.getElementById('opts-fuel-val');
  if (el) el.value = v || '0';
  const elFlex = document.getElementById('opts-fuel-flex');
  const elDgs  = document.getElementById('opts-fuel-dgs');
  if (elFlex) elFlex.value = localStorage.getItem('nto_fuel_flex') || '0';
  if (elDgs)  elDgs.value  = localStorage.getItem('nto_fuel_dgs')  || '0';
  // Remplir le tableau historique
  const tbl = document.getElementById('opts-fuel-table');
  if (tbl && typeof NTO_FUEL !== 'undefined') {
    const now = new Date();
    const curKey = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
    tbl.innerHTML = Object.entries(NTO_FUEL).reverse().map(([k, r]) => {
      const isCur = k === curKey;
      return '<div style="padding:4px 6px;border-radius:3px;' + (isCur ? 'background:var(--purple-bg);border:1px solid var(--purple);' : '') + '">'
        + '<span style="color:var(--text3)">' + k + '</span> '
        + '<strong style="color:' + (isCur ? 'var(--purple)' : 'var(--text)') + '">' + (r*100).toFixed(1) + '%</strong>'
        + (isCur ? ' ◀' : '') + '</div>';
    }).join('');
  }
}


// ╔══════════════════════════════════════════════════════════════════════════╗

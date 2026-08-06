// ║  OFFLINE INDICATOR — polling serveur AU3                                 ║
// ╚══════════════════════════════════════════════════════════════════════════╝
let _serverOnline = true;
function _checkServerStatus() {
  fetch(API_URL + '/ping', { method: 'POST', body: '{}' })
    .then(r => { if (!_serverOnline) { _serverOnline = true; _updateOnlineIndicator(); } })
    .catch(() => { if (_serverOnline) { _serverOnline = false; _updateOnlineIndicator(); } });
}
function _updateOnlineIndicator() {
  let el = document.getElementById('offline-banner');
  if (!_serverOnline) {
    if (!el) {
      el = document.createElement('div');
      el.id = 'offline-banner';
      el.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:10000;background:var(--red);color:#fff;text-align:center;padding:5px 0;font-size:12px;font-weight:600;letter-spacing:.5px';
      el.textContent = 'HORS LIGNE — Serveur AutoIt non joignable';
      document.body.appendChild(el);
    }
  } else {
    if (el) el.remove();
  }
}
setInterval(_checkServerStatus, 30000);

// ╔══════════════════════════════════════════════════════════════════════════╗

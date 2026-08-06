// ║  TOAST                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝
let _toastTimer=null;
function toast(msg) {
  const el=document.getElementById('toast');
  el.textContent=msg; el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer=setTimeout(()=>el.classList.remove('show'),3000);
}

// ╔══════════════════════════════════════════════════════════════════════════╗

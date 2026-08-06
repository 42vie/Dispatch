// ║  UNDO — historique de 5 actions                                          ║
// ╚══════════════════════════════════════════════════════════════════════════╝
const _undoStack = [];
const UNDO_MAX = 5;
function pushUndo(label) {
  _undoStack.push({ label, master: JSON.parse(JSON.stringify(g_master)), cpData: JSON.parse(JSON.stringify(g_cpData)) });
  if (_undoStack.length > UNDO_MAX) _undoStack.shift();
}
function undoLast() {
  if (!_undoStack.length) return toast('Rien à annuler.');
  const snap = _undoStack.pop();
  g_master = snap.master;
  g_cpData = snap.cpData;
  renderMaster(); renderKanban(); renderCP(); updateStats(); markDirty();
  toast('Annulé : ' + snap.label);
}

// ╔══════════════════════════════════════════════════════════════════════════╗

// ║  RDL OVERDUE ALERT                                                       ║
// ╚══════════════════════════════════════════════════════════════════════════╝
function isRdlOverdue(rdl) {
  if (!rdl) return false;
  const parts = rdl.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (!parts) return false;
  let y = +parts[3]; if (y < 100) y += 2000;
  const d = new Date(y, +parts[2]-1, +parts[1]); d.setHours(0,0,0,0);
  return d < TODAY;
}

// ╔══════════════════════════════════════════════════════════════════════════╗

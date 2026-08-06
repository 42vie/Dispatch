// ║  TABS                                                                    ║
// ╚══════════════════════════════════════════════════════════════════════════╝
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
    document.querySelectorAll('.tabpane').forEach(p=>p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('tab-'+tab.dataset.tab).classList.add('active');
    if (tab.dataset.tab==='workflow') renderKanban();
    if (tab.dataset.tab==='contacts') renderContacts();
  });
});

// ╔══════════════════════════════════════════════════════════════════════════╗

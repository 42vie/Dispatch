// ║  SAVE / LOAD JSON                                                        ║
// ╚══════════════════════════════════════════════════════════════════════════╝
document.getElementById('btn-save').onclick = () => {
  const data=JSON.stringify({master:g_master,rawData:g_rawData,cpData:g_cpData,contacts:g_contacts},null,2);
  downloadText('DispatchMaster_'+new Date().toISOString().slice(0,16).replace('T','_')+'.json', data);
  toast('Sauvegarde téléchargée.');
};

// ╔══════════════════════════════════════════════════════════════════════════╗

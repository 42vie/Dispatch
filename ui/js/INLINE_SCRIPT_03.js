(function(){
'use strict';
var wfFilter = localStorage.getItem('dispatch_workflow_transp_filter') || 'all';
function id(x){ return document.getElementById(x); }
function rows(){ try { return Array.isArray(g_master) ? g_master : []; } catch(e){ return []; } }
function stat(r){ try { return String(typeof statutNum === 'function' ? statutNum(r.statut) : (String(r.statut || '1').match(/^(\d)/) || ['','1'])[1]); } catch(e){ return '1'; } }
function norm(s){ return String(s || '').toLowerCase(); }
function inWorkflow(r){ var n = Number(stat(r)); return n >= 1 && n <= 9; }
function isCp(r){ try { if(typeof isCP === 'function') return isCP(r.client); } catch(e){} var c = norm(r.client); return /arrow|scc|computacenter|also|dexcel|dexxon|ingram|synnex|mc3/.test(c); }
function matchFilter(r, f){
  var tr = norm(r.transp), client = norm(r.client), email = String(r.email || '').trim();
  if(f === 'all') return true;
  if(f === 'dgs') return tr.includes('dgs');
  if(f === 'flex') return tr.includes('flex') || tr.includes('(7)');
  if(f === 'ups') return tr.includes('ups');
  if(f === 'efds') return tr.includes('efds') || client.includes('efds');
  if(f === 'groussard') return tr.includes('groussard');
  if(f === 'cp') return isCp(r);
  if(f === 'noemail') return !email;
  if(f === 'groups') return String(r.file || '').includes(' + ');
  return true;
}
function count(f){ return rows().filter(function(r){ return inWorkflow(r) && matchFilter(r, f); }).length; }
function injectPanel(){
  var tab = id('tab-workflow'); if(!tab || id('cr-workflow-filter-panel')) return;
  var kanban = id('kanban-wrap'); if(!kanban) return;
  var panel = document.createElement('div');
  panel.id = 'cr-workflow-filter-panel';
  panel.innerHTML = '<div id="cr-workflow-chip-row"></div><div id="cr-workflow-filter-count"></div>';
  tab.insertBefore(panel, kanban);
  panel.querySelector('#cr-workflow-chip-row').addEventListener('click', function(e){
    var b = e.target.closest('[data-wf-filter]'); if(!b) return;
    wfFilter = b.dataset.wfFilter || 'all';
    localStorage.setItem('dispatch_workflow_transp_filter', wfFilter);
    updateChips(); applyWorkflowFilter();
  });
  updateChips();
}
function updateChips(){
  var row = id('cr-workflow-chip-row'); if(!row) return;
  var chips = [
    ['all','Tous '+count('all'),''],
    ['dgs','DGS '+count('dgs'),''],
    ['flex','Flex '+count('flex'),''],
    ['ups','UPS '+count('ups'),''],
    ['efds','EFDS '+count('efds'),''],
    ['groussard','Groussard '+count('groussard'),''],
    ['cp','CP '+count('cp'),'good'],
    ['noemail','Sans email '+count('noemail'),'warn'],
    ['groups','Groupes '+count('groups'),'']
  ];
  row.innerHTML = chips.map(function(c){
    return '<button class="cr-wf-chip '+c[2]+' '+(wfFilter===c[0]?'active':'')+'" data-wf-filter="'+c[0]+'">'+c[1]+'</button>';
  }).join('');
}
function recordForCard(card){
  var file = card && card.dataset ? card.dataset.file : '';
  if(!file) return null;
  return rows().find(function(r){ return r.file === file; }) || null;
}
function applyWorkflowFilter(){
  var total = 0, shown = 0;
  document.querySelectorAll('#kanban-wrap .k-item').forEach(function(card){
    var r = recordForCard(card); total++;
    var show = !r || matchFilter(r, wfFilter);
    card.classList.toggle('cr-wf-filter-hidden', !show);
    if(show) shown++;
  });
  var c = id('cr-workflow-filter-count');
  if(c) c.textContent = wfFilter === 'all' ? '' : shown + '/' + total + ' cartes affichées';
  updateChips();
}
function patchRenderKanban(){
  try{
    if(typeof renderKanban === 'function' && !renderKanban.__wfChipsV11){
      var old = renderKanban;
      var w = function(){
        var res = old.apply(this, arguments);
        setTimeout(function(){ injectPanel(); applyWorkflowFilter(); if(typeof addCmrToKanban === 'function') addCmrToKanban(); }, 0);
        return res;
      };
      w.__wfChipsV11 = true;
      renderKanban = w;
      window.renderKanban = w;
    }
  } catch(e){ console.warn('workflow chips patch renderKanban', e); }
}
function init(){ injectPanel(); patchRenderKanban(); setTimeout(applyWorkflowFilter, 80); }
if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', function(){ setTimeout(init, 450); }); else setTimeout(init, 450);
window.addEventListener('load', function(){ setTimeout(init, 900); });
})();

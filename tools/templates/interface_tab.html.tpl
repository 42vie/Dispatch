<!-- ============================================================
     {{NAME_UPPER}}_UI_V1 - onglet {{NAME_UPPER}} (genere par tools\NewFeature.au3 le {{DATE}})
     ============================================================ -->
<style id="{{NAME_LOWER}}-ui-v1-style">
#tab-{{NAME_LOWER}}{height:100%;overflow:auto;background:var(--bg);}
.{{NAME_LOWER}}-page{padding:18px;display:flex;flex-direction:column;gap:14px;max-width:900px;margin:0 auto;}
.{{NAME_LOWER}}-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow-sm);padding:18px;}
.{{NAME_LOWER}}-card h3{margin:0 0 8px;font-size:15px;font-weight:900;color:var(--text);}
.{{NAME_LOWER}}-sub{font-size:11.5px;color:var(--text3);line-height:1.45;margin-bottom:12px;}
.{{NAME_LOWER}}-row2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.{{NAME_LOWER}}-field input{width:100%;border:1px solid var(--border2);border-radius:9px;background:var(--surface2);color:var(--text);font-size:12.5px;padding:9px 10px;outline:none;box-sizing:border-box;}
.{{NAME_LOWER}}-actions{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 2px;}
.{{NAME_LOWER}}-table-wrap{max-height:320px;overflow:auto;border:1px solid var(--border);border-radius:12px;margin-top:10px;}
.{{NAME_LOWER}}-table{width:100%;border-collapse:collapse;font-size:11.5px;}
.{{NAME_LOWER}}-table th{position:sticky;top:0;background:var(--surface2);text-align:left;padding:8px 9px;border-bottom:1px solid var(--border);color:var(--text2);}
.{{NAME_LOWER}}-table td{padding:7px 9px;border-bottom:1px solid var(--border);vertical-align:middle;}
.{{NAME_LOWER}}-empty{color:var(--text3);font-size:11.5px;text-align:center;padding:18px;}
@media(max-width:640px){.{{NAME_LOWER}}-row2{grid-template-columns:1fr}}
</style>
<script id="{{NAME_LOWER}}-ui-v1-script">
(function(){
'use strict';
var items=[];
function id(x){return document.getElementById(x)}
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function toastMsg(s){if(typeof toast==='function')toast(s);else console.log(s)}
function apiAction(payload){if(typeof apiCall==='function')return apiCall('action',payload);return fetch('/api/action',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(function(r){return r.json().catch(function(){return {status:r.ok?'ok':'error'}});});}

function renderItems(){
  var b=id('{{NAME_LOWER}}-body'); if(!b) return;
  b.innerHTML = items.length ? items.map(function(it){
    return '<tr><td><b>'+esc(it.label||it.id)+'</b></td><td>'+esc(it.value)+'</td><td>'+esc(it.date)+'</td>'+
      '<td><button class="btn btn-ghost {{NAME_LOWER}}-mini" data-del="'+esc(it.id)+'" style="height:22px;padding:0 7px;font-size:10px">🗑</button></td></tr>';
  }).join('') : '<tr><td colspan="4" class="{{NAME_LOWER}}-empty">Aucun element.</td></tr>';
}
function load(){
  apiAction({action:'{{NAME_UPPER}}_LIST'}).then(function(r){
    items = (r && r.items) || [];
    renderItems();
  }).catch(function(e){ toastMsg('Erreur chargement : '+(e.message||e)); });
}
function add(){
  var label=id('{{NAME_LOWER}}-label').value.trim();
  var value=id('{{NAME_LOWER}}-value').value.trim();
  if(!label){ toastMsg('Le libelle est obligatoire.'); return; }
  apiAction({action:'{{NAME_UPPER}}_ADD', label:label, value:value}).then(function(r){
    if(r && r.status==='ok'){ id('{{NAME_LOWER}}-label').value=''; id('{{NAME_LOWER}}-value').value=''; toastMsg('Ajoute.'); load(); }
    else toastMsg('Erreur : '+((r&&r.message)||'inconnue'));
  }).catch(function(e){ toastMsg('Erreur : '+(e.message||e)); });
}
function del(itemId){
  apiAction({action:'{{NAME_UPPER}}_DELETE', id:itemId}).then(function(){ toastMsg('Supprime.'); load(); })
    .catch(function(e){ toastMsg('Erreur : '+(e.message||e)); });
}

function ensureTab(){
  var tabbar=id('tabbar'), content=id('content'); if(!tabbar||!content) return;
  var tab=document.querySelector('.tab[data-tab="{{NAME_LOWER}}"]');
  if(!tab){ tab=document.createElement('div'); tab.className='tab'; tab.dataset.tab='{{NAME_LOWER}}'; tab.textContent='{{TAB_LABEL}}';
    var opt=document.querySelector('.tab[data-tab="options"]'); tabbar.insertBefore(tab, opt||null); }
  var pane=id('tab-{{NAME_LOWER}}');
  if(!pane){
    pane=document.createElement('div'); pane.className='tabpane'; pane.id='tab-{{NAME_LOWER}}';
    pane.innerHTML = '<div class="{{NAME_LOWER}}-page">' +
      '<section class="{{NAME_LOWER}}-card">' +
      '<h3>{{TAB_LABEL}}</h3>' +
      '<div class="{{NAME_LOWER}}-sub">Onglet genere automatiquement -- personnalise les champs, le stockage et la logique dans src/features/{{NAME_LOWER}}/{{NAME_UPPER}}_Tool.au3.</div>' +
      '<div class="{{NAME_LOWER}}-row2"><div class="{{NAME_LOWER}}-field"><input id="{{NAME_LOWER}}-label" type="text" placeholder="Libelle"></div><div class="{{NAME_LOWER}}-field"><input id="{{NAME_LOWER}}-value" type="text" placeholder="Valeur"></div></div>' +
      '<div class="{{NAME_LOWER}}-actions"><button class="btn btn-success btn-sm" id="{{NAME_LOWER}}-add">+ Ajouter</button><button class="btn btn-ghost btn-sm" id="{{NAME_LOWER}}-reload">Rafraichir</button></div>' +
      '<div class="{{NAME_LOWER}}-table-wrap"><table class="{{NAME_LOWER}}-table"><thead><tr><th>Libelle</th><th>Valeur</th><th>Date</th><th></th></tr></thead><tbody id="{{NAME_LOWER}}-body"></tbody></table></div>' +
      '</section></div>';
    content.appendChild(pane);
    id('{{NAME_LOWER}}-add').onclick=add;
    id('{{NAME_LOWER}}-reload').onclick=load;
    id('{{NAME_LOWER}}-body').addEventListener('click', function(e){
      var d=e.target.closest('[data-del]'); if(!d) return;
      if(confirm('Supprimer "'+d.dataset.del+'" ?')) del(d.dataset.del);
    });
  }
}

function switchToTab(){
  document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active')});
  document.querySelectorAll('.tabpane').forEach(function(p){p.classList.remove('active')});
  var t=document.querySelector('.tab[data-tab="{{NAME_LOWER}}"]'), p=id('tab-{{NAME_LOWER}}');
  if(t) t.classList.add('active'); if(p) p.classList.add('active');
  load();
}

document.addEventListener('click', function(e){
  var tab=e.target.closest('.tab[data-tab="{{NAME_LOWER}}"]');
  if(tab){ e.preventDefault(); e.stopPropagation(); ensureTab(); switchToTab(); }
}, true);

function init(){ ensureTab(); }
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', function(){ setTimeout(init,700); });
else setTimeout(init,700);
window.addEventListener('load', function(){ setTimeout(init,1300); });
})();
</script>

# Dispatch HPE — Documentation & Roadmap

> Application interne de gestion des dossiers de transport — Expeditors CDG  
> Stack : HTML/CSS/JS (vanilla) · AutoIt · Réseau F:\

---

## Table des matières

1. [Présentation](#1-présentation)
2. [Architecture actuelle](#2-architecture-actuelle)
3. [Fonctionnalités existantes](#3-fonctionnalités-existantes)
4. [Roadmap UI/UX — Améliorations préconisées](#4-roadmap-uiux--améliorations-préconisées)
   - 4.1 Dashboard opérationnel
   - 4.2 Chips filtres rapides
   - 4.3 Indicateur de synchronisation réseau
   - 4.4 Barre de progression par dossier (Kanban)
   - 4.5 Vue « Mon Desk » persistante
   - 4.6 Side Panel (drawer) à la place de la modale
   - 4.7 Timeline dossier
   - 4.8 Raccourcis clavier — cheat sheet
5. [Améliorations fonctionnelles](#5-améliorations-fonctionnelles)
6. [Priorités recommandées](#6-priorités-recommandées)
7. [Convention de code](#7-convention-de-code)

---

## 1. Présentation

**Dispatch HPE** est un outil de pilotage des dossiers de livraison HPE géré par l'équipe transport Expeditors CDG.  
Il couvre l'intégralité du workflow : import Excel E.TMS → CC → RDV → Pré-alerte → File Closing → COMAT → NTO → Terminé.

L'interface tourne dans un navigateur local (fichier HTML standalone) et communique avec AutoIt pour piloter E.TMS.

---

## 2. Architecture actuelle

```
Interface.html        ← UI complète (HTML + CSS + JS inline)
validate.js           ← Validation des données
migrate.js            ← Migration schéma localStorage
merge.js              ← Fusion dossiers / groupes
DataManager.js        ← Lecture/écriture JSON réseau (F:\)
```

**Stockage :**
- `localStorage` du navigateur (données actives)
- `dispatch_state.json` sur réseau partagé `F:\` (sync multi-postes)

**Communication AutoIt :**  
Via fichier `action.json` posé sur le réseau, lu en polling par le script `.au3`.

---

## 3. Fonctionnalités existantes

| Module | Description |
|---|---|
| **Dispatch** | Table principale, tri/filtre colonnes, badges statut, jours sans réponse |
| **Workflow Kanban** | 8 colonnes statut, drag & drop, sélection multi, barre d'actions flottante |
| **Channel Partners** | Regroupement CP (Arrow, SCC, Computacenter…), envoi RDV groupé |
| **Contacts Clients** | Carnet d'adresses partagé, auto-complétion |
| **COMAT** | Automatisation E.TMS séquentielle (solo & batch) |
| **NTO** | Calcul tarification transport (Flex IDF / DGS Province / UPS) avec taux fuel |
| **Options** | Config réseau, fuel, PJ mails, diagnostic, stockage |
| **Dark mode** | Thème sombre complet (`Ctrl+D`) |

---

## 4. Roadmap UI/UX — Améliorations préconisées

### 4.1 Dashboard opérationnel ⭐ PRIORITÉ HAUTE

**Objectif :** Ajouter un nouvel onglet `Dashboard` avec une vue synthétique temps réel.

**Contenu préconisé :**

```
┌─────────────────────────────────────────────────────────────┐
│  KPIs du jour            Répartition par statut (donut)     │
│  ┌──────┐ ┌──────┐      ████ 1.CC   ██ 2.RDV   █ 3.Att     │
│  │  42  │ │  12  │      ██ 4.PA    ████ 5.FC   █ 6.COM      │
│  │Total │ │ CC ✓ │                                          │
│  └──────┘ └──────┘                                          │
│                                                             │
│  Alertes urgentes        Charge par opérateur               │
│  ⚠ 3 dossiers J+3        Mohamed  ██████ 14                 │
│  ⚠ 5 RDL dépassées       Sarah    ████ 9                    │
│  ⚠ 7 sans email          Karim    ███ 6                     │
│                                                             │
│  Évolution semaine (barres)    Top transporteurs            │
│  Lun Mar Mer Jeu Ven            DGS 45% / Flex 35% / UPS 20%│
└─────────────────────────────────────────────────────────────┘
```

**Implémentation :**
- Ajouter `<div class="tab" data-tab="dashboard">📊 Dashboard</div>` dans la tabbar
- Grille CSS (`display: grid; grid-template-columns: 1fr 1fr; gap: 16px`)
- Donut chart en SVG pur (pas de lib externe)
- Barres en CSS (`width: calc(valeur / max * 100%)`)
- Refresh automatique toutes les 30 secondes via `setInterval`

---

### 4.2 Chips filtres rapides ⭐ PRIORITÉ HAUTE

**Objectif :** Remplacer/compléter le menu opérateur par des boutons filtres 1-clic au-dessus de la table Dispatch.

**Chips préconisés :**

```
[ Tout ] [ ⚠ Retard J+2 ] [ 📧 Sans email ] [ ⏳ Sans contact ] [ 🔗 Groupables ] [ ✓ CC en attente ]
```

**CSS :**
```css
.chip {
  height: 24px; padding: 0 10px; border-radius: 12px;
  font-size: 11px; font-weight: 500; cursor: pointer;
  border: 1px solid var(--border2); background: var(--surface2);
  color: var(--text2); transition: all .15s;
}
.chip.active {
  background: var(--blue-bg); border-color: var(--blue);
  color: var(--blue); font-weight: 600;
}
```

**Logique JS :**
```javascript
const CHIP_FILTERS = {
  retard:    r => daysSince(r._dateCreated) >= 2 && +statutNum(r.statut) <= 7,
  no_email:  r => !r.email,
  no_contact:r => !r.contact && !r.tel,
  groupable: r => /* logique siblings existante */,
  cc_wait:   r => r.cc !== 'Cc' && +statutNum(r.statut) === 1,
};
```

---

### 4.3 Indicateur de synchronisation réseau ⭐ PRIORITÉ HAUTE

**Objectif :** Remplacer les boutons `⬆ Réseau` / `⬇ Réseau` peu lisibles par un indicateur visuel dans le header.

**Design préconisé :**
```
● Sync  14:32        (vert = OK)
● Sync  —            (orange = en attente)
● Erreur réseau      (rouge = échec)
```

**CSS :**
```css
.sync-dot {
  width: 8px; height: 8px; border-radius: 50%;
  flex-shrink: 0; transition: background .3s;
}
.sync-ok     { background: var(--green); }
.sync-pending{ background: var(--orange); animation: pulse 1.2s infinite; }
.sync-error  { background: var(--red); }

@keyframes pulse {
  0%, 100% { opacity: 1; }
  50%       { opacity: 0.4; }
}
```

**Conserver** les boutons ⬆⬇ en Options > Réseau pour les actions manuelles.

---

### 4.4 Barre de progression Kanban

**Objectif :** Ajouter une mini progress bar sur chaque carte Kanban pour visualiser l'avancement pipeline (1→8).

**HTML (dans renderKanban) :**
```html
<div class="k-progress">
  <div class="k-progress-bar" style="width: calc(${sn} / 8 * 100%); 
    background: var(${k.c})"></div>
</div>
```

**CSS :**
```css
.k-progress {
  height: 3px; background: var(--border);
  border-radius: 2px; margin-top: 5px; overflow: hidden;
}
.k-progress-bar {
  height: 100%; border-radius: 2px;
  transition: width .3s ease;
}
```

---

### 4.5 Vue « Mon Desk » persistante

**Objectif :** Mémoriser le filtre opérateur entre les sessions et ajouter un bouton toggle rapide.

**Comportement préconisé :**
- Au chargement : lire `localStorage.getItem('dispatch_operator')` → pré-sélectionner automatiquement ✅ (déjà partiellement fait)
- Ajouter un bouton **👤 Mon Desk** dans la toolbar qui toggle entre "mon filtre" et "Tous"
- Afficher le nom de l'opérateur actif dans le header avec un badge coloré

**JS :**
```javascript
function toggleMyDesk() {
  const myName = localStorage.getItem('dispatch_my_name') || '';
  if (!myName) return toast('Configurez votre prénom dans Options > Réseau');
  g_operatorFilter = g_operatorFilter ? '' : myName;
  localStorage.setItem('dispatch_operator', g_operatorFilter);
  renderMaster(); renderKanban(); updateStats();
}
```

---

### 4.6 Side Panel (drawer) — remplacement de la modale

**Objectif :** Remplacer la modale d'édition plein écran par un panneau latéral qui glisse depuis la droite, sans masquer la table.

**Structure HTML :**
```html
<div id="side-panel" class="side-panel">
  <div class="sp-header">
    <h3 id="sp-title">Édition dossier</h3>
    <button onclick="closeSidePanel()">✕</button>
  </div>
  <div class="sp-body"><!-- même contenu que modal-edit --></div>
  <div class="sp-footer"><!-- boutons --></div>
</div>
<div id="side-overlay" class="side-overlay" onclick="closeSidePanel()"></div>
```

**CSS :**
```css
.side-panel {
  position: fixed; top: 0; right: -480px; width: 480px; height: 100vh;
  background: var(--surface); border-left: 1px solid var(--border);
  box-shadow: -4px 0 24px rgba(0,0,0,.15);
  z-index: 300; transition: right .25s cubic-bezier(.4,0,.2,1);
  display: flex; flex-direction: column; overflow: hidden;
}
.side-panel.open { right: 0; }
.side-overlay {
  display: none; position: fixed; inset: 0;
  background: rgba(0,0,0,.25); z-index: 299;
}
.side-overlay.open { display: block; }
```

**Avantage :** L'utilisateur voit la table derrière pendant qu'il édite.

---

### 4.7 Timeline dossier

**Objectif :** Afficher l'historique des changements de statut dans le panneau d'édition.

**Structure de données à ajouter :**
```javascript
// Dans chaque record g_master :
r._history = r._history || [];
// À chaque changement de statut :
r._history.push({
  date: new Date().toISOString(),
  statut: newStatut,
  operator: r.operator || ''
});
```

**Rendu HTML :**
```html
<div class="timeline">
  <div class="tl-item">
    <div class="tl-dot" style="background:var(--green)"></div>
    <div class="tl-content">
      <span class="tl-label">5. Prêt pour FC</span>
      <span class="tl-date">04/08/2026 14:30 · Mohamed</span>
    </div>
  </div>
</div>
```

**CSS :**
```css
.timeline { display: flex; flex-direction: column; gap: 0; }
.tl-item  { display: flex; gap: 10px; align-items: flex-start; padding: 6px 0; }
.tl-dot   { width: 8px; height: 8px; border-radius: 50%; margin-top: 4px; flex-shrink: 0; }
.tl-content { font-size: 11px; }
.tl-label { font-weight: 600; color: var(--text); display: block; }
.tl-date  { color: var(--text3); }
```

---

### 4.8 Raccourcis clavier — cheat sheet

**Objectif :** Ajouter un bouton `?` dans le header qui affiche une modale avec tous les raccourcis.

**Raccourcis à documenter :**

| Raccourci | Action |
|---|---|
| `Ctrl+F` | Focuser la recherche |
| `Ctrl+S` | Sauvegarder |
| `Ctrl+D` | Basculer dark mode |
| `Ctrl+Z` | Annuler dernière action |
| `Échap` | Fermer modale / panel |
| `Entrée` (dans champ dossier) | Lancer action E.TMS |
| `Double-clic` (ligne table) | Ouvrir édition |
| `Double-clic` (carte Kanban) | Ouvrir édition |

**JS :**
```javascript
document.addEventListener('keydown', e => {
  if (e.key === '?') openModal('modal-shortcuts');
  if (e.key === 'Escape') closeAllModals();
});
```

---

## 5. Améliorations fonctionnelles

### 5.1 Sauvegarde automatique plus robuste
- Ajouter un `setInterval(autoSave, 5 * 60 * 1000)` toutes les 5 minutes
- Afficher un timestamp "Dernière sauvegarde : 14:32" dans le header
- Implémenter un système de versions (`dispatch_state_YYYYMMDD_HHMM.json`) pour rollback

### 5.2 Export PDF / rapport journalier
- Bouton **📄 Rapport PDF** dans le header
- Utiliser `window.print()` avec une feuille de style `@media print` dédiée
- Contenu : table Dispatch filtrée + KPIs du jour + date

### 5.3 Notifications navigateur
```javascript
// Alerter si un dossier dépasse J+3
if (Notification.permission === 'granted') {
  new Notification('Dispatch HPE', {
    body: `${overdueCount} dossier(s) sans réponse depuis J+3`,
    icon: '/favicon.ico'
  });
}
```

### 5.4 Import Excel amélioré
- Afficher un **preview** des lignes avant import (modal de confirmation)
- Détecter les colonnes automatiquement (mapping flexible) pour s'adapter aux exports E.TMS
- Rapport post-import : X ajoutés, Y mis à jour, Z ignorés

### 5.5 Recherche globale améliorée
- Étendre la recherche aux champs `email`, `tel`, `comment`, `transp`
- Surligner les termes trouvés dans les cellules
- Shortcut `Ctrl+G` pour recherche globale (tous onglets)

---

## 6. Priorités recommandées

### Phase 1 — Impact immédiat (1–2 jours)

| # | Amélioration | Effort | Impact |
|---|---|---|---|
| 1 | **Chips filtres rapides** | Faible | 🔴 Élevé |
| 2 | **Indicateur sync réseau** | Faible | 🔴 Élevé |
| 3 | **Vue Mon Desk toggle** | Faible | 🟠 Moyen |

### Phase 2 — Valeur ajoutée (3–5 jours)

| # | Amélioration | Effort | Impact |
|---|---|---|---|
| 4 | **Dashboard** | Moyen | 🔴 Élevé |
| 5 | **Barre progression Kanban** | Faible | 🟠 Moyen |
| 6 | **Cheat sheet raccourcis** | Faible | 🟡 Faible |

### Phase 3 — Expérience avancée (1–2 semaines)

| # | Amélioration | Effort | Impact |
|---|---|---|---|
| 7 | **Side Panel drawer** | Élevé | 🟠 Moyen |
| 8 | **Timeline dossier** | Moyen | 🟠 Moyen |
| 9 | **Export PDF rapport** | Moyen | 🟠 Moyen |
| 10 | **Notifications navigateur** | Faible | 🟡 Faible |

---

## 7. Convention de code

### Nommage
- **Variables globales** : préfixe `g_` (ex: `g_master`, `g_cpData`)
- **Fonctions render** : préfixe `render` (ex: `renderMaster`, `renderKanban`)
- **Fonctions métier** : camelCase explicite (ex: `calcTaxable`, `transpTarget`)
- **IDs HTML** : kebab-case (ex: `modal-edit`, `k-body-5`)
- **Classes CSS** : kebab-case (ex: `btn-primary`, `k-item`)

### CSS
- Toutes les couleurs passent par des variables CSS (`:root` / `html.dark`)
- Jamais de couleur hardcodée dans le JS ou le HTML
- Dark mode : toujours tester les deux thèmes avant de commiter

### Performances
- `renderMaster()` et `renderKanban()` reconstruisent le DOM → appeler uniquement si nécessaire
- Utiliser `debounce()` sur les événements `input`
- `g_master` est la source de vérité — ne jamais lire le DOM pour récupérer des données

### Sauvegarde
- Tout changement utilisateur → appeler `markDirty()` puis `autoSave()` (debounced 2s)
- Les données critiques → `reseauForceSave()` en plus

---

*Documentation générée le 04/08/2026 — Dispatch HPE v2.x*

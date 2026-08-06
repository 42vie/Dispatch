# Plan complet de modularisation Dispatch

## 1. Objectif

Modulariser le projet Dispatch actuel, composé de :

- `Dispatch.au3` : serveur HTTP local, API, actions ETMS/COMAT/FC/CMR/EDOC, mails, réseau, stockage et diagnostics.
- `Interface-2.html` : interface complète Dispatch avec table, Kanban, NTO, Options, Contacts, Workflow, Dashboard et CMR.

La modularisation doit fonctionner sur le PC du travail avec uniquement :

- AutoIt/SciTE déjà disponible ;
- le navigateur ;
- des fichiers `.au3`, `.html`, `.css`, `.js`, `.json`, `.tsv` et `.ini`.

Aucune installation de Node.js, npm, framework, bundler ou serveur externe n'est nécessaire.

Le comportement applicatif doit rester identique. On sépare le code, mais on conserve :

- les mêmes fichiers de données ;
- les mêmes endpoints API ;
- les mêmes variables globales ;
- le même `gmaster` côté navigateur ;
- la même UX ;
- un seul fichier AutoIt à compiler.

---

## 2. Architecture finale

```text
Dispatch/
│
├── MainDispatch.au3
├── Dispatch.exe
│
├── src/
│   ├── core/
│   │   ├── Globals.au3
│   │   ├── Constants.au3
│   │   ├── Config.au3
│   │   ├── HttpServer.au3
│   │   ├── HttpRouter.au3
│   │   ├── HttpResponse.au3
│   │   ├── JsonUtils.au3
│   │   ├── DateUtils.au3
│   │   ├── FileUtils.au3
│   │   ├── BackupUtils.au3
│   │   ├── Audit.au3
│   │   └── Security.au3
│   │
│   ├── api/
│   │   ├── ApiPing.au3
│   │   ├── ApiState.au3
│   │   ├── ApiContacts.au3
│   │   ├── ApiNetwork.au3
│   │   ├── ApiActions.au3
│   │   ├── ApiJobs.au3
│   │   └── ApiConfig.au3
│   │
│   ├── services/
│   │   ├── StateService.au3
│   │   ├── ContactsService.au3
│   │   ├── NetworkStateService.au3
│   │   ├── ConfigService.au3
│   │   └── JobService.au3
│   │
│   ├── features/
│   │   ├── etms/
│   │   │   └── Action_ETMS.au3
│   │   ├── comat/
│   │   │   ├── Action_COMAT.au3
│   │   │   └── Batch_COMAT.au3
│   │   ├── fc/
│   │   │   ├── Action_FC.au3
│   │   │   ├── Batch_FC.au3
│   │   │   └── Audit_FC.au3
│   │   ├── cmr/
│   │   │   ├── Action_CMR.au3
│   │   │   ├── CMR_Engine.au3
│   │   │   ├── CMR_Mail.au3
│   │   │   ├── CMR_EDOC.au3
│   │   │   └── CMR_PDF.au3
│   │   ├── edoc/
│   │   │   ├── Action_EDOC.au3
│   │   │   ├── EDOC_Web.au3
│   │   │   └── EDOC_Master.au3
│   │   ├── mail/
│   │   │   ├── Mail_Common.au3
│   │   │   ├── Mail_RDV.au3
│   │   │   ├── Mail_Alertes.au3
│   │   │   └── Mail_CP.au3
│   │   └── diagnostics/
│   │       ├── Diagnostic.au3
│   │       ├── StorageInfo.au3
│   │       └── ContactsCleanup.au3
│   │
│   └── legacy/
│       └── Dispatch_Legacy_Backup.au3
│
├── ui/
│   ├── index.html
│   ├── css/
│   │   ├── 00-reset.css
│   │   ├── 01-variables.css
│   │   ├── 02-layout.css
│   │   ├── 03-components.css
│   │   ├── 04-table.css
│   │   ├── 05-kanban.css
│   │   ├── 06-modals.css
│   │   ├── 07-options.css
│   │   ├── 08-dashboard.css
│   │   ├── 09-cmr.css
│   │   └── 10-responsive.css
│   └── js/
│       ├── 00-bootstrap.js
│       ├── 01-state.js
│       ├── 02-utils.js
│       ├── 03-api.js
│       ├── 04-storage.js
│       ├── 05-contacts.js
│       ├── 06-dispatch-table.js
│       ├── 07-kanban.js
│       ├── 08-groups.js
│       ├── 09-workflow.js
│       ├── 10-actions-etms.js
│       ├── 11-actions-comat.js
│       ├── 12-actions-fc.js
│       ├── 13-actions-cmr.js
│       ├── 14-actions-edoc.js
│       ├── 15-mail.js
│       ├── 16-nto.js
│       ├── 17-options.js
│       ├── 18-dashboard.js
│       ├── 19-modals.js
│       ├── 20-undo.js
│       ├── 21-sync.js
│       ├── 22-events.js
│       └── 99-init.js
│
├── dispatch.json
├── data.json
├── status.json
├── contacts.tsv
├── contacts.json
├── config.ini
├── audit.log
└── backups/
```

---

## 3. Règles générales

### 3.1 Un seul point d'entrée AutoIt

Le seul fichier ouvert dans SciTE et compilé est :

```text
MainDispatch.au3
```

Il doit contenir :

- les includes AutoIt standards ;
- les includes des modules ;
- `MainDispatch()` ;
- la boucle principale.

Les fichiers fonctionnels ne doivent pas contenir de deuxième `Main()`, de deuxième boucle serveur ou de deuxième initialisation globale.

### 3.2 Un seul fichier de globals

Tous les globals communs doivent être définis dans :

```text
src/core/Globals.au3
```

Les autres modules utilisent ces globals, mais ne les redéclarent pas.

### 3.3 Pas de modules JavaScript ES au départ

Ne pas utiliser dans la première migration :

```html
<script type="module">
```

Ne pas utiliser `import` ou `export`.

Les fichiers JavaScript sont chargés comme des scripts classiques dans un ordre contrôlé. Cela préserve la compatibilité avec les variables globales existantes : `gmaster`, `grawData`, `gcpData`, `gcontacts`, `apiCall`, `renderMaster`, `renderKanban`, etc.

### 3.4 Même contrat API

Les noms d'actions et endpoints actuels doivent rester inchangés :

- `/api/ping` ;
- `/api/load` ;
- `/api/save` ;
- `/api/load-data` ;
- `/api/save-data` ;
- `/api/load-status` ;
- `/api/save-status` ;
- `/api/load-contacts` ;
- `/api/save-contacts` ;
- `/api/net-save` ;
- `/api/net-load` ;
- `/api/net-list` ;
- `/api/net-check` ;
- `/api/action` ;
- `/api/job-status`.

---

## 4. Fichier `MainDispatch.au3`

```autoit
#NoTrayIcon
Opt("MustDeclareVars", 0)
Opt("GUIOnEventMode", 0)

#include <File.au3>
#include <String.au3>
#include <Date.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <WinAPI.au3>

; Globals et constantes
#include "src\core\Globals.au3"
#include "src\core\Constants.au3"

; Utilitaires
#include "src\core\JsonUtils.au3"
#include "src\core\DateUtils.au3"
#include "src\core\FileUtils.au3"
#include "src\core\BackupUtils.au3"
#include "src\core\Security.au3"
#include "src\core\Audit.au3"

; Configuration et services
#include "src\core\Config.au3"
#include "src\services\StateService.au3"
#include "src\services\ContactsService.au3"
#include "src\services\NetworkStateService.au3"
#include "src\services\ConfigService.au3"
#include "src\services\JobService.au3"

; Fonctions métier
#include "src\features\etms\Action_ETMS.au3"

#include "src\features\comat\Action_COMAT.au3"
#include "src\features\comat\Batch_COMAT.au3"

#include "src\features\fc\Action_FC.au3"
#include "src\features\fc\Batch_FC.au3"
#include "src\features\fc\Audit_FC.au3"

#include "src\features\mail\Mail_Common.au3"
#include "src\features\mail\Mail_RDV.au3"
#include "src\features\mail\Mail_Alertes.au3"
#include "src\features\mail\Mail_CP.au3"

#include "src\features\edoc\Action_EDOC.au3"
#include "src\features\edoc\EDOC_Web.au3"
#include "src\features\edoc\EDOC_Master.au3"

#include "src\features\cmr\CMR_Engine.au3"
#include "src\features\cmr\CMR_Mail.au3"
#include "src\features\cmr\CMR_EDOC.au3"
#include "src\features\cmr\CMR_PDF.au3"
#include "src\features\cmr\Action_CMR.au3"

; Diagnostics
#include "src\features\diagnostics\Diagnostic.au3"
#include "src\features\diagnostics\StorageInfo.au3"
#include "src\features\diagnostics\ContactsCleanup.au3"

; API et serveur, en dernier
#include "src\api\ApiPing.au3"
#include "src\api\ApiState.au3"
#include "src\api\ApiContacts.au3"
#include "src\api\ApiNetwork.au3"
#include "src\api\ApiConfig.au3"
#include "src\api\ApiJobs.au3"
#include "src\api\ApiActions.au3"

#include "src\core\HttpResponse.au3"
#include "src\core\HttpRouter.au3"
#include "src\core\HttpServer.au3"

MainDispatch()

Func MainDispatch()
    DispatchInitialize()
    DispatchStartServer()
    DispatchOpenInterface()

    While 1
        DispatchProcessClients()
        DispatchProcessBackgroundJobs()
        Sleep(10)
    WEnd
EndFunc
```

---

## 5. Répartition détaillée de `Dispatch.au3`

| Code actuel | Nouveau fichier |
|---|---|
| Tous les `Global` | `src/core/Globals.au3` |
| Constantes délais, limites, classes Windows | `src/core/Constants.au3` |
| Démarrage, port TCP, ouverture HTML | `src/core/HttpServer.au3` |
| Réception des headers/body HTTP | `src/core/HttpServer.au3` |
| `SendHttpResponse` | `src/core/HttpResponse.au3` |
| `ElseIf $sURL = ...` | `src/core/HttpRouter.au3` et `src/api/` |
| `JsonEscape`, `GetJsonValue`, `GetJsonArrayValue` | `src/core/JsonUtils.au3` |
| `BackupRotate` | `src/core/BackupUtils.au3` |
| `ValidateNetPath`, `ValidateId`, `SanitizeString` | `src/core/Security.au3` |
| `AuditLog`, health check et logs | `src/core/Audit.au3` |
| Lecture/écriture des fichiers JSON | `src/services/StateService.au3` |
| Lecture/écriture contacts TSV/JSON | `src/services/ContactsService.au3` |
| Lecture/écriture de `config.ini` | `src/services/ConfigService.au3` |
| Sauvegarde/lecture réseau F: | `src/services/NetworkStateService.au3` |
| Suivi des jobs ETMS | `src/services/JobService.au3` |
| `/api/ping` | `src/api/ApiPing.au3` |
| `/api/load`, `/api/save`, `/api/load-data`, `/api/save-data`, patches | `src/api/ApiState.au3` |
| `/api/load-contacts`, `/api/save-contacts` | `src/api/ApiContacts.au3` |
| `/api/net-save`, `/api/net-load`, `/api/net-list`, `/api/net-check` | `src/api/ApiNetwork.au3` |
| `/api/save-config`, `/api/load-bkd-config`, `/api/save-pj-config`, `/api/save-cp-config` | `src/api/ApiConfig.au3` |
| `/api/job-status` | `src/api/ApiJobs.au3` |
| `Switch $sAction` | `src/api/ApiActions.au3` |
| `ETMSCMD` | `src/features/etms/Action_ETMS.au3` |
| `COMATMULTI`, `COMATSOLO`, pause/stop/skip | `src/features/comat/Action_COMAT.au3` |
| Batch COMAT | `src/features/comat/Batch_COMAT.au3` |
| `KANBAN5` et actions FC | `src/features/fc/Action_FC.au3` |
| Batch FC | `src/features/fc/Batch_FC.au3` |
| Audit FC | `src/features/fc/Audit_FC.au3` |
| `MAILRDV` | `src/features/mail/Mail_RDV.au3` |
| Alertes et actions mails | `src/features/mail/Mail_Alertes.au3` |
| `BATCHCP` | `src/features/mail/Mail_CP.au3` |
| Fonctions mail communes | `src/features/mail/Mail_Common.au3` |
| `CMRGENERATE`, `CMRSTATUS`, `CMRSTOP`, `CMRFORCE` | `src/features/cmr/Action_CMR.au3` |
| Bloc CMR complet | `src/features/cmr/CMR_Engine.au3` |
| Fonctions mail CMR | `src/features/cmr/CMR_Mail.au3` |
| Fonctions EDOC CMR | `src/features/cmr/CMR_EDOC.au3` |
| Reconstruction PDF CMR | `src/features/cmr/CMR_PDF.au3` |
| Actions EDOC | `src/features/edoc/Action_EDOC.au3` |
| EDOC Web | `src/features/edoc/EDOC_Web.au3` |
| EDOC Master GUI | `src/features/edoc/EDOC_Master.au3` |
| `DIAG` | `src/features/diagnostics/Diagnostic.au3` |
| `STORAGEINFO` | `src/features/diagnostics/StorageInfo.au3` |
| `CLEANCONTACTS` | `src/features/diagnostics/ContactsCleanup.au3` |
| Bloc HPE BL Queue Snapshot | `src/features/cmr/CMR_Engine.au3` au départ |

---

## 6. Règles pour `Globals.au3`

```autoit
; src/core/Globals.au3

Global $giTrackCount = 0
Global $gaTrackIDs[1]
Global $ghTracker = 0

Global $bFCStop = False
Global $bFCPause = False
Global $bFCSkip = False
Global $iFCStepCurrent = 0
Global $gsFCAuditLog = ""
Global $gbFCAudit = True

Global $bCOMATStop = False
Global $bCOMATPause = False
Global $bCOMATSkip = False

Global $giPort = 9500
Global $giMainSocket = -1

Global $gsHtmlFile = @ScriptDir & "\ui\index.html"
Global $gsSaveFile = @ScriptDir & "\dispatch.json"
Global $gsDataFile = @ScriptDir & "\data.json"
Global $gsStatusFile = @ScriptDir & "\status.json"
Global $gsContactsFile = @ScriptDir & "\contacts.tsv"
Global $gsContactsLegacyFile = @ScriptDir & "\contacts.json"
Global $gsConfigFile = @ScriptDir & "\config.ini"

Global $gsCMRStatus = "Prêt"
Global $gbCMRRunning = False
Global $gbCMRStop = False
Global $gbCMRForce = False
```

Avant chaque déplacement, rechercher les globals existants dans `Dispatch.au3` et les ajouter une seule fois dans ce fichier.

---

## 7. Règles spécifiques au module CMR

`Dispatch_CMR_STABLE_FULL.au3` ne doit pas être inclus tel quel si le fichier contient :

- une fonction `Main()` ;
- une deuxième boucle `While 1` ;
- des `Opt(...)` de démarrage ;
- des globals déjà définis dans `Dispatch.au3` ;
- une deuxième initialisation du serveur.

Il doit devenir un module de fonctions :

```autoit
; src/features/cmr/CMR_Engine.au3

Func CMRSTABLERunFromData($sData)
    ; logique CMR existante
EndFunc
```

Le seul `Main()` doit rester dans `MainDispatch.au3`.

---

## 8. Architecture HTML finale

Le fichier `ui/index.html` ne contient plus le CSS ni la logique JavaScript. Il contient :

- le squelette HTML ;
- les onglets ;
- les tables ;
- les conteneurs Kanban ;
- les modales ;
- les boutons ;
- les inputs ;
- les zones de statut ;
- les liens CSS ;
- les balises `<script>`.

```html
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>DispatchMaster</title>

  <link rel="stylesheet" href="/css/00-reset.css">
  <link rel="stylesheet" href="/css/01-variables.css">
  <link rel="stylesheet" href="/css/02-layout.css">
  <link rel="stylesheet" href="/css/03-components.css">
  <link rel="stylesheet" href="/css/04-table.css">
  <link rel="stylesheet" href="/css/05-kanban.css">
  <link rel="stylesheet" href="/css/06-modals.css">
  <link rel="stylesheet" href="/css/07-options.css">
  <link rel="stylesheet" href="/css/08-dashboard.css">
  <link rel="stylesheet" href="/css/09-cmr.css">
  <link rel="stylesheet" href="/css/10-responsive.css">
</head>
<body>
  <!-- Reprendre ici les blocs HTML de Interface-2.html sans les styles inline si possible. -->

  <script src="/js/00-bootstrap.js"></script>
  <script src="/js/01-state.js"></script>
  <script src="/js/02-utils.js"></script>
  <script src="/js/03-api.js"></script>
  <script src="/js/04-storage.js"></script>
  <script src="/js/05-contacts.js"></script>
  <script src="/js/06-dispatch-table.js"></script>
  <script src="/js/07-kanban.js"></script>
  <script src="/js/08-groups.js"></script>
  <script src="/js/09-workflow.js"></script>
  <script src="/js/10-actions-etms.js"></script>
  <script src="/js/11-actions-comat.js"></script>
  <script src="/js/12-actions-fc.js"></script>
  <script src="/js/13-actions-cmr.js"></script>
  <script src="/js/14-actions-edoc.js"></script>
  <script src="/js/15-mail.js"></script>
  <script src="/js/16-nto.js"></script>
  <script src="/js/17-options.js"></script>
  <script src="/js/18-dashboard.js"></script>
  <script src="/js/19-modals.js"></script>
  <script src="/js/20-undo.js"></script>
  <script src="/js/21-sync.js"></script>
  <script src="/js/22-events.js"></script>
  <script src="/js/99-init.js"></script>
</body>
</html>
```

---

## 9. Répartition précise du CSS

| Bloc actuel dans `Interface-2.html` | Nouveau fichier |
|---|---|
| `:root`, couleurs, variables CSS | `01-variables.css` |
| `box-sizing`, `html`, `body`, reset | `00-reset.css` |
| header, tabs, layout général | `02-layout.css` |
| boutons, badges, inputs, labels, toast | `03-components.css` |
| table Dispatch et cellules | `04-table.css` |
| Kanban, cartes, colonnes | `05-kanban.css` |
| overlays, modales, spinner | `06-modals.css` |
| Options, NTO, Channel Partners, BKD | `07-options.css` |
| Dashboard et classes `cr-*` | `08-dashboard.css` |
| CMR et queue CMR | `09-cmr.css` |
| media queries | `10-responsive.css` |

Lors du premier split, ne pas renommer les classes CSS. Il faut seulement déplacer le contenu.

---

## 10. Répartition précise du JavaScript

### `00-bootstrap.js`

Variables globales et constantes :

```javascript
"use strict";

var gmaster = [];
var grawData = {};
var gcpData = [];
var gcontacts = [];
var gcpConfig = [];

var goperatorFilter = "";
var goperator = "";
var geditIdx = -1;
var ggroupIdx = -1;

var APIURL = window.location.origin;
var dirty = false;
var saveTimer = null;
var undoStack = [];

const UNDOMAX = 5;
const SAVEDELAY = 1500;
```

### `01-state.js`

Déplacer :

- `gmaster` et ses fonctions de manipulation ;
- `grawData` ;
- `gcpData` ;
- labels et couleurs de statuts ;
- `statutNum()` ;
- `statutLabel()` ;
- recherche de dossiers ;
- `markDirty()` ;
- `stampRecord()`.

### `02-utils.js`

Déplacer :

- échappement HTML ;
- échappement d'attributs ;
- formatage dates/nombres ;
- helpers DOM ;
- `toast()` ;
- `calcTaxable()` ;
- `calcTransp()` ;
- fonctions générales de tableau.

### `03-api.js`

Déplacer :

- `apiCall()` ;
- `sendAction()` ;
- `sendActionHeader()` ;
- `netSave()` ;
- `netLoad()` ;
- `netListFiles()` ;
- tous les appels `fetch()` vers AutoIt.

### `04-storage.js`

Déplacer :

- IndexedDB ;
- `idbOpen()` ;
- `idbPut()` ;
- `idbGet()` ;
- `idbSave()` ;
- `idbLoad()` ;
- `autoSave()` ;
- `saveContactsChunked()` ;
- import/export JSON.

### `05-contacts.js`

Déplacer :

- affichage des contacts ;
- recherche ;
- ajout/modification/suppression ;
- remplissage automatique ;
- sauvegarde TSV/JSON ;
- configuration des Channel Partners si liée aux contacts.

### `06-dispatch-table.js`

Déplacer :

- `renderMaster()` ;
- génération du tableau ;
- recherche ;
- filtres ;
- édition des lignes ;
- sélection ;
- suppression des terminés ;
- statistiques de la table.

### `07-kanban.js`

Déplacer :

- `renderKanban()` ;
- création des cartes ;
- déplacement ;
- changement de statut ;
- synchronisation table/Kanban ;
- mise à jour depuis une carte.

### `08-groups.js`

Déplacer :

- `openGroupManual()` ;
- `grpmParseFiles()` ;
- `grpmPreview()` ;
- `grpmAutoTransp()` ;
- `grpmSave()` ;
- ajout manuel ;
- retrait partiel de groupe ;
- validation CC partiel ;
- séparation solo/groupe.

### `09-workflow.js`

Déplacer :

- export workflow ;
- import workflow ;
- fusion ;
- export JSON complet ;
- import JSON ;
- export Excel ;
- fusion des dossiers, contacts et CP.

### `10-actions-etms.js`

Déplacer :

- commandes ETMS ;
- boutons ETMS ;
- spinner ;
- ouverture du dossier ;
- actions `ETMSCMD`.

### `11-actions-comat.js`

Déplacer :

- COMAT solo ;
- COMAT multi ;
- traitement séquentiel ;
- pause/stop/skip ;
- barre de progression ;
- `comatStartBatch()` ;
- `comatProcessNext()`.

### `12-actions-fc.js`

Déplacer :

- calcul des dates ;
- règles Flex/UPS/autres ;
- cutoff 14 h 30 ;
- BKD ;
- DLY ;
- modale FC ;
- actions FC ;
- traitement séquentiel ;
- barre de progression FC.

### `13-actions-cmr.js`

Déplacer :

- saisie manuelle des BL ;
- séparation par ligne vide ou `---` ;
- ajout depuis Kanban ;
- génération CMR ;
- statut ;
- arrêt/forçage ;
- affichage de la queue.

### `14-actions-edoc.js`

Déplacer :

- initialisation EDOC ;
- scan ;
- upload ;
- règles ;
- EDOC Master ;
- actions EDOC CMR.

### `15-mail.js`

Déplacer :

- mail RDV ;
- mail alertes ;
- mail CP ;
- mail CMR ;
- sélection des dossiers ;
- construction des payloads.

### `16-nto.js`

Déplacer :

- `NTODGS` ;
- `NTOUTE` ;
- `NTOFUEL` ;
- tranches ;
- calcul DGS ;
- calcul UTE ;
- fuel ;
- affichage ;
- options tarifaires.

### `17-options.js`

Déplacer :

- fuel ;
- Channel Partners ;
- BKD global ;
- réseau ;
- opérateur ;
- thème sombre ;
- paramètres PJ.

### `18-dashboard.js`

Déplacer :

- création du Dashboard ;
- KPIs ;
- alertes ;
- filtres ;
- répartition par statut/opérateur/transporteur ;
- panneau des informations manquantes.

### `19-modals.js`

Déplacer :

- `openModal()` ;
- `closeModal()` ;
- fermeture par Escape ;
- modale identité ;
- modale édition ;
- modale groupe ;
- modale FC ;
- modale NTO ;
- modale paramètres.

### `20-undo.js`

Déplacer :

- `pushUndo()` ;
- `undoLast()` ;
- limitation aux cinq actions ;
- snapshots de `gmaster` et `gcpData`.

### `21-sync.js`

Déplacer :

- `refreshDataSilent()` ;
- synchronisation toutes les deux minutes ;
- chargement serveur ;
- réconciliation IndexedDB/serveur ;
- indicateur hors ligne ;
- fusion selon `ts`.

### `22-events.js`

Déplacer :

- listeners des boutons ;
- listeners des tabs ;
- listeners des inputs ;
- raccourcis clavier ;
- écoute globale `markDirty()` ;
- événements de navigation.

### `99-init.js`

```javascript
window.addEventListener("load", async function () {
  await initialiseIndexedDB();
  await chargerDonneesInitiales();

  renderMaster();
  renderKanban();
  renderContacts();
  updateStats();
  updateDashboard();
  populateOperatorFilter();
  bindAllEvents();

  setInterval(function () {
    refreshData(true);
  }, 120000);
});
```

---

## 11. Adaptation du serveur AutoIt pour CSS et JS

Le serveur actuel doit pouvoir servir tous les fichiers de `ui/css/` et `ui/js/`.

Ajouter un routage équivalent à :

```autoit
If StringLeft($sURL, 5) = "/css/" Then
    ServeStaticFile($iSocket, @ScriptDir & "\ui" & $sURL, "text/css")
    Return
EndIf

If StringLeft($sURL, 4) = "/js/" Then
    ServeStaticFile($iSocket, @ScriptDir & "\ui" & $sURL, "application/javascript")
    Return
EndIf

If $sURL = "/" Or $sURL = "/index.html" Then
    ServeStaticFile($iSocket, @ScriptDir & "\ui\index.html", "text/html")
    Return
EndIf
```

Fonction commune :

```autoit
Func ServeStaticFile($iSocket, $sPath, $sType)
    If StringInStr($sPath, "..") Then
        SendHttpResponse($iSocket, 403, "text/plain", "Forbidden")
        Return
    EndIf

    If Not FileExists($sPath) Then
        SendHttpResponse($iSocket, 404, "text/plain", "File not found")
        Return
    EndIf

    Local $hFile = FileOpen($sPath, 256)
    If $hFile = -1 Then
        SendHttpResponse($iSocket, 500, "text/plain", "Cannot read file")
        Return
    EndIf

    Local $sContent = FileRead($hFile)
    FileClose($hFile)

    SendHttpResponse($iSocket, 200, $sType, $sContent)
EndFunc
```

La validation `..` est obligatoire pour empêcher qu'une URL puisse demander un fichier situé en dehors du dossier `ui`.

---

## 12. Procédure de migration sans casse

### Étape 0 : sauvegarde

Créer une copie complète :

```text
Dispatch_Legacy_Backup/
├── Dispatch.au3
├── Interface-2.html
└── tous les fichiers actuels
```

Ne jamais modifier directement la version opérationnelle sans copie.

### Étape 1 : externaliser uniquement le CSS

1. Copier `Interface-2.html` vers `Interface-2.monolithique.html`.
2. Extraire le bloc `<style>` dans `ui/css/all.css`.
3. Remplacer le bloc `<style>` par :

```html
<link rel="stylesheet" href="/css/all.css">
```

4. Modifier le serveur AutoIt pour servir `/css/all.css`.
5. Tester toute l'interface.

### Étape 2 : externaliser uniquement le JavaScript

1. Extraire tous les blocs `<script>` dans `ui/js/all.js`.
2. Remplacer les scripts inline par :

```html
<script src="/js/all.js"></script>
```

3. Tester le démarrage, le chargement des données, les actions ETMS, FC, COMAT, CMR, EDOC et les sauvegardes.

### Étape 3 : créer les fichiers JavaScript fonctionnels

Copier les sections de `all.js` dans les fichiers fonctionnels, sans changer les noms de fonctions.

Ordre recommandé :

1. `00-bootstrap.js`
2. `01-state.js`
3. `02-utils.js`
4. `03-api.js`
5. `04-storage.js`
6. `05-contacts.js`
7. `06-dispatch-table.js`
8. `07-kanban.js`
9. `08-groups.js`
10. `09-workflow.js`
11. `10-actions-etms.js`
12. `11-actions-comat.js`
13. `12-actions-fc.js`
14. `13-actions-cmr.js`
15. `14-actions-edoc.js`
16. `15-mail.js`
17. `16-nto.js`
18. `17-options.js`
19. `18-dashboard.js`
20. `19-modals.js`
21. `20-undo.js`
22. `21-sync.js`
23. `22-events.js`
24. `99-init.js`

### Étape 4 : créer `MainDispatch.au3`

1. Copier `Dispatch.au3` vers `MainDispatch.au3`.
2. Déplacer les globals dans `Globals.au3`.
3. Déplacer les fonctions utilitaires.
4. Déplacer les services de stockage.
5. Déplacer les fonctionnalités une par une.
6. Déplacer les routes API.
7. Garder le serveur et la boucle principale pour la fin.
8. Compiler uniquement `MainDispatch.au3`.

### Étape 5 : validation fonctionnelle

Tester dans cet ordre :

1. démarrage du serveur ;
2. ouverture de l'interface ;
3. `/api/ping` ;
4. chargement initial ;
5. sauvegarde `status.json` ;
6. sauvegarde `data.json` ;
7. contacts ;
8. réseau F: ;
9. édition d'un dossier ;
10. Kanban ;
11. COMAT solo ;
12. COMAT multi ;
13. FC ;
14. mails ;
15. CMR ;
16. EDOC ;
17. NTO ;
18. import/export ;
19. synchronisation ;
20. diagnostic.

---

## 13. Checklist anti-régression

### AutoIt

- [ ] Un seul `Main()`.
- [ ] Une seule boucle principale.
- [ ] Un seul démarrage TCP.
- [ ] Un seul `Globals.au3`.
- [ ] Aucun global redéfini dans les modules.
- [ ] Les chemins utilisent toujours `@ScriptDir`.
- [ ] `index.html` est trouvé depuis le dossier de l'exécutable.
- [ ] Les fichiers JSON restent accessibles.
- [ ] Les sauvegardes sont conservées.
- [ ] Les actions longues répondent toujours rapidement au navigateur.
- [ ] Les chemins réseau sont toujours validés.
- [ ] Le body HTTP reste limité par `MAXBODYSIZE`.

### JavaScript

- [ ] `gmaster` n'est déclaré qu'une fois.
- [ ] `grawData` n'est déclaré qu'une fois.
- [ ] `gcpData` n'est déclaré qu'une fois.
- [ ] `gcontacts` n'est déclaré qu'une fois.
- [ ] `apiCall()` est chargé avant les modules d'actions.
- [ ] Les fonctions de state sont chargées avant les fonctions de rendu.
- [ ] Les fonctions de rendu sont chargées avant les événements.
- [ ] `99-init.js` est chargé en dernier.
- [ ] Aucun `import/export` n'est utilisé.
- [ ] IndexedDB reste optionnel avec fallback API.
- [ ] Les appels API utilisent les mêmes noms d'actions.

### HTML/CSS

- [ ] Les IDs HTML existants sont conservés.
- [ ] Les classes CSS existantes sont conservées.
- [ ] Les modales gardent leurs IDs.
- [ ] Les boutons gardent leurs IDs.
- [ ] Les scripts CSS/JS sont chargés avec des chemins absolus `/css/` et `/js/`.
- [ ] Les fichiers CSS et JS sont servis par AutoIt.

---

## 14. Version de production finale

Après validation, le dossier à transmettre à l'utilisateur doit ressembler à ceci :

```text
Dispatch/
├── Dispatch.exe
├── MainDispatch.au3                 ; facultatif en production
├── ui/
│   ├── index.html
│   ├── css/
│   └── js/
├── dispatch.json
├── data.json
├── status.json
├── contacts.tsv
├── contacts.json
├── config.ini
├── audit.log
└── backups/
```

En développement, conserver tout le dossier `src/`. En production, `Dispatch.exe` n'a besoin que des fichiers UI et des fichiers de données réellement utilisés.

---

## 15. Résultat attendu

Le workflow final doit être :

```text
SciTE
  ↓ F5
MainDispatch.au3
  ↓
Serveur HTTP AutoIt local
  ↓
ui/index.html
  ↓
CSS + JavaScript séparés
  ↓
API AutoIt
  ↓
JSON / TSV / INI / réseau F:
```

L'utilisateur final doit seulement faire :

```text
Double-clic sur Dispatch.exe
```

Le découpage ne fragmente pas la mémoire. Il sépare uniquement les responsabilités :

- AutoIt garde les fichiers et services communs ;
- le JavaScript garde un seul état `gmaster` ;
- les fonctionnalités deviennent des fichiers indépendants ;
- l'API reste le lien entre l'interface et AutoIt ;
- le comportement actuel est conservé.

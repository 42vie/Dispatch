# Dispatch - Version Finale

## Structure finale

```
Dispatch/
├── MainDispatch.au3      ← Source AutoIt (F5 dans SciTE)
├── Dispatch.exe          ← Executable compil (àīī gnrer avec BUILD.bat)
├── BUILD.bat             ← Script de compilation automatique
│
├── src/
│   ├── core/             ← Globals, HTTP server, utils
│   ├── api/              ← Endpoints API (/api/ping, /api/action...)
│   ├── services/         ← Services mtier (State, Contacts, Network...)
│   └── features/         ← Features (ETMS, COMAT, FC, CMR, EDOC, Mail...)
│
├── ui/
│   ├── index.html        ← Interface web
│   ├── css/              ← Styles CSS
│   └── js/               ← JavaScript client
│
└── Fichiers de donnes
    ├── dispatch.json
    ├── data.json
    ├── status.json
    ├── contacts.tsv
    └── config.ini
```

## Utiliser en production

### Mthode 1: Script automatique (recommand)

1. Double-cliquer sur `BUILD.bat`
2. Attre que `Dispatch.exe` soit gnr
3. Copier le dossier complet avec `Dispatch.exe` et `ui/`
4. Lancer `Dispatch.exe`

### Mthode 2: SciTE (dveloppement)

1. Ouvrir `MainDispatch.au3` dans SciTE
2. F5 pour dmarrer le serveur
3. L'interface s'ouvre sur `http://localhost:9500`

### Mthode 3: Compilation manuelle

1. Dans SciTE: `Tools` → `Build` sur `MainDispatch.au3`
2. Gnr `Dispatch.exe`
3. Copier avec `ui/` et fichiers `.json/.tsv/.ini`

## Fonctionnalits

- ✅ Serveur HTTP local (port 9500)
- ✅ Interface web responsive
- ✅ API REST complte
- ✅ Gestion des dossiers (Dispatch)
- ✅ COMAT batch
- ✅ File Closing (FC)
- ✅ CMR (CMR BL)
- ✅ EDOC upload
- ✅ Mails (RDV, alertes, CP)
- ✅ Contacts + Channel Partners
- ✅ NTO (tarification)
- ✅ Synchronisation rseau F:
- ✅ IndexedDB local
- ✅ Thme sombre

## API Endpoints

- `GET /api/ping` - Health check
- `POST /api/action` - Actions mtier (ETMSCMD, COMATMULTI, KANBAN5, etc.)
- `GET /api/load` - Charger donnes
- `POST /api/save` - Sauvegarder donnes
- `GET /api/load-contacts` - Charger contacts
- `POST /api/save-contacts` - Sauvegarder contacts
- `POST /api/net-save` - Sauvegarde rseau
- `GET /api/net-load` - Chargement rseau
- `GET /api/net-list` - Liste fichiers rseau
- `GET /api/net-check` - Vrification rseau
- `POST /api/job-status` - Statut jobs

## Dpannage

### Le serveur ne dmarre pas

- Vrifier que le port 9500 n'est pas utilis
- Vrifier que `config.ini` est prsent
- Regarder `audit.log` pour les erreurs

### L'interface ne s'ouvre pas

- Vrifier que `ui/index.html` existe
- Ouvrir manuellement `http://localhost:9500`

### Erreurs API

- Tester `/api/ping` dans le navigateur
- Vrifier les logs dans `audit.log`
- Vrifier les permissions sur les fichiers JSON

## Prochaines tapes

1. Compiler `Dispatch.exe` avec `BUILD.bat`
2. Tester toutes les fonctionnalits
3. Dployer avec `ui/` et donnes
4. Former les utilisateurs

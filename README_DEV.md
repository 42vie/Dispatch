# Dispatch - Guide de développement

## Architecture

```
Dispatch/
├── MainDispatch.au3          ← Point d'unique (F5 dans SciTE)
├── Dispatch.exe              ← Version compilé§§e (prod)
│
├── src/
│   ├── core/                 ← Globals, HTTP, utils
│   ├── api/                  ← Endpoints API
│   ├── services/             ← Services métier
│   └── features/             ← Features (ETMS, COMAT, FC, CMR, EDOC...)
│
├── ui/
│   ├── index.html            ← Interface
│   ├── css/                  ← Styles
│   └── js/                   ← Logique client
│
└── Fichiers données
    ├── dispatch.json
    ├── data.json
    ├── status.json
    ├── contacts.tsv
    └── config.ini
```

## Développement

### Lancer en dev

1. Ouvrir `MainDispatch.au3` dans SciTE
2. F5 pour démarrer
3. Le serveur HTTP démarre sur `http://localhost:9500`
4. L'interface s'ouvre automatiquement

### Compiler pour prod

1. Dans SciTE, `Tools` → `Build` sur `MainDispatch.au3`
2. Ça génè§§re `Dispatch.exe`
3. Copier avec `ui/` et fichiers `.json/.tsv/.ini`

### Structure des modules

Chaque module AutoIt doit :
- Inclure `Globals.au3` et `Constants.au3`
- Ne pas redé§§finir de variables globales
- Exposer des fonctions claires

Exemple : `Action_COMAT.au3` expose `Action_COMAT_Run()`

## API

Endpoints disponibles :

- `GET /api/ping` - Health check
- `POST /api/action` - Actions métier
- `GET /api/load` - Charger données
- `POST /api/save` - Sauvegarder
- `GET /api/load-contacts` - Charger contacts
- etc.

## Tests

Lancer `TEST_SERVER.au3` pour vérifier le serveur.

## Dépannage

### Le serveur ne démarre pas
- Vérifier port 9500 non utilisé
- Vérifier `config.ini` présent

### L'interface ne s'ouvre pas
- Vérifier `ui/index.html` présent
- Vérifier chemin dans `Globals.au3`

### Erreurs API
- Vérifier logs dans `audit.log`
- Tester `/api/ping` dans navigateur

## Prochaines étapes

1. Complé§§ter les modules manquants
2. Tester chaque feature
3. Compiler `Dispatch.exe`
4. Déployer avec `ui/` et données

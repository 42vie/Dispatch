# Dispatch

Dispatch est un outil d'automatisation développé en AutoIt, orchestré·¹ une interface HTML et un système de blocs d'action configurables.

## Architecture

- **Core AutoIt** :
  - `Dispatch.au3` : point d'entré·¹ e / orchestrateur principal.
  - `MainDispatch.au3` : logique métier principale.
  - `cmr.au3`, `Robot EDOC V2.au3` : modules spécialisé.1;s (CMR, robot EDOC).
- **Interface** :
  - `Interface.html` : front-end HTML.
  - `INTERFACE_BLOCKS_MAP.json` : mapping entre UI et blocs d'action.
  - Dossier `ui/` : ressources UI complémentaires.
- **Blocs d'action** :
  - Dossier `action_blocks/` : implé·¹ mentation des différentes actions automatisé·¹ es.
- **Configuration & données** :
  - `config.ini` : configuration générale.
  - `MANIFEST.json` : métadonné·¹ es / description des composants.
  - `contacts.json`, `contacts.tsv`, `data.json`, `dispatch.json`, `status.json` : données utilisées par l'outil.
  - `RECONSTRUCTION_CHECK.json` : checks / validation de reconstruction.
- **Scripts utilitaires** :
  - `BUILD_Dispatch.bat` : script de build / compilation.
  - `Dispatch.exe.NOT_COMPILED.txt` : note indiquant que l'exé·¹ cutable n'est pas compilé·¹ dans le repo.

## Installation & build

1. Cloner le repo :
   ```bash
   git clone https://github.com/42vie/Dispatch.git
   cd Dispatch
   ```
2. Avoir AutoIt installé sur ta machine Windows.
3. Compiler le projet :
   - Utiliser `BUILD_Dispatch.bat` ou ton outil de compilation AutoIt habituel.
   - Le fichier `Dispatch.exe.NOT_COMPILED.txt` rappelle que l'exé·¹ cutable n'est pas commité·¹ .

## Utilisation

- Lancer `Dispatch.exe` (aprè·¹ s compilation) ou exé·¹ cuter `Dispatch.au3` via AutoIt.
- L'interface HTML (`Interface.html`) sert de front-end pour dé• 1;clencher et superviser les actions.
- Les paramè·¹ tres et comportements sont configuré·¹ s via `config.ini` et les fichiers JSON.

## Structure des dossiers

- `action_blocks/` : chaque bloc implé·¹ mente une action spécifique appelé.1;e par le core.
- `src/` : code source complémentaire (modules, helpers).
- `ui/` : assets et composants UI.

## Logs & audit

- `audit.log` : journal d'audit des actions exé·¹ cuté·¹ es.

## Licence & contact

Projet interne / privé.
Pour toute question, ouvrir une issue ou contacter les mainteneurs.
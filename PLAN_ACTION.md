# Plan d'action – Dispatch

## Objectif

Amé·¹ liorer la maintenabilité·¹ , la lisibilité·¹ et l'extensibilité·¹ de Dispatch tout en gardant une chaî.1;ne de build simple.

## Axes de travail

### 1. Documentation & onboarding

- [ ] Finaliser le README avec :
  - description claire du projet,
  - sché·¹ ma d'architecture (core / UI / action_blocks),
  - instructions de build et d'usage.
- [ ] Ajouter un `CONTRIBUTING.md` (conventions de code, branches, PRs).
- [ ] Documenter chaque bloc d'action dans `action_blocks/` (README par bloc ou commentaires standardisé·¹ s).

### 2. Architecture & refactor

- [ ] Clarifier les responsabilité.1;s :
  - `Dispatch.au3` : bootstrap + wiring.
  - `MainDispatch.au3` : logique métier centrale.
  - `cmr.au3`, `Robot EDOC V2.au3` : domaines spécifiques.
- [ ] Centraliser la configuration :
  - s'assurer que `config.ini` et les JSON (`MANIFEST.json`, `INTERFACE_BLOCKS_MAP.json`, etc.) sont cohé·¹ rents et non redondants.
- [ ] Nettoyer les fichiers inutilisé·¹ s ou obsolè·¹ tes.

### 3. Interface & mapping

- [ ] Mettre à jour `INTERFACE_BLOCKS_MAP.json` pour reflé·¹ chir exactement les blocs disponibles.
- [ ] Vé.1;rifier la synchronisation entre :
  - les blocs dans `action_blocks/`,
  - le mapping JSON,
  - l'interface HTML.
- [ ] Ajouter des tests manuels / checklists pour chaque action visible dans l'UI.

### 4. Données & intégrité·¹

- [ ] Valider le format et l'usage de :
  - `contacts.json` / `contacts.tsv`,
  - `data.json`,
  - `dispatch.json`,
  - `status.json`,
  - `RECONSTRUCTION_CHECK.json`.
- [ ] Mettre en place des contrô.1;les de cohé·¹ rence (via `RECONSTRUCTION_CHECK.json` ou script dédié).

### 5. Build & distribution

- [ ] Clarifier le processus de build :
  - rôle exact de `BUILD_Dispatch.bat`,
  - environnement requis (version AutoIt, outils).
- [ ] Documenter comment gé.1;né·¹ rer `Dispatch.exe` proprement.
- [ ] Ajouter un mode "dev" vs "prod" si pertinent.

### 6. Logs, audit & debug

- [ ] Standardiser le format de `audit.log`.
- [ ] Ajouter des niveaux de log (INFO, WARN, ERROR).
- [ ] Pré.1;voir un mé.1;canisme de rotation / archivage des logs.

## Prochaines étapes immédiates

1. Valider ce plan d'action.
2. Prioriser les tâches (ex : documentation → refactor → interface → build).
3. Créer des issues GitHub correspondant à chaque point ci-dessus.
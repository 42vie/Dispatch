# Plan d’action — Dispatch

## Objectif
Modulariser progressivement Dispatch sans modifier son comportement : conserver les données, les endpoints API, les variables globales, l’UX et la compatibilité AutoIt/SciTE, sans Node.js ni framework externe.

## État initial — 2026-08-06
- [x] README/plan de modularisation analysé.
- [x] Dépôt identifié : `42vie/Dispatch`.
- [ ] Sauvegarde logique de la version monolithique.
- [ ] Inventaire détaillé des globals, endpoints et fonctions.
- [ ] Extraction du CSS.
- [ ] Extraction du JavaScript.
- [ ] Découpage du JavaScript en scripts classiques ordonnés.
- [ ] Découpage AutoIt en modules.
- [ ] Création de `MainDispatch.au3` comme point d’entrée unique.
- [ ] Routage sécurisé des fichiers statiques `/css/` et `/js/`.
- [ ] Validation fonctionnelle et anti-régression.

## Règles
- Un seul `Main()` et une seule boucle serveur.
- Un seul fichier de globals.
- Aucun `import`/`export` JavaScript pendant la première migration.
- Conservation des endpoints `/api/*` existants.
- Conservation des IDs/classes HTML et des noms de fonctions.
- Validation de sécurité des chemins statiques et conservation des sauvegardes.

## Journal des changements
### 2026-08-06 — Initialisation
- Création de ce fichier de suivi.
- Prochaine étape : inventorier précisément le code avant toute extraction.

## Validation
- [ ] Compilation AutoIt réussie.
- [ ] Démarrage du serveur et ouverture de l’interface.
- [ ] Vérification des API, données, contacts, réseau, Kanban, actions métier, CMR, EDOC, NTO, import/export et synchronisation.

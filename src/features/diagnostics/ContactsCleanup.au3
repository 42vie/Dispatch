; ============================================================================
; ContactsCleanup.au3
; Nettoyage contacts.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

; Aucun bloc dedie trouve automatiquement dans Dispatch.txt.
; Fichier cree pour respecter l architecture cible et recevoir la migration.
;
; La logique de nettoyage des contacts existe deja et fonctionne :
; voir _CleanContactsFiles() dans src/services/ContactsService.au3 (ligne 9).
; Elle n'a pas ete deplacee ici pour eviter tout risque de regression
; sur un chemin de code deja en production.

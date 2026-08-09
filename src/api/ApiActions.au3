; ============================================================================
; ApiActions.au3
; Routes d actions.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

; Aucun bloc dedie trouve automatiquement dans Dispatch.txt.
; Fichier cree pour respecter l architecture cible et recevoir la migration.
;
; La route /api/action est geree par le grand Switch $sAction dans
; src/core/HttpRouter.au3, avec des dizaines de Case (CMR_*, FC_*, EDOC_*,
; SP_*, HPE_*, COMAT_*, ...) qui partagent des variables Local declarees
; une seule fois avant le Switch. Extraire ces Case vers ce fichier casserait
; ce partage et risquerait de regresser un routeur HTTP deja en production.
; Laisse intentionnellement vide -- voir HttpRouter.au3 pour la logique reelle.

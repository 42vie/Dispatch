; ============================================================================
; Action_CMR.au3
; Actions CMR API/UI.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

; Aucun bloc dedie trouve automatiquement dans Dispatch.txt.
; Fichier cree pour respecter l architecture cible et recevoir la migration.
;
; Les actions CMR (CMR_RUN, CMR_STATUS, CMR_STOP, SP_*, ...) vivent
; actuellement comme des Case du grand Switch $sAction dans
; src/core/HttpRouter.au3 (route /api/action), avec la logique metier dans
; src/features/cmr/CMR_Engine.au3, CMR_Results.au3, CMR_SP.au3, CMR_UI.au3,
; CMR_Mail.au3, CMR_EDOC.au3, CMR_PDF.au3. Ces Case partagent des variables
; Local declarees une seule fois avant le Switch (ex: $sCmd_a, $sData_a) ;
; les extraire isolement vers ce fichier casserait ce partage et risquerait
; de regresser un routeur HTTP deja en production. Laisse intentionnellement
; vide -- voir les fichiers ci-dessus pour la logique reelle.

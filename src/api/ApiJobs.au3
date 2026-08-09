; ============================================================================
; ApiJobs.au3
; Routes jobs/batch.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

; Aucun bloc dedie trouve automatiquement dans Dispatch.txt.
; Fichier cree pour respecter l architecture cible et recevoir la migration.
;
; Les jobs/batch (FC_PAUSE, FC_PLAY, FC_STOP, ...) vivent comme des Case
; du meme grand Switch $sAction dans src/core/HttpRouter.au3 (route
; /api/action), avec la logique metier dans src/services/JobService.au3 et
; src/features/fc/. Meme risque que ApiActions.au3 : ces Case partagent des
; variables Local declarees avant le Switch, donc laisse intentionnellement
; vide -- voir HttpRouter.au3 et JobService.au3 pour la logique reelle.

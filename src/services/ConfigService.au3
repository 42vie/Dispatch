; ============================================================================
; ConfigService.au3
; Service config.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

; Aucun bloc dedie trouve automatiquement dans Dispatch.txt.
; Fichier cree pour respecter l architecture cible et recevoir la migration.
;
; La logique de configuration (profils/actions EDOC, config BKD) existe deja
; et fonctionne : voir src/core/Config.au3 (_InitConfig, _LoadActions,
; _RulesGUI, ...) pour la logique metier et src/api/ApiConfig.au3 pour les
; routes HTTP save-bkd-config/load-bkd-config. Rien n'a ete deplace ici pour
; eviter tout risque de regression sur un chemin de code deja en production.

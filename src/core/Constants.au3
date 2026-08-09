; ============================================================================
; Constants.au3
; Constantes extraites et centralisables.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
#include-once

; Chemin du fichier de config EDOC (profils/actions/mots-cles). Deplace
; depuis CMR_EDOC.au3 pour centraliser cette constante avec les autres --
; valeur inchangee, toujours "robot_v15_config.ini" (a ne pas confondre
; avec le config.ini a la racine du projet, qui sert a autre chose :
; host/port serveur, taux fuel, chemins contacts/backups).
Global Const $CFG_FILE = @ScriptDir & "\robot_v15_config.ini"

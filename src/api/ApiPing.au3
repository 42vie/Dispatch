; ============================================================================
; ApiPing.au3
; Routes ping extraites depuis le routeur principal si presentes.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
; Deplace depuis HttpRouter.au3 (bloc "/api/ping") pour que ce fichier
; contienne reellement ce que son nom annonce, comme prevu par
; l'architecture cible du projet.

Func _ApiPing_Handle($iSocket)
    _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
EndFunc

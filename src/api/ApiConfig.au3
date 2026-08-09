; ============================================================================
; ApiConfig.au3
; Routes config/options.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
; Deplace depuis HttpRouter.au3 (blocs "/api/save-bkd-config",
; "/api/load-bkd-config") -- logique inchangee, juste regroupee ici comme
; prevu par l'architecture cible du projet. Les routes "save-pj-config" /
; "save-config" / "save-cp-config" restent dans le Switch $sAction de
; HttpRouter.au3 : elles partagent des variables locales declarees avant
; le Switch, les en extraire isolement serait plus risque que ce qu'elles
; apportent.

Func _ApiConfig_SaveBkd($iSocket, $sBody)
    Local $sIniBKD = @ScriptDir & "\dispatch_config.ini"
    Local $sModeBKD = _GetJsonValue($sBody, "mode")
    Local $sCutoffBKD = _GetJsonValue($sBody, "cutoff")
    Local $sCustomBKD = _GetJsonValue($sBody, "customDate")

    If $sModeBKD = "" Then $sModeBKD = "auto"
    If $sCutoffBKD = "" Then $sCutoffBKD = "14:30"

    Local $bOkBKD = True
    If IniWrite($sIniBKD, "BKD", "Mode", $sModeBKD) = 0 Then $bOkBKD = False
    If IniWrite($sIniBKD, "BKD", "Cutoff", $sCutoffBKD) = 0 Then $bOkBKD = False
    If IniWrite($sIniBKD, "BKD", "CustomDate", $sCustomBKD) = 0 Then $bOkBKD = False

    If $bOkBKD Then
        _AuditLog("SAVE", "BKD config - mode=" & $sModeBKD & " cutoff=" & $sCutoffBKD & " custom=" & $sCustomBKD)
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
    Else
        _AuditLog("ERROR", "BKD config - impossible d'ecrire dispatch_config.ini")
        _SendHttpResponse($iSocket, 500, "application/json", '{"status":"error","reason":"cannot_write_bkd_config"}')
    EndIf
EndFunc

Func _ApiConfig_LoadBkd($iSocket)
    Local $sIniBKD2 = @ScriptDir & "\dispatch_config.ini"
    Local $sModeBKD2 = IniRead($sIniBKD2, "BKD", "Mode", "auto")
    Local $sCutoffBKD2 = IniRead($sIniBKD2, "BKD", "Cutoff", "14:30")
    Local $sCustomBKD2 = IniRead($sIniBKD2, "BKD", "CustomDate", "")

    Local $sRespBKD = '{"mode":"' & _JsonEscape($sModeBKD2) & _
                      '","cutoff":"' & _JsonEscape($sCutoffBKD2) & _
                      '","customDate":"' & _JsonEscape($sCustomBKD2) & '"}'

    _SendHttpResponse($iSocket, 200, "application/json", $sRespBKD)
EndFunc

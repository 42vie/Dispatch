; ============================================================================
; ApiNetwork.au3
; Routes/profils reseau.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
; Deplace depuis HttpRouter.au3 (blocs "/api/net-save", "/api/net-load",
; "/api/net-list") -- logique inchangee, juste regroupee ici comme prevu
; par l'architecture cible du projet.

Func _ApiNetwork_Save($iSocket, $sURL, $sBody)
    ; /api/net-save?path=F:\...\state.json — le body EST le JSON à écrire
    Local $sNetPath = StringMid($sURL, 20) ; après "/api/net-save?path="
    $sNetPath = _URIDecode($sNetPath)
    If $sNetPath <> "" Then
        _Net_SaveState($sNetPath, $sBody)
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
    Else
        _SendHttpResponse($iSocket, 400, "application/json", '{"error":"missing path"}')
    EndIf
EndFunc

Func _ApiNetwork_Load($iSocket, $sURL)
    ; /api/net-load?path=F:\...\state.json — retourne le contenu du fichier
    Local $sNetPath2 = StringMid($sURL, 20) ; après "/api/net-load?path="
    $sNetPath2 = _URIDecode($sNetPath2)
    Local $sNetJSON = _Net_LoadState($sNetPath2)
    _SendHttpResponse($iSocket, 200, "application/json", $sNetJSON)
EndFunc

Func _ApiNetwork_List($iSocket, $sURL)
    ; /api/net-list?pattern=F:\...\dispatch_state_*.json — liste les fichiers correspondants
    Local $sPattern = StringMid($sURL, 22) ; après "/api/net-list?pattern="
    $sPattern = _URIDecode($sPattern)
    Local $sDir = StringRegExpReplace($sPattern, "\\[^\\]*$", "")
    Local $sGlob = StringRegExpReplace($sPattern, "^.*\\", "")
    Local $sFileList = '["' ; on construit un array JSON
    Local $hSearch2 = FileFindFirstFile($sDir & "\" & $sGlob)
    Local $bFirst = True
    If $hSearch2 <> -1 Then
        While True
            Local $sFound = FileFindNextFile($hSearch2)
            If @error Then ExitLoop
            If Not $bFirst Then $sFileList &= ',"'
            $sFileList &= StringReplace($sDir & "\" & $sFound, "\", "\\") & '"'
            $bFirst = False
        WEnd
        FileClose($hSearch2)
    EndIf
    If $bFirst Then
        $sFileList = "[]"
    Else
        $sFileList &= "]"
    EndIf
    _SendHttpResponse($iSocket, 200, "application/json", $sFileList)
EndFunc

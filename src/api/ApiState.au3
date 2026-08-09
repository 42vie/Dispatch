; ============================================================================
; ApiState.au3
; Routes et helpers state/status.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
; Deplace depuis HttpRouter.au3 (blocs "/api/load", "/api/save",
; "/api/save-status", "/api/load-status", "/api/save-data", "/api/load-data")
; -- logique inchangee, juste regroupee ici comme prevu par l'architecture
; cible du projet.

Func _ApiState_Load($iSocket)
    Local $sJson = "{}"
    If FileExists($g_sSaveFile) Then
        Local $hJsonRead = FileOpen($g_sSaveFile, 256) ; 256 = UTF-8 sans BOM
        If $hJsonRead <> -1 Then
            $sJson = FileRead($hJsonRead)
            FileClose($hJsonRead)
        EndIf
    EndIf
    _SendHttpResponse($iSocket, 200, "application/json", $sJson)
EndFunc

Func _ApiState_Save($iSocket, $sBody)
    Local $hFile = FileOpen($g_sSaveFile, 2 + 256) ; 256 = UTF-8 sans BOM
    FileWrite($hFile, $sBody)
    FileClose($hFile)
    _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
EndFunc

Func _ApiState_SaveStatus($iSocket, $sBody)
    _AuditLog("SAVE", "status — " & StringLen($sBody) & " bytes")
    Local $hFileS = FileOpen($g_sStatusFile, 2 + 256)
    FileWrite($hFileS, $sBody)
    FileClose($hFileS)
    _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
EndFunc

Func _ApiState_LoadStatus($iSocket)
    Local $sJsonS = "[]"
    If FileExists($g_sStatusFile) Then
        Local $hReadS = FileOpen($g_sStatusFile, 256)
        If $hReadS <> -1 Then
            $sJsonS = FileRead($hReadS)
            FileClose($hReadS)
        EndIf
    EndIf
    _SendHttpResponse($iSocket, 200, "application/json", $sJsonS)
EndFunc

Func _ApiState_SaveData($iSocket, $sBody)
    _AuditLog("SAVE", "data — " & StringLen($sBody) & " bytes")
    _BackupRotate($g_sDataFile, 5)
    Local $hFileD = FileOpen($g_sDataFile, 2 + 256)
    FileWrite($hFileD, $sBody)
    FileClose($hFileD)
    _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
EndFunc

Func _ApiState_LoadData($iSocket)
    Local $sJsonD = "{}"
    If FileExists($g_sDataFile) Then
        Local $hReadD = FileOpen($g_sDataFile, 256)
        If $hReadD <> -1 Then
            $sJsonD = FileRead($hReadD)
            FileClose($hReadD)
        EndIf
    EndIf
    _SendHttpResponse($iSocket, 200, "application/json", $sJsonD)
EndFunc

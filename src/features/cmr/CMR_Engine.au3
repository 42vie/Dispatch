; ============================================================================
; CMR_Engine.au3
; Moteur CMR.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _CMR_RunFromData($sData)
    $g_bCMRRunning = True
    $g_sCMRStatus = "CMR : préparation queue..."
    $g_sCMRLastJSON = '{"status":"running","message":"preparation"}'
    _AuditLog("CMR", "RunFromData raw=" & StringLeft($sData, 500))

    _QueueClearRows()
    _ClearResults()

    Local $sTxt = StringReplace($sData, "@@BL@@", @LF & @LF)
    $sTxt = StringReplace($sTxt, ",", @LF)
    $sTxt = StringReplace($sTxt, "|", @LF)
    Local $hQ = FileOpen($g_sCMRQueueFile, 2 + 256)
    If $hQ <> -1 Then
        FileWrite($hQ, $sTxt)
        FileClose($hQ)
    EndIf

    Local $sSplitData = StringReplace($sData, "@@BL@@", Chr(30))
    Local $aGroups = StringSplit($sSplitData, Chr(30), 2)
    For $i = 0 To UBound($aGroups) - 1
        Local $sBlock = StringReplace($aGroups[$i], ",", @LF)
        $sBlock = StringReplace($sBlock, "|", @LF)
        If StringStripWS($sBlock, 3) <> "" Then _QueueAppendBlock($sBlock)
    Next

    If UBound($g_aQueueBlocks) = 0 Then
        $g_sCMRStatus = "CMR : queue vide."
        $g_sCMRLastJSON = '{"status":"error","message":"queue_vide","results":[]}'
        $g_bCMRRunning = False
        Return False
    EndIf

    $g_sCMRStatus = "CMR : génération en cours..."
    $g_sCMRLastJSON = '{"status":"running","message":"generation","results":[]}'
    Local $bOk = _RunGenerateQueue()

    $g_sCMRLastJSON = _CMR_BuildResultsJSON($bOk)
    $g_bCMRRunning = False
    If $bOk Then
        $g_sCMRStatus = "CMR : terminé. " & UBound($g_aBL) & " BL généré(s)."
    Else
        $g_sCMRStatus = "CMR : terminé avec erreurs. " & UBound($g_aBL) & " BL généré(s)."
    EndIf
    _AuditLog("CMR", $g_sCMRStatus)
    Return $bOk
EndFunc

; ============================================================================
; Audit_FC.au3
; Audit File Closing.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _FC_AuditInit($sLabel)
    $g_sFC_AuditLog = ""
    _FC_AuditLog("====== AUDIT FC : " & $sLabel & " ======")
    _FC_AuditLog("PC       : " & @ComputerName)
    _FC_AuditLog("User     : " & @UserName)
    _FC_AuditLog("OS       : " & @OSVersion & " " & @OSArch)
    _FC_AuditLog("RAM Free : " & Round(MemGetStats()[2] / 1024, 0) & " MB / " & Round(MemGetStats()[1] / 1024, 0) & " MB")
    _FC_AuditLog("CPU      : " & @CPUArch)
    Local $aProc = ProcessList("ETMS.exe")
    If IsArray($aProc) Then
        _FC_AuditLog("ETMS PID : " & ($aProc[0][0] > 0 ? $aProc[1][1] : "NON TROUVE"))
    Else
        _FC_AuditLog("ETMS PID : NON TROUVE")
    EndIf
EndFunc

Func _FC_AuditLog($sMsg)
    If Not $g_bFC_Audit Then Return
    Local $sLine = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & "  " & $sMsg
    $g_sFC_AuditLog &= $sLine & @CRLF
EndFunc

Func _FC_AuditStep($iStep, $sDesc)
    _FC_AuditLog("── STEP " & $iStep & " : " & $sDesc & " ──")
EndFunc

Func _FC_AuditWinState($sClass, $sLabel)
    Local $bExists = WinExists($sClass)
    Local $sState  = "exists=" & $bExists
    If $bExists Then
        Local $hW  = WinGetHandle($sClass)
        Local $aPos = WinGetPos($hW)
        Local $sTitle = WinGetTitle($hW)
        $sState &= " | hwnd=" & $hW & " | title=" & $sTitle
        If IsArray($aPos) Then $sState &= " | pos=" & $aPos[0] & "," & $aPos[1] & " size=" & $aPos[2] & "x" & $aPos[3]
        $sState &= " | state=" & WinGetState($hW)
    EndIf
    _FC_AuditLog("  WIN[" & $sLabel & "] " & $sState)
EndFunc

Func _FC_AuditCtrl($hWnd, $sCtrl, $sLabel)
    Local $sTxt = ControlGetText($hWnd, "", $sCtrl)
    Local $hCtrl = ControlGetHandle($hWnd, "", $sCtrl)
    _FC_AuditLog("  CTRL[" & $sLabel & "] handle=" & $hCtrl & " | text='" & StringLeft($sTxt, 120) & "'")
EndFunc

Func _FC_AuditTiming($sLabel, $nMs)
    Local $sSuffix = ""
    If $nMs > 5000 Then $sSuffix = " *** LENT ***"
    If $nMs > 10000 Then $sSuffix = " *** TRES LENT ***"
    If $nMs > 20000 Then $sSuffix = " *** CRITIQUE ***"
    _FC_AuditLog("  TIMING[" & $sLabel & "] " & Round($nMs, 0) & " ms" & $sSuffix)
EndFunc

Func _FC_AuditFileCheck($sPath)
    _FC_AuditLog("  FILE[" & $sPath & "]")
    If FileExists($sPath) Then
        Local $iSize = FileGetSize($sPath)
        Local $sTime = FileGetTime($sPath, 0, 1)
        _FC_AuditLog("    exists=TRUE | size=" & $iSize & " octets (" & Round($iSize/1024, 1) & " KB) | modified=" & $sTime)
    Else
        _FC_AuditLog("    exists=FALSE *** FICHIER INTROUVABLE ***")
        ; Verifier le dossier parent
        Local $sDir = StringRegExpReplace($sPath, "\\[^\\]+$", "")
        If FileExists($sDir) Then
            _FC_AuditLog("    dossier parent OK : " & $sDir)
            ; Lister les .eds dans le dossier
            Local $hSearch = FileFindFirstFile($sDir & "\*.eds")
            If $hSearch <> -1 Then
                Local $sFiles = ""
                Local $iCount = 0
                While 1
                    Local $sFile = FileFindNextFile($hSearch)
                    If @error Then ExitLoop
                    $sFiles &= $sFile & ", "
                    $iCount += 1
                WEnd
                FileClose($hSearch)
                _FC_AuditLog("    .eds trouves (" & $iCount & ") : " & StringTrimRight($sFiles, 2))
            Else
                _FC_AuditLog("    AUCUN .eds dans le dossier !")
            EndIf
        Else
            _FC_AuditLog("    *** DOSSIER PARENT INTROUVABLE : " & $sDir & " ***")
        EndIf
    EndIf
EndFunc

Func _FC_AuditSave($sNum)
    If $g_sFC_AuditLog = "" Then Return ""
    Local $sDir = @ScriptDir & "\logs"
    If Not FileExists($sDir) Then DirCreate($sDir)
    Local $sFile = $sDir & "\FC_AUDIT_" & $sNum & "_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".log"
    Local $hFile = FileOpen($sFile, 2)
    If $hFile = -1 Then Return ""
    FileWrite($hFile, $g_sFC_AuditLog)
    FileClose($hFile)
    Return $sFile
EndFunc

Func _FC_AuditShow($sNum)
    If Not $g_bFC_Audit Then Return
    Local $sFile = _FC_AuditSave($sNum)
    If $sFile = "" Then Return

    ; GUI avec le rapport complet
    Local $hGUI = GUICreate("AUDIT FC — " & $sNum, 750, 550, -1, -1)
    GUISetBkColor(0x1E1E1E, $hGUI)
    GUISetFont(9, 400, 0, "Consolas")
    Local $idEdit = GUICtrlCreateEdit($g_sFC_AuditLog, 5, 5, 740, 495, BitOR(0x0004, 0x0800, 0x00200000))
    GUICtrlSetBkColor($idEdit, 0x1E1E1E)
    GUICtrlSetColor($idEdit, 0x00FF00)
    Local $idBtnCopy = GUICtrlCreateButton("Copier", 5, 505, 120, 35)
    Local $idBtnOpen = GUICtrlCreateButton("Ouvrir le .log", 130, 505, 150, 35)
    Local $idBtnClose = GUICtrlCreateButton("Fermer", 625, 505, 120, 35)
    GUISetState(@SW_SHOW, $hGUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idBtnClose
                ExitLoop
            Case $idBtnCopy
                ClipPut($g_sFC_AuditLog)
                ToolTip("Copié !", Default, Default, "Audit FC", 1)
                Sleep(1000)
                ToolTip("")
            Case $idBtnOpen
                ShellExecute($sFile)
        EndSwitch
    WEnd
    GUIDelete($hGUI)
EndFunc

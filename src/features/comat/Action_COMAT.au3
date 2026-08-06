; ============================================================================
; Action_COMAT.au3
; Action COMAT solo.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _Run_COMAT_Single($Num)
    $Num = StringStripWS($Num, 8)
    If $Num = "" Then Return
    Local $hWnd = WinGetHandle("[CLASS:TfmBrowser]")
    If $hWnd = 0 Or Not WinExists($hWnd) Then
        _NotifyError("COMAT", "Fenêtre E.TMS introuvable.")
        $bCOMAT_Stop = True
        Return
    EndIf
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)

    _COMAT_Spinner("COMAT [" & $Num & "] 1/5 - LOG J...")
    Send("{PGUP}")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlFocus($hWnd, "", $COMAT_LOG_CTRL)
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG " & $Num)
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    _COMAT_SmartSleep($COMAT_DELAY_LOAD)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 2/5 - F3...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    Send("{F3}")
    _COMAT_SmartSleep($COMAT_DELAY_L)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 3/5 - F5 x4...")
    Local $k
    For $k = 1 To 4
        Send("{F5}")
        _COMAT_SmartSleep($COMAT_DELAY_M)
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Next
    _COMAT_SmartSleep($COMAT_DELAY_L)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 4/5 - F1 + TAB + C...")
    Send("{F1}")
    _COMAT_SmartSleep($COMAT_DELAY_L)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    For $k = 1 To 6
        Send("{TAB}")
        _COMAT_SmartSleep($COMAT_DELAY_S)
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Next
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Send("C")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    For $k = 1 To 4
        Send("{F5}")
        If $k = 4 Then
            _COMAT_SmartSleep($COMAT_DELAY_L)
        Else
            _COMAT_SmartSleep($COMAT_DELAY_M)
        EndIf
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Next
    _COMAT_SmartSleep(400)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 5/5 - Retour LOG...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlFocus($hWnd, "", $COMAT_LOG_CTRL)
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    _COMAT_SmartSleep(1000)
    ToolTip("")
EndFunc

Func _Action_COMAT_Solo($Num)
    $Num = StringStripWS($Num, 8)
    If $Num = "" Then
    MsgBox(48+262144, "Erreur", "Aucun numéro de dossier.")
    Return
EndIf
    If MsgBox(1+32+262144, "COMAT Solo", "Lancer COMAT sur le dossier : " & $Num & " ?") = 2 Then Return
    _Run_COMAT_Single($Num)
    ToolTip("")
    MsgBox(64+262144, "COMAT", "Dossier " & $Num & " traité.")
EndFunc

Func _COMAT_Spinner($sTxt)
    ToolTip($sTxt, 0, 0, "Robot E.TMS — COMAT", 1)
EndFunc

Func _COMAT_SmartSleep($iMs)
    Local $iSlept = 0
    While $iSlept < $iMs
        _Tracker_PollButtons()
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
        _COMAT_WaitIfPaused()
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
        Sleep(100)
        $iSlept += 100
    WEnd
EndFunc

; ==============================================================================
; URI DECODE (pour les query params)
; ==============================================================================

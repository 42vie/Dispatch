; ============================================================================
; Action_ETMS.au3
; Actions ETMS.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _GetWindowETMS()
    Return WinGetHandle("[CLASS:TfmBrowser]")
EndFunc

Func _GetETMSInstance($hWnd)
    Local $sTitle = WinGetTitle($hWnd)
    If StringInStr($sTitle, "(LOG)")   Then Return "91"
    If StringInStr($sTitle, "(NOTES)") Then Return "83"
    If StringInStr($sTitle, "(REFS)")  Then Return "109"
    If StringInStr($sTitle, "(DIMST)") Then Return "300"
    If StringInStr($sTitle, "(HIST)")  Then Return "207"
    Return "91"
EndFunc

Func _ActionETMS($sBouton, $sNumDossier)
    If $sBouton = "EDOC" Then
        If _ActionEDOC($sNumDossier) Then Return
    EndIf
    Local $hWnd = WinGetHandle("[CLASS:TfmBrowser]")
    If Not WinExists($hWnd) Then
        _NotifyError("E.TMS", "E.TMS est fermé ou introuvable. Ouvrez E.TMS avant de lancer une action.")
        Return
    EndIf

    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)

    Local $sInst = _GetETMSInstance($hWnd)
    Local $sCtrl = "[CLASS:TEIEdit; INSTANCE:" & $sInst & "]"

    ; ══ F8 seul : envoyer F8 sans écrire de commande ══
    If $sBouton = "F8" Then
        ControlFocus($hWnd, "", $sCtrl)
        ControlSend($hWnd, "", $sCtrl, "{F8}")
        _AuditLog("ETMS", "BG: F8 (execute)")
        Return
    EndIf

    ; Préparer la commande
    Send("{PGUP}")
    Sleep(150)
    Local $sCommande = $sBouton & " " & $sNumDossier
    If $sBouton = "LOG X" Then $sCommande = "LOG X"

    ControlFocus($hWnd, "", $sCtrl)
    ControlSetText($hWnd, "", $sCtrl, $sCommande)
    Sleep(150)
    ControlSend($hWnd, "", $sCtrl, "{F8}")
    _AuditLog("ETMS", "BG: " & $sCommande)
EndFunc

; ==============================================================================
; CALCUL DATES / JOURS OUVRÉS
; ==============================================================================

Func _OpenEdsInEtms()
    Local $hWnd = WinGetHandle($ETMS_WINDOW)
    If $hWnd = 0 Then
        MsgBox(16+262144, $APP_TITLE, "ETMS introuvable : [CLASS:TfmBrowser].")
        Return False
    EndIf
    If Not FileExists($EDS_PATH) Then
        MsgBox(16+262144, $APP_TITLE, "EDS introuvable :" & @CRLF & $EDS_PATH)
        Return False
    EndIf
    Local $ok = False, $try = 0
    While Not $ok And $try < 3
        $try += 1
        If $try > 1 And WinExists($FILEOPEN_WIN) Then WinClose($FILEOPEN_WIN)
        WinActivate($hWnd)
        WinWaitActive($hWnd, "", 3)
        Sleep(500)
        ControlClick($hWnd, "", $ETMS_TOOLBAR, "LEFT", 1, $TOOLBAR_X, $TOOLBAR_Y)
        Local $t = TimerInit()
        While Not WinExists($FILEOPEN_WIN)
            Sleep(100)
            If _IsPressed("1B") Then Return False
            If TimerDiff($t) > 10000 Then ExitLoop
        WEnd
        If Not WinExists($FILEOPEN_WIN) Then ContinueLoop
        WinActivate($FILEOPEN_WIN)
        WinWaitActive($FILEOPEN_WIN, "", 3)
        ControlSetText($FILEOPEN_WIN, "", $FILEOPEN_EDIT, "")
        Sleep(100)
        ControlSetText($FILEOPEN_WIN, "", $FILEOPEN_EDIT, $EDS_PATH)
        Sleep(200)
        Send("{ENTER}")
        Sleep(800)
        If Not WinExists($FILEOPEN_WIN) Then $ok = True
    WEnd
    Return $ok
EndFunc

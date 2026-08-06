; ============================================================================
; JobService.au3
; Service batch/queue.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _FC_WaitIfPaused()
    While $bFC_Pause And Not $bFC_Stop And Not $bFC_Skip
        _Spinner("EN PAUSE — cliquez Play dans la fenêtre de contrôle")
        _Tracker_PollButtons()
        Sleep(80)
    WEnd
EndFunc

; Sleep réactif FC : découpe en tranches de 100ms + poll GUI

Func _COMAT_WaitIfPaused2()
    While $bCOMAT_Pause And Not $bCOMAT_Stop And Not $bCOMAT_Skip
        _Tracker_PollButtons()
        Sleep(100)
    WEnd
EndFunc

; Les HotKeys restent comme fallback (si la GUI n'est pas visible)

Func _HK_FC_PauseToggle()
    $bFC_Pause = Not $bFC_Pause
EndFunc

Func _HK_FC_Stop()
    $bFC_Stop = True
    $bFC_Pause = False
EndFunc

Func _HK_COMAT_PauseToggle()
    $bCOMAT_Pause = Not $bCOMAT_Pause
EndFunc

Func _HK_COMAT_Stop()
    $bCOMAT_Stop = True
    $bCOMAT_Pause = False
EndFunc

; ==============================================================================
; E.TMS ET EDOC
; ==============================================================================

Func _Tracker_Start($sTitle, $aList)
    $g_iTrackCount = UBound($aList)
    If $g_iTrackCount = 0 Then Return False
    ReDim $g_aTrackIDs[$g_iTrackCount]

    ; GUI toujours au premier plan (0x00000008 = WS_EX_TOPMOST) mais sans voler le focus
    $g_hTracker = GUICreate($sTitle, 320, 460, @DesktopWidth - 340, 50, -1, 0x00000008)
    GUISetBkColor(0x2D2D30, $g_hTracker)
    GUISetFont(9, 400, 0, "Segoe UI")

    ; Barre de progression
    $g_idTrackProg = GUICtrlCreateProgress(10, 10, 300, 18)

    ; Label info
    $g_idTrackLbl = GUICtrlCreateLabel("0 / " & $g_iTrackCount & " dossiers", 10, 33, 300, 18, 1)
    GUICtrlSetColor(-1, 0xFFFFFF)

    ; Info dossier en cours
    $g_idBatchInfo = GUICtrlCreateLabel("", 10, 52, 300, 18, 1)
    GUICtrlSetColor(-1, 0xFFCC00)

    ; ══ BOUTONS DE CONTRÔLE ══
    $g_idBtnPause = GUICtrlCreateButton("⏸ Pause", 10, 74, 72, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0xFF9900)

    $g_idBtnPlay = GUICtrlCreateButton("▶ Play", 86, 74, 72, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0x00CC55)
    GUICtrlSetState(-1, $GUI_DISABLE)

    $g_idBtnSkip = GUICtrlCreateButton("⏭ Passer", 162, 74, 78, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0x3399FF)

    $g_idBtnStop = GUICtrlCreateButton("⏹ Stop", 244, 74, 66, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0xFF4444)

    ; Liste des dossiers
    $g_idTrackLV = GUICtrlCreateListView("Statut|Numero J", 10, 110, 300, 340)
    GUICtrlSetBkColor(-1, 0x1E1E1E)
    GUICtrlSetColor(-1, 0xCCCCCC)
    GUICtrlSendMsg($g_idTrackLV, 0x101E, 0, 80)
    GUICtrlSendMsg($g_idTrackLV, 0x101E, 1, 190)
    For $i = 0 To $g_iTrackCount - 1
        $g_aTrackIDs[$i] = GUICtrlCreateListViewItem("Attente|" & $aList[$i], $g_idTrackLV)
        GUICtrlSetColor($g_aTrackIDs[$i], 0xAAAAAA)
    Next

    GUISetState(@SW_SHOWNOACTIVATE, $g_hTracker)
    Return True
EndFunc

Func _Tracker_PollButtons()
    If $g_hTracker = 0 Then Return
    Local $iMsg = GUIGetMsg()
    Switch $iMsg
        Case $g_idBtnPause
            $bFC_Pause = True
            $bCOMAT_Pause = True
            GUICtrlSetState($g_idBtnPause, $GUI_DISABLE)
            GUICtrlSetState($g_idBtnPlay, $GUI_ENABLE)
            GUICtrlSetData($g_idBatchInfo, "⏸ EN PAUSE — Cliquez Play pour reprendre")
        Case $g_idBtnPlay
            $bFC_Pause = False
            $bCOMAT_Pause = False
            GUICtrlSetState($g_idBtnPlay, $GUI_DISABLE)
            GUICtrlSetState($g_idBtnPause, $GUI_ENABLE)
        Case $g_idBtnSkip
            $bFC_Skip = True
            $bCOMAT_Skip = True
            $bFC_Pause = False
            $bCOMAT_Pause = False
            GUICtrlSetState($g_idBtnPlay, $GUI_DISABLE)
            GUICtrlSetState($g_idBtnPause, $GUI_ENABLE)
        Case $g_idBtnStop
            $bFC_Stop = True
            $bCOMAT_Stop = True
            $bFC_Pause = False
            $bCOMAT_Pause = False
    EndSwitch
EndFunc

Func _Tracker_End()
    If $g_hTracker Then
        GUICtrlSetData($g_idTrackProg, 100)
        GUICtrlSetData($g_idTrackLbl, "Traitement terminé !")
        GUICtrlSetData($g_idBatchInfo, "✓ Tout est fait")
        GUICtrlSetState($g_idBtnPause, $GUI_DISABLE)
        GUICtrlSetState($g_idBtnPlay, $GUI_DISABLE)
        GUICtrlSetState($g_idBtnSkip, $GUI_DISABLE)
        GUICtrlSetState($g_idBtnStop, $GUI_DISABLE)
        Sleep(2000)
        GUIDelete($g_hTracker)
        $g_hTracker = 0
    EndIf
EndFunc

; ==============================================================================
; COMAT
; ==============================================================================

Func _COMAT_WaitIfPaused()
    While $bCOMAT_Pause And Not $bCOMAT_Stop And Not $bCOMAT_Skip
        _COMAT_Spinner("EN PAUSE — cliquez Play dans la fenêtre de contrôle")
        _Tracker_PollButtons()
        Sleep(80)
    WEnd
EndFunc

; Sleep réactif : découpe le sleep en tranches de 100ms et poll les boutons GUI
; Permet Pause/Skip/Stop instantané même pendant les longues attentes

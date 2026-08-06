; ============================================================================
; Batch_COMAT.au3
; Batch COMAT.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _Batch_COMAT($sData)
    If $sData = "" Then Return
    $bCOMAT_Stop = False
    $bCOMAT_Pause = False
    $bCOMAT_Skip = False
    HotKeySet("{F9}", "_HK_COMAT_PauseToggle")
    HotKeySet("{ESCAPE}", "_HK_COMAT_Stop")

    Local $aJobs = StringSplit($sData, "|")
    Local $aValid[$aJobs[0]]
    For $i = 1 To $aJobs[0]
        Local $aInfos = StringSplit($aJobs[$i], ";")
        $aValid[$i-1] = $aInfos[1]
    Next
    _Tracker_Start("COMAT en masse", $aValid)
    Local $iDone = 0, $iStopped = 0
    Local $sRemaining = ""
    For $i = 1 To $aJobs[0]
        Local $aDetails = StringSplit($aJobs[$i], ";")
        If $aDetails[0] >= 1 Then
            Local $sNumJ = $aDetails[1]
            $bCOMAT_Skip = False
            _Tracker_Update($i-1, 1)
            _COMAT_WaitIfPaused2()
            If $bCOMAT_Stop Then
                _Tracker_Update($i-1, 3)
                $iStopped = 1
                ; Collecter les dossiers restants (celui-ci + suivants)
                For $r = $i To $aJobs[0]
                    Local $aR = StringSplit($aJobs[$r], ";")
                    If $aR[0] >= 1 Then $sRemaining &= $aR[1] & @CRLF
                Next
                ExitLoop
            EndIf
            If $bCOMAT_Skip Then
                _Tracker_Update($i-1, 4)
                $bCOMAT_Skip = False
                ContinueLoop
            EndIf
            _Run_COMAT_Single($sNumJ)
            If $bCOMAT_Stop Then
                _Tracker_Update($i-1, 3)
                $iStopped = 1
                ; Collecter les dossiers restants (suivants seulement, celui-ci peut être partiel)
                For $r = $i + 1 To $aJobs[0]
                    Local $aR2 = StringSplit($aJobs[$r], ";")
                    If $aR2[0] >= 1 Then $sRemaining &= $aR2[1] & @CRLF
                Next
                ExitLoop
            EndIf
            If $bCOMAT_Skip Then
                _Tracker_Update($i-1, 4)
                $bCOMAT_Skip = False
            Else
                _Tracker_Update($i-1, 2)
                $iDone += 1
            EndIf
            _Tracker_PollButtons()
            _COMAT_SmartSleep(150)
        EndIf
    Next
    HotKeySet("{F9}")
    HotKeySet("{ESCAPE}")
    _Tracker_End()
    ; Bilan final
    If $iStopped And $sRemaining <> "" Then
        ; Copier les dossiers restants dans le presse-papier pour reprise facile
        ClipPut(StringStripWS($sRemaining, 2))
        MsgBox(48+262144, "COMAT — Arrêté", _
            $iDone & " dossier(s) traité(s) sur " & $aJobs[0] & "." & @CRLF & @CRLF & _
            "Dossiers restants (copiés dans le presse-papier) :" & @CRLF & $sRemaining)
    ElseIf $iStopped Then
        MsgBox(48+262144, "COMAT — Arrêté", $iDone & " dossier(s) traité(s) sur " & $aJobs[0] & ".")
    EndIf
    $bCOMAT_Stop = False
    $bCOMAT_Pause = False
    $bCOMAT_Skip = False
EndFunc

Func _GUI_COMAT_Multi()
    Local $hComat    = GUICreate("COMAT MULTI", 310, 440, -1, -1, -1, 0x00000008)
    GUISetBkColor(0x2D2D30)
    GUISetFont(9, 400, 0, "Segoe UI")
    GUICtrlCreateLabel("Numéros J (un par ligne) :", 10, 10, 290, 18)
    GUICtrlSetColor(-1, 0xFFFFFF)
    Local $idEdit    = GUICtrlCreateEdit("", 10, 30, 290, 310, BitOR(0x0004, 0x0040, 0x2000))
    GUICtrlSetBkColor(-1, 0x1E1E1E)
    GUICtrlSetColor(-1, 0xFFFFFF)
    Local $idBtnRun  = GUICtrlCreateButton("> LANCER COMAT MULTI", 10, 352, 290, 45)
    GUICtrlSetFont(-1, 10, 800)
    GUICtrlSetColor(-1, 0x007ACC)
    Local $idBtnStop = GUICtrlCreateButton("STOP", 10, 405, 290, 22)
    GUICtrlSetFont(-1, 8, 400)
    GUICtrlSetColor(-1, 0xCC0000)
    GUISetState(@SW_SHOW, $hComat)
    Local $msg_gui  = 0
    Local $sDataGui = ""
    Local $aLinesGui[1]
    Local $aValidGui[100]
    Local $iTotalGui = 0
    Local $sNumGui   = ""
    While 1
        $msg_gui = GUIGetMsg()
        Switch $msg_gui
            Case -3
                GUIDelete($hComat)
                Return
            Case $idBtnStop
                $bCOMAT_Stop = True
                GUIDelete($hComat)
                Return
            Case $idBtnRun
                $sDataGui = GUICtrlRead($idEdit)
                GUIDelete($hComat)
                If StringStripWS($sDataGui, 8) = "" Then
    MsgBox(48+262144, "Vide", "La liste est vide !")
    Return
EndIf
                $aLinesGui = StringSplit(StringStripCR($sDataGui), @LF)
                $iTotalGui = 0
                ReDim $aValidGui[$aLinesGui[0] + 1]
                For $j = 1 To $aLinesGui[0]
                    $sNumGui = StringStripWS($aLinesGui[$j], 8)
                    If $sNumGui <> "" Then
                        $aValidGui[$iTotalGui] = $sNumGui
                        $iTotalGui += 1
                    EndIf
                Next
                If $iTotalGui = 0 Then
    MsgBox(48+262144, "Vide", "Aucun numéro valide.")
    Return
EndIf
                ReDim $aValidGui[$iTotalGui]
                If MsgBox(1+32+262144, "Confirmation", $iTotalGui & " dossier(s) à traiter. GO ?") = 2 Then Return
                _Tracker_Start("COMAT Multi - Suivi", $aValidGui)
                HotKeySet("{F9}", "_HK_COMAT_PauseToggle")
                HotKeySet("{ESCAPE}", "_HK_COMAT_Stop")
                $bCOMAT_Stop = False
                $bCOMAT_Pause = False
                $bCOMAT_Skip = False
                Local $iDoneG = 0, $iStoppedG = 0
                Local $sRemainingG = ""
                For $j = 0 To $iTotalGui - 1
                    $bCOMAT_Skip = False
                    _Tracker_Update($j, 1)
                    _COMAT_WaitIfPaused2()
                    If $bCOMAT_Stop Then
                        _Tracker_Update($j, 3)
                        $iStoppedG = 1
                        For $rr = $j To $iTotalGui - 1
                            $sRemainingG &= $aValidGui[$rr] & @CRLF
                        Next
                        ExitLoop
                    EndIf
                    If $bCOMAT_Skip Then
                        _Tracker_Update($j, 4)
                        $bCOMAT_Skip = False
                        ContinueLoop
                    EndIf
                    _Run_COMAT_Single($aValidGui[$j])
                    If $bCOMAT_Stop Then
                        _Tracker_Update($j, 3)
                        $iStoppedG = 1
                        For $rr = $j + 1 To $iTotalGui - 1
                            $sRemainingG &= $aValidGui[$rr] & @CRLF
                        Next
                        ExitLoop
                    EndIf
                    If $bCOMAT_Skip Then
                        _Tracker_Update($j, 4)
                        $bCOMAT_Skip = False
                    Else
                        _Tracker_Update($j, 2)
                        $iDoneG += 1
                    EndIf
                    _Tracker_PollButtons()
                    _COMAT_SmartSleep(300)
                Next
                HotKeySet("{F9}")
                HotKeySet("{ESCAPE}")
                _Tracker_End()
                If $iStoppedG And $sRemainingG <> "" Then
                    ClipPut(StringStripWS($sRemainingG, 2))
                    MsgBox(48+262144, "COMAT — Arrêté", _
                        $iDoneG & " dossier(s) traité(s) sur " & $iTotalGui & "." & @CRLF & @CRLF & _
                        "Dossiers restants (copiés dans le presse-papier) :" & @CRLF & $sRemainingG)
                ElseIf $iStoppedG Then
                    MsgBox(48+262144, "COMAT — Arrêté", $iDoneG & " dossier(s) traité(s) sur " & $iTotalGui & ".")
                Else
                    MsgBox(64+262144, "Terminé", "Traitement COMAT terminé — " & $iDoneG & " dossier(s).")
                EndIf
                $bCOMAT_Stop = False
                $bCOMAT_Pause = False
                $bCOMAT_Skip = False
                Return
        EndSwitch
    WEnd
EndFunc

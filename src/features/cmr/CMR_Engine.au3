; ============================================================================
; CMR_Engine.au3
; Moteur CMR : generation de la queue de BL, boucle principale,
; gestion des outputs ETMS, dialogues modernes (avec repli headless quand
; l'app tourne sans fenetre _Main), et l'ecran GUI historique (_Main, non
; utilise depuis le passage a l'interface web -- conserve pour reference).
; ----------------------------------------------------------------------------
; Deplace depuis StateService.au3 (bloc CMR_FULL_FUSION_MODULE_APPENDED) pour
; que le code CMR vive dans src/features/cmr/, comme le reste du projet.
; ============================================================================

; Point d'entree reel de la generation CMR, appele par HttpRouter.au3 sur
; l'action "CMR_GENERATE" (POST /api/action). C'est ce que le bouton
; "Test CMR" de DispatchDoctor declenche.
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

Func _Main()
    DirCreate(@ScriptDir & "\Config")
    DirCreate($INPUT_PATH)
    DirCreate($OUTPUT_PATH)
    DirCreate($HTML_PATH)
    DirCreate($PDF_PATH)
    DirCreate($SNAPSHOT_PATH)
    DirCreate(@ScriptDir & "\HTML") ; fallback local PRO CLEAN
    _SP_InitDefaultsIfNeeded()

    $g_sOperator = IniRead($INI_PATH, "USER", "Operator", "")
    If StringStripWS($g_sOperator, 3) = "" Then
        $g_sOperator = InputBox($APP_TITLE, "Nom utilisateur / Operator", @UserName, "", 420, 140)
        If @error Or StringStripWS($g_sOperator, 3) = "" Then $g_sOperator = @UserName
        IniWrite($INI_PATH, "USER", "Operator", $g_sOperator)
    EndIf

    $g_hGui = GUICreate($APP_TITLE, 820, 610, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_MINIMIZEBOX))
    GUISetFont(9, 400, 0, "Segoe UI", $g_hGui)
    GUISetBkColor(0xF3F6FB, $g_hGui)
    Local $idTab = GUICtrlCreateTab(10, 10, 800, 590)

    GUICtrlCreateTabItem("Queue")
    Local $lblQueue = GUICtrlCreateLabel("Queue BL visuelle : 1 ligne dans la liste = 1 BL à générer", 24, 45, 520, 18)
    GUICtrlSetFont($lblQueue, 9, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("Colle ici les J d’un BL, puis clique Ajouter BL. Pour plusieurs BL d’un coup : sépare par ligne vide ou --- puis Import auto.", 24, 66, 360, 34)
    $idQueue = GUICtrlCreateEdit("", 24, 103, 350, 95, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_WANTRETURN))
    Local $bAddQueue = GUICtrlCreateButton("Ajouter BL", 24, 204, 110, 28)
    Local $bImportQueue = GUICtrlCreateButton("Import auto", 144, 204, 110, 28)
    Local $bRemoveQueue = GUICtrlCreateButton("Retirer", 264, 204, 110, 28)
    $idQueueList = GUICtrlCreateListView("OK|#|Nb|J dans le BL|Statut", 24, 240, 350, 238, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_VSCROLL))
    _GUICtrlListView_SetExtendedListViewStyle($idQueueList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES))
    _GUICtrlListView_SetColumnWidth($idQueueList, 0, 34)
    _GUICtrlListView_SetColumnWidth($idQueueList, 1, 35)
    _GUICtrlListView_SetColumnWidth($idQueueList, 2, 35)
    _GUICtrlListView_SetColumnWidth($idQueueList, 3, 180)
    _GUICtrlListView_SetColumnWidth($idQueueList, 4, 70)
    Local $bPreviewQueue = GUICtrlCreateButton("Résumé", 24, 488, 110, 30)
    Local $bClearQueue = GUICtrlCreateButton("Vider queue", 144, 488, 110, 30)
    Local $bCheckQueue = GUICtrlCreateButton("Tout cocher", 264, 488, 110, 30)

    Local $lblWorkflow = GUICtrlCreateLabel("Workflow conseillé", 405, 45, 240, 18)
    GUICtrlSetFont($lblWorkflow, 9, 700, 0, "Segoe UI")
    Local $bGenerateAll = _CreateHoverButton("1. Générer tous les BL/PDF", 405, 72, 330, 44)
    $idBtnRefreshResultsTop = _CreateHoverButton("2. Rafraîchir résultats", 405, 126, 330, 34)
    $idBtnForceNext = _CreateHoverButton("Forcer output / passer au suivant", 405, 170, 330, 36)
    GUICtrlSetState($idBtnForceNext, $GUI_DISABLE)
    $idBtnStopQueue = _CreateHoverButton("Stop queue après BL en cours", 405, 214, 330, 34)
    Local $lblForceHint = GUICtrlCreateLabel("À utiliser si ETMS a fini mais que l’attente outputs reste bloquée.", 405, 250, 330, 16)
    GUICtrlSetFont($lblForceHint, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor($lblForceHint, 0x64748B)
    GUICtrlSetBkColor($lblForceHint, 0xF3F6FB)
    Local $bUser = _CreateHoverButton("Utilisateur", 405, 270, 330, 30)
    $idLastBL = GUICtrlCreateLabel("Dernier BL : lecture...", 405, 310, 330, 42)
    GUICtrlSetColor($idLastBL, 0x344054)
    Local $lblStatus = GUICtrlCreateLabel("Résumé / progression", 405, 362, 240, 18)
    GUICtrlSetFont($lblStatus, 9, 700, 0, "Segoe UI")
    $idStatus = GUICtrlCreateEdit("Prêt.", 405, 384, 330, 103, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_READONLY))

    GUICtrlCreateTabItem("Résultats")
    Local $lblResults = GUICtrlCreateLabel("BL générés depuis snapshots", 24, 45, 260, 18)
    GUICtrlSetFont($lblResults, 9, 700, 0, "Segoe UI")
    $idResults = GUICtrlCreateListView("OK|BL|Entreprise|SP|Livraison", 24, 68, 485, 410, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_VSCROLL))
    _GUICtrlListView_SetExtendedListViewStyle($idResults, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES))
    _GUICtrlListView_SetColumnWidth($idResults, 0, 38)
    _GUICtrlListView_SetColumnWidth($idResults, 1, 135)
    _GUICtrlListView_SetColumnWidth($idResults, 2, 145)
    _GUICtrlListView_SetColumnWidth($idResults, 3, 105)
    _GUICtrlListView_SetColumnWidth($idResults, 4, 95)
    GUICtrlCreateLabel("Sélection", 535, 45, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    $idBtnRefreshResults = _CreateHoverButton("Rafraîchir", 535, 68, 95, 28)
    Local $bClearResults = GUICtrlCreateButton("Vider", 640, 68, 95, 28)
    Local $bCheckResults = GUICtrlCreateButton("Tout cocher", 535, 102, 95, 28)
    Local $bUncheckResults = GUICtrlCreateButton("Décocher", 640, 102, 95, 28)

    GUICtrlCreateLabel("BL / PDF", 535, 145, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    Local $bEditBL = _CreateHoverButton("Modifier / enregistrer", 535, 168, 200, 34)
    Local $bOpenPdf = _CreateHoverButton("Ouvrir PDF", 535, 210, 200, 30)

    GUICtrlCreateLabel("Actions", 535, 262, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    $idBtnMail = _CreateHoverButton("Créer mail", 535, 285, 200, 36)
    $idBtnEdocOnly = _CreateHoverButton("Upload EDOC seul", 535, 329, 200, 36)
    $idBtnEdocMail = _CreateHoverButton("EDOC + mail", 535, 373, 200, 40)
    Local $bSummaryResults = GUICtrlCreateButton("Résumé résultats", 535, 427, 200, 28)

    GUICtrlCreateLabel("Simple : coche les BL à traiter. Si rien n’est coché, l’action utilise la ligne sélectionnée. Modifier / enregistrer sauvegarde le HTML et recrée directement le PDF utilisé pour le mail/EDOC.", 24, 495, 710, 38)

    GUICtrlCreateTabItem("SP")
    GUICtrlCreateLabel("Transporteur / SP", 24, 45, 180, 18)
    $idSPCarrier = GUICtrlCreateCombo("", 24, 66, 300, 24, $CBS_DROPDOWN)
    GUICtrlSetData($idSPCarrier, _SP_ComboData(), "CARTAGE FLEX TRANSPORT")
    Local $bLoadSP = GUICtrlCreateButton("Charger", 340, 64, 85, 28)
    Local $bSaveSP = GUICtrlCreateButton("Sauver", 435, 64, 85, 28)
    Local $bDeleteSP = GUICtrlCreateButton("Supprimer", 530, 64, 95, 28)
    GUICtrlCreateLabel("To", 24, 102, 40, 18)
    $idSPTo = GUICtrlCreateInput("", 70, 99, 665, 22)
    GUICtrlCreateLabel("CC", 24, 129, 40, 18)
    $idSPCC = GUICtrlCreateInput("", 70, 126, 665, 22)
    GUICtrlCreateLabel("BCC", 24, 156, 40, 18)
    $idSPBCC = GUICtrlCreateInput("", 70, 153, 665, 22)
    GUICtrlCreateLabel("Objet", 24, 183, 45, 18)
    $idSPSubject = GUICtrlCreateInput("", 70, 180, 665, 22)
    GUICtrlCreateLabel("PDF", 24, 210, 45, 18)
    $idSPPdf = GUICtrlCreateInput("", 70, 207, 665, 22)
    GUICtrlCreateLabel("Corps mail - Entrée = saut de ligne", 24, 240, 420, 18)
    $idSPBody = GUICtrlCreateEdit("", 24, 262, 710, 105, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_WANTRETURN))
    GUICtrlCreateLabel("Signature additionnelle optionnelle", 24, 376, 560, 18)
    $idSPSign = GUICtrlCreateEdit("", 24, 398, 710, 70, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_WANTRETURN))
    GUICtrlCreateLabel("Variables : {J_SUBJECT}, {J_FILE}, {DELIVERY}, {DELIVERY_SHORT}, {DATE_JOUR}, {OPERATOR}, {CARRIER}, {NL}", 24, 478, 720, 18)
    $idSPKnown = GUICtrlCreateLabel("", 24, 502, 720, 20)
    GUICtrlCreateTabItem("")

    _SP_LoadToControls("CARTAGE FLEX TRANSPORT")
    _SP_UpdateKnownLabel()
    GUISetState(@SW_SHOW, $g_hGui)
    $g_iLastHoverTick = TimerInit()
    _LastBL_Refresh()

    While 1
        Local $msg = GUIGetMsg()
        _UpdateHoverButtons()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                If _HandleCloseRequest() Then ExitLoop
            Case $bAddQueue
                _QueueAddFromInput(False)
            Case $bImportQueue
                _QueueAddFromInput(True)
            Case $bRemoveQueue
                _QueueRemoveSelected()
            Case $bPreviewQueue
                _Status(_QueuePreviewText())
            Case $bClearQueue
                _QueueClearRows()
                GUICtrlSetData($idQueue, "")
                _Status("Queue vidée.")
            Case $bCheckQueue
                _QueueSetAllChecks(True)
            Case $bGenerateAll
                _RunGenerateQueue()
            Case $idBtnRefreshResultsTop, $idBtnRefreshResults
                _RefreshResultsList()
            Case $idBtnStopQueue
                $g_bStopQueue = True
                _Status("Stop demandé. Le BL en cours se termine puis la queue s’arrête.")
            Case $idBtnForceNext
                $g_bForceNext = True
                _Status("Forçage demandé : contrôle des outputs puis passage à l’étape suivante.")
            Case $bUser
                _ChangeUser()
            Case $bOpenPdf
                Local $idxOpen = _SelectedResultIndex()
                If $idxOpen >= 0 Then
                    _EnsureLatestPdfByIndex($idxOpen)
                    If FileExists($g_aPDF[$idxOpen]) Then ShellExecute($g_aPDF[$idxOpen])
                EndIf
            Case $bEditBL
                Local $idxEdit = _SelectedResultIndex()
                If $idxEdit >= 0 Then _OpenPreviewEditorByIndex($idxEdit)
            Case $idBtnMail
                _SendMailOnlySmart()
            Case $idBtnEdocOnly
                _UploadEdocOnlySmart()
            Case $idBtnEdocMail
                _SendEdocMailSmart()
            Case $bCheckResults
                _ResultsSetAllChecks(True)
            Case $bUncheckResults
                _ResultsSetAllChecks(False)
            Case $bSummaryResults
                _Status(_ResultsDetailedSummary())
            Case $bClearResults
                _ClearResults()
                _Status("Résultats vidés.")
            Case $bLoadSP
                _SP_LoadToControls(GUICtrlRead($idSPCarrier))
            Case $bSaveSP
                _SP_SaveFromControls()
                GUICtrlSetData($idSPCarrier, "")
                GUICtrlSetData($idSPCarrier, _SP_ComboData(), GUICtrlRead($idSPCarrier))
                _SP_UpdateKnownLabel()
                _Status("SP sauvegardé : " & GUICtrlRead($idSPCarrier))
            Case $bDeleteSP
                If MsgBox(52, $APP_TITLE, "Supprimer ce SP ?" & @CRLF & GUICtrlRead($idSPCarrier)) = 6 Then
                    _SP_Delete(GUICtrlRead($idSPCarrier))
                    GUICtrlSetData($idSPCarrier, "")
                    GUICtrlSetData($idSPCarrier, _SP_ComboData(), "Autre")
                    _SP_LoadToControls("Autre")
                    _SP_UpdateKnownLabel()
                EndIf
        EndSwitch
    WEnd
    GUIDelete($g_hGui)
EndFunc


Func _ChangeUser()
    Local $sNew = InputBox($APP_TITLE, "Nom utilisateur / Operator", $g_sOperator, "", 420, 140)
    If Not @error And StringStripWS($sNew, 3) <> "" Then
        $g_sOperator = StringStripWS($sNew, 3)
        IniWrite($INI_PATH, "USER", "Operator", $g_sOperator)
        _Status("Utilisateur sauvegardé : " & $g_sOperator)
    EndIf
EndFunc


Func _Status($txt)
    $g_sCMRStatus = $txt
    If $idStatus <> 0 Then GUICtrlSetData($idStatus, $txt)
    _AuditLog("CMR", StringLeft(StringReplace($txt, @CRLF, " | "), 500))
EndFunc


; ============================================================
; QUEUE + SNAPSHOT
; ============================================================

Func _RunGenerateQueue()
    Local $groups = _QueueGroups()
    If UBound($groups) = 0 Then
        _ModernInfo("Queue vide", "Colle au moins un J puis clique sur Ajouter BL ou Import auto.", "warning")
        Return False
    EndIf

    _ClearResults()
    $g_bStopQueue = False
    $g_bGenerating = True
    $g_bForceNext = False
    $g_bForceSkipCurrent = False
    $g_bHardClose = False
    _SetProcessingUi(True)

    For $i = 0 To UBound($groups) - 1
        If $g_bStopQueue Or $g_bHardClose Then ExitLoop
        Local $a = _QueueGroupToArray($groups[$i])
        If UBound($a) = 0 Then ContinueLoop

        $g_bForceNext = False
        $g_bForceSkipCurrent = False
        _QueueSetStatusByBlock($groups[$i], "En cours")
        _Status("Génération BL " & ($i + 1) & "/" & UBound($groups) & @CRLF & _JoinArray($a, " / "))

        Local $ok = _GenerateOneBL($a, $i + 1, UBound($groups))
        If $g_bHardClose Then ExitLoop

        If Not $ok Then
            If $g_bForceSkipCurrent Then
                _QueueSetStatusByBlock($groups[$i], "SKIP")
                _Status("BL ignoré par forçage utilisateur :" & @CRLF & _JoinArray($a, " / ") & @CRLF & @CRLF & "Passage au BL suivant.")
                _RefreshResultsList()
                ContinueLoop
            EndIf
            _QueueSetStatusByBlock($groups[$i], "KO")
            Local $r = _ModernErrorChoice("Erreur génération BL " & ($i + 1), _
                "BL concerné :" & @CRLF & _JoinArray($a, " / ") & @CRLF & @CRLF & "Que veux-tu faire ?", _
                "Continuer", "Relancer", "Arrêter")
            If $r = 3 Then ExitLoop
            If $r = 2 Then $i = $i - 1
        Else
            _QueueSetStatusByBlock($groups[$i], "OK")
        EndIf
        _RefreshResultsList()
    Next

    _SetProcessingUi(False)
    $g_bGenerating = False
    $g_bForceNext = False
    $g_bForceSkipCurrent = False

    If $g_bHardClose Then
        _Status("Fermeture demandée pendant traitement.")
        Return False
    EndIf

    _Status("Génération terminée." & @CRLF & _ResultsSummary() & @CRLF & "Chaque BL a son snapshot outputs.")
    Return True
EndFunc


Func _GenerateOneBL(ByRef $aNums, $index, $total)
    _ResetFiles()
    If Not _CreateInputFiles($aNums) Then Return False
    If Not _InputReady($aNums) Then Return False
    If Not _OpenEdsInEtms() Then Return False
    If Not _WaitOutputsReady($aNums, $index, $total) Then Return False

    Local $snap = _SnapshotOutputs($aNums, $index)
    If $snap = "" Then Return False

    Local $carrier = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 7)
    Local $delivery = _SnapGetDelivery($snap)
    Local $company = _SnapGetCompany($snap)
    Local $html = _GenerateHtmlFromSnapshot($aNums, $snap)
    If Not FileExists($html) Then Return False
    Local $pdf = _HtmlToPdf($html, $aNums, $carrier, $delivery)
    If Not FileExists($pdf) Then Return False
    _AddResult($aNums, $html, $pdf, $carrier, $delivery, $company, $snap)
    Return True
EndFunc


Func _WaitOutputsReady(ByRef $aNums, $index, $total)
    Local $t = TimerInit()
    While TimerDiff($t) < $OUTPUT_WAIT_MS
        _PumpGuiDuringProcess()
        If $g_bHardClose Then Return False
        If $g_bStopQueue Then Return False

        If $g_bForceNext Then
            $g_bForceNext = False
            If _OutputsOK() Then
                _Status("Forçage accepté." & @CRLF & "Les outputs existent : passage sans attendre la stabilité complète." & @CRLF & _JoinArray($aNums, " / "))
                Return True
            Else
                $g_bForceSkipCurrent = True
                _Status("Forçage demandé mais outputs incomplets." & @CRLF & "BL marqué SKIP puis passage au suivant :" & @CRLF & _JoinArray($aNums, " / "))
                Return False
            EndIf
        EndIf

        If _OutputsOK() And _OutputsStableFast() Then Return True
        _Status("BL " & $index & "/" & $total & " - attente outputs stables" & @CRLF & _JoinArray($aNums, " / ") & @CRLF & @CRLF & _
                "Si ETMS a terminé mais que l’outil reste bloqué, clique sur :" & @CRLF & "Forcer output / passer au suivant")
        Sleep(250)
    WEnd

    If _OutputsOK() Then
        _Status("Timeout atteint, mais outputs présents : passage à la suite.")
        Return True
    EndIf
    Return False
EndFunc


Func _OutputsStable()
    Local $s1 = _OutputsSignature()
    If $s1 = "" Then Return False
    Sleep($OUTPUT_STABLE_MS)
    Local $s2 = _OutputsSignature()
    Return ($s1 = $s2 And $s2 <> "")
EndFunc


Func _OutputsStableFast()
    Local $s1 = _OutputsSignature()
    If $s1 = "" Then Return False
    Local $t = TimerInit()
    While TimerDiff($t) < $OUTPUT_STABLE_MS
        _PumpGuiDuringProcess()
        If $g_bHardClose Or $g_bStopQueue Or $g_bForceNext Then Return False
        Sleep(100)
    WEnd
    Local $s2 = _OutputsSignature()
    Return ($s1 = $s2 And $s2 <> "")
EndFunc


Func _OutputsSignature()
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    Local $sig = ""
    For $i = 0 To 5
        Local $p = $OUTPUT_PATH & $a[$i]
        If Not FileExists($p) Or FileGetSize($p) <= 0 Then Return ""
        $sig &= $a[$i] & ":" & FileGetSize($p) & ":" & FileGetTime($p, 0, 1) & "|"
    Next
    Return $sig
EndFunc


Func _SnapshotOutputs(ByRef $aNums, $index)
    Local $name = @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_BL" & $index & "_" & _Safe(_JoinArray($aNums, "_"))
    Local $dir = $SNAPSHOT_PATH & $name & "\"
    DirCreate($dir)
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        Local $src = $OUTPUT_PATH & $a[$i]
        Local $dst = $dir & $a[$i]
        If Not FileExists($src) Or FileGetSize($src) <= 0 Then Return ""
        FileCopy($src, $dst, 1) ; PRO CLEAN : dossier déjà créé, overwrite simple plus rapide
        If Not FileExists($dst) Or FileGetSize($dst) <= 0 Then Return ""
    Next
    Return $dir
EndFunc


Func _QueueAddFromInput($bImportMulti)
    Local $txt = GUICtrlRead($idQueue)
    If StringStripWS($txt, 3) = "" Then
        MsgBox(48, $APP_TITLE, "Colle au moins un J dans la zone de gauche.")
        Return False
    EndIf
    Local $groups
    If $bImportMulti Then
        $groups = _QueueGroupsFromText($txt)
    Else
        Local $a = _QueueGroupToArray($txt)
        If UBound($a) = 0 Then Return False
        Local $one[1]
        $one[0] = _JoinArray($a, @LF)
        $groups = $one
    EndIf
    For $i = 0 To UBound($groups) - 1
        _QueueAppendBlock($groups[$i])
    Next
    GUICtrlSetData($idQueue, "")
    _QueueRefreshList()
    _Status("BL ajoutés dans la queue : " & UBound($groups))
    Return True
EndFunc


Func _QueueAppendBlock($sBlock)
    Local $a = _QueueGroupToArray($sBlock)
    If UBound($a) = 0 Then Return False
    Local $clean = _JoinArray($a, @LF)
    Local $n = UBound($g_aQueueBlocks)
    Local $nQS = UBound($g_aQueueStatus)
    ReDim $g_aQueueBlocks[$n + 1]
    ReDim $g_aQueueStatus[$n + 1]
    $g_aQueueBlocks[$n] = $clean
    $g_aQueueStatus[$n] = "Prêt"
    Return True
EndFunc


Func _QueueRefreshList()
    If $idQueueList = 0 Then Return False
    _GUICtrlListView_DeleteAllItems($idQueueList)
    For $i = 0 To UBound($g_aQueueBlocks) - 1
        Local $a = _QueueGroupToArray($g_aQueueBlocks[$i])
        _GUICtrlListView_AddItem($idQueueList, "", $i)
        _GUICtrlListView_AddSubItem($idQueueList, $i, $i + 1, 1)
        _GUICtrlListView_AddSubItem($idQueueList, $i, UBound($a), 2)
        _GUICtrlListView_AddSubItem($idQueueList, $i, _JoinArray($a, " / "), 3)
        _GUICtrlListView_AddSubItem($idQueueList, $i, $g_aQueueStatus[$i], 4)
        _GUICtrlListView_SetItemChecked($idQueueList, $i, True)
    Next
    Return True
EndFunc


Func _QueueRemoveSelected()
    Local $aSel = _GUICtrlListView_GetSelectedIndices($idQueueList, True)
    If Not IsArray($aSel) Or $aSel[0] = 0 Then
        MsgBox(48, $APP_TITLE, "Sélectionne une ligne dans la queue.")
        Return False
    EndIf
    Local $idx = Number($aSel[1])
    If $idx < 0 Or $idx >= UBound($g_aQueueBlocks) Then Return False
    For $i = $idx To UBound($g_aQueueBlocks) - 2
        $g_aQueueBlocks[$i] = $g_aQueueBlocks[$i + 1]
        $g_aQueueStatus[$i] = $g_aQueueStatus[$i + 1]
    Next
    ReDim $g_aQueueBlocks[UBound($g_aQueueBlocks) - 1]
    ReDim $g_aQueueStatus[UBound($g_aQueueStatus) - 1]
    _QueueRefreshList()
    Return True
EndFunc


Func _QueueClearRows()
    Local $nQB = UBound($g_aQueueBlocks)
    Local $nQS = UBound($g_aQueueStatus)
    ReDim $g_aQueueBlocks[0]
    ReDim $g_aQueueStatus[0]
    If $idQueueList <> 0 Then _GUICtrlListView_DeleteAllItems($idQueueList)
EndFunc


Func _QueueSetAllChecks($b)
    Local $count = _GUICtrlListView_GetItemCount($idQueueList)
    For $i = 0 To $count - 1
        _GUICtrlListView_SetItemChecked($idQueueList, $i, $b)
    Next
EndFunc


Func _QueueSetStatusByBlock($sBlock, $status)
    Local $aTarget = _QueueGroupToArray($sBlock)
    Local $target = _JoinArray($aTarget, @LF)
    For $i = 0 To UBound($g_aQueueBlocks) - 1
        If $g_aQueueBlocks[$i] = $target Then
            $g_aQueueStatus[$i] = $status
            If $idQueueList <> 0 Then _GUICtrlListView_SetItemText($idQueueList, $i, $status, 4)
            Return True
        EndIf
    Next
    Return False
EndFunc


Func _QueueGroups()
    If UBound($g_aQueueBlocks) > 0 Then
        Local $res[0]
        Local $hasChecked = False
        Local $i = 0

        For $i = 0 To UBound($g_aQueueBlocks) - 1
            If _GUICtrlListView_GetItemChecked($idQueueList, $i) Then
                $hasChecked = True
                ReDim $res[UBound($res) + 1]
                $res[UBound($res) - 1] = $g_aQueueBlocks[$i]
            EndIf
        Next

        If $hasChecked Then Return $res
        Return $g_aQueueBlocks
    EndIf

    Return _QueueGroupsFromText(GUICtrlRead($idQueue))
EndFunc


; ============================================================
; INPUT / ETMS / LIVE OUTPUT
; ============================================================

Func _LastBL_Refresh()
    If $idLastBL <> 0 Then GUICtrlSetData($idLastBL, _LastBL_Text())
    Return True
EndFunc


Func _InputReady(ByRef $aNums)
    Local $try = 0
    Local $last = ""
    While $try < 25
        $try += 1
        If FileExists($INPUT_CSV) And FileGetSize($INPUT_CSV) > 0 And FileExists($INPUT_GEN) And FileGetSize($INPUT_GEN) > 0 Then
            Local $sContent = FileRead($INPUT_CSV)
            Local $sGen = FileRead($INPUT_GEN)
            Local $first = _CleanJ($aNums[0])
            Local $tracking = $first & $g_sOperator
            Local $ok = True
            If Not StringInStr($sGen, $first & "," & $tracking) Then
                $ok = False
                $last = "inputGEN.csv ne contient pas : " & $first & "," & $tracking
            EndIf
            For $i = 0 To UBound($aNums) - 1
                Local $sJ = _CleanJ($aNums[$i])
                If $sJ <> "" And Not StringInStr($sContent, $sJ & "," & $tracking) Then
                    $ok = False
                    $last = "input.csv ne contient pas : " & $sJ & "," & $tracking
                    ExitLoop
                EndIf
            Next
            If $ok Then Return True
        Else
            $last = "Fichiers INPUT absents ou vides :" & @CRLF & $INPUT_CSV & " size=" & FileGetSize($INPUT_CSV) & @CRLF & $INPUT_GEN & " size=" & FileGetSize($INPUT_GEN)
        EndIf
        Sleep(250)
    WEnd
    _Status("INPUT KO après contrôle :" & @CRLF & $last)
    Return False
EndFunc


Func _OutputsOK()
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        Local $p = $OUTPUT_PATH & $a[$i]
        If Not FileExists($p) Then Return False
        If FileGetSize($p) <= 0 Then Return False
    Next
    Return True
EndFunc


Func _GetCarrierLive()
    Return _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 7)
EndFunc


Func _GetDeliveryLive()
    Local $d = _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 8)
    If _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 9) <> "" Then $d &= " " & _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 9)
    Return $d
EndFunc


Func _SetProcessingUi($bProcessing)
    If $idBtnForceNext <> 0 Then
        If $bProcessing Then
            GUICtrlSetState($idBtnForceNext, $GUI_ENABLE)
        Else
            GUICtrlSetState($idBtnForceNext, $GUI_DISABLE)
        EndIf
    EndIf
EndFunc


Func _PumpGuiDuringProcess()
    Local $msg = GUIGetMsg()
    _UpdateHoverButtons()
    Switch $msg
        Case 0
            Return
        Case $GUI_EVENT_CLOSE
            _HandleCloseRequest()
        Case $idBtnStopQueue
            $g_bStopQueue = True
            _Status("Stop demandé. Le BL en cours se termine puis la queue s’arrête.")
        Case $idBtnForceNext
            $g_bForceNext = True
            _Status("Forçage demandé : contrôle des outputs puis passage à l’étape suivante.")
        Case $idBtnMail
            _SendMailOnlySmart()
        Case $idBtnEdocOnly
            _UploadEdocOnlySmart()
        Case $idBtnEdocMail
            _SendEdocMailSmart()
        Case $idBtnRefreshResults, $idBtnRefreshResultsTop
            _RefreshResultsList()
    EndSwitch
EndFunc


Func _HandleCloseRequest()
    If Not $g_bGenerating Then Return True
    Local $choice = _ModernCloseDialog()
    Switch $choice
        Case 1
            Return False
        Case 2
            $g_bStopQueue = True
            _Status("Fermeture demandée : arrêt propre après le BL en cours.")
            Return False
        Case 3
            $g_bHardClose = True
            $g_bStopQueue = True
            Return True
    EndSwitch
    Return False
EndFunc


Func _ModernCloseDialog()
    Local $r = MsgBox(35 + 262144, $APP_TITLE, "Un BL est encore en cours." & @CRLF & @CRLF & "Oui = continuer" & @CRLF & "Non = arrêt propre après BL en cours" & @CRLF & "Annuler = fermer maintenant")
    If $r = 6 Then Return 1
    If $r = 7 Then Return 2
    Return 3
EndFunc


Func _ModernInfo($sTitle, $sMessage, $sType = "info")
    If $g_hGui = 0 Then
        _Status($sTitle & " - " & $sMessage)
        Return
    EndIf
    MsgBox(64, $sTitle, $sMessage)
EndFunc


Func _ModernErrorChoice($sTitle, $sMessage, $sBtn1, $sBtn2, $sBtn3)
    If $g_hGui = 0 Then
        _Status($sTitle & " - " & $sMessage & " / action auto: continuer")
        Return 1
    EndIf
    Local $r = MsgBox(51 + 262144, $sTitle, $sMessage & @CRLF & @CRLF & "Oui = " & $sBtn1 & @CRLF & "Non = " & $sBtn2 & @CRLF & "Annuler = " & $sBtn3)
    If $r = 6 Then Return 1
    If $r = 7 Then Return 2
    Return 3
EndFunc


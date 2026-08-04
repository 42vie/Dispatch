#Region ;**** AutoIt3Wrapper ****
#AutoIt3Wrapper_Outfile=HPE_BL_QUEUE_SNAPSHOT.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_UseUpx=n ; PRO CLEAN: UPX désactivé, souvent contre-productif sur ce type d'outil
#EndRegion

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>
#include <ComboConstants.au3>
#include <ListBoxConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <MsgBoxConstants.au3>
#include <Misc.au3>

Opt("GUIOnEventMode", 0)
Opt("MustDeclareVars", 1)

; ============================================================
; HPE BL QUEUE SNAPSHOT - VERSION PRO CLEAN OPTIMISÉE
; Base originale conservée, fonctions préservées.
; Optimisations majeures : cache CSV, throttling GUI hover, browser cache, snapshots plus rapides.
; ============================================================
; HPE BL QUEUE SNAPSHOT
; Correction majeure : chaque BL duplique ses outputs dans un dossier snapshot.
; Les HTML/PDF/mails sont générés depuis le snapshot du BL, jamais depuis OUTPUT live.
; ============================================================

Global Const $APP_TITLE = "HPE BL Queue Snapshot"
Global Const $BASE_PATH = "F:\Scripting\Export\EXPORT_HPE_CMR_001\"
Global Const $INPUT_PATH = $BASE_PATH & "Database\INPUT\"
Global Const $OUTPUT_PATH = $BASE_PATH & "Database\OUTPUT\"
Global Const $HTML_PATH = $BASE_PATH & "Database\HTML\"
Global Const $PDF_PATH = @ScriptDir & "\PDF\"
Global Const $SNAPSHOT_PATH = @ScriptDir & "\Snapshots\"
Global Const $INPUT_CSV = $INPUT_PATH & "input.csv"
Global Const $INPUT_GEN = $INPUT_PATH & "inputGEN.csv"
Global Const $LASTBL_MARKER = $INPUT_PATH & "last_bl_queue_snapshot.ini"
Global Const $EDS_PATH = $BASE_PATH & "EXPORT_HPE_CMR_006.eds"
Global Const $INI_PATH = @ScriptDir & "\Config\HPE_BL_QUEUE_SNAPSHOT.ini"
Global Const $MAIL_AUTOSEND = False
Global Const $OUTPUT_WAIT_MS = 180000
Global Const $OUTPUT_STABLE_MS = 1500

Global Const $ETMS_WINDOW = "[CLASS:TfmBrowser]"
Global Const $ETMS_TOOLBAR = "[CLASS:TRzToolbar; INSTANCE:1]"
Global Const $EDOC_UPLOAD_TITLE = "Upload Documents CDG"
Global Const $EDOC_UPLOAD_BTN = "TButton2"
Global Const $FILEOPEN_WIN = "[CLASS:TRzShellOpenSaveForm]"
Global Const $FILEOPEN_EDIT = "[CLASS:TRzEdit; INSTANCE:1]"
Global Const $TOOLBAR_X = 54
Global Const $TOOLBAR_Y = 9

Global $g_sOperator = ""
Global $g_bStopQueue = False
Global $g_aBL[0], $g_aHTML[0], $g_aPDF[0], $g_aCarrier[0], $g_aDelivery[0], $g_aCompany[0], $g_aSnap[0]
Global $g_aQueueBlocks[0], $g_aQueueStatus[0]
Global $idQueue = 0, $idQueueList = 0, $idStatus = 0, $idResults = 0, $idLastBL = 0
Global $idSPCarrier = 0, $idSPTo = 0, $idSPCC = 0, $idSPBCC = 0, $idSPSubject = 0, $idSPPdf = 0, $idSPBody = 0, $idSPSign = 0, $idSPKnown = 0
; ============================================================
; UX / PROCESS CONTROL + HOVER BUTTONS
; ============================================================
Global $g_hGui = 0
Global $g_bGenerating = False
Global $g_bForceNext = False
Global $g_bForceSkipCurrent = False
Global $g_bHardClose = False
Global $idBtnStopQueue = 0
Global $idBtnForceNext = 0
Global $idBtnMail = 0
Global $idBtnEdocOnly = 0
Global $idBtnEdocMail = 0
Global $idBtnRefreshResults = 0
Global $idBtnRefreshResultsTop = 0
Global $g_aHoverIds[0], $g_aHoverNormal[0], $g_aHoverHover[0], $g_aHoverIsHover[0]

; ============================================================
; PRO CLEAN / PERFORMANCE CACHE
; ============================================================
Global $g_oCsvCache = ObjCreate("Scripting.Dictionary") ; clé = chemin fichier + date/taille, valeur = tableau de lignes non vides
Global $g_sBrowserCache = ""
Global $g_iLastHoverTick = 0
Global Const $HOVER_THROTTLE_MS = 60


_Main()

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
    GUICtrlSetData($idStatus, $txt)
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

Func _AddResult(ByRef $aNums, $html, $pdf, $carrier, $delivery, $company, $snap)
    Local $n = UBound($g_aBL)
    ReDim $g_aBL[$n + 1]
    ReDim $g_aHTML[$n + 1]
    ReDim $g_aPDF[$n + 1]
    ReDim $g_aCarrier[$n + 1]
    ReDim $g_aDelivery[$n + 1]
    ReDim $g_aCompany[$n + 1]
    ReDim $g_aSnap[$n + 1]
    $g_aBL[$n] = _JoinArray($aNums, " / ")
    $g_aHTML[$n] = $html
    $g_aPDF[$n] = $pdf
    $g_aCarrier[$n] = $carrier
    $g_aDelivery[$n] = $delivery
    $g_aCompany[$n] = $company
    $g_aSnap[$n] = $snap
EndFunc

Func _ClearResults()
    ReDim $g_aBL[0]
    ReDim $g_aHTML[0]
    ReDim $g_aPDF[0]
    ReDim $g_aCarrier[0]
    ReDim $g_aDelivery[0]
    ReDim $g_aCompany[0]
    ReDim $g_aSnap[0]
    If $idResults <> 0 Then _GUICtrlListView_DeleteAllItems($idResults)
EndFunc

Func _CsvCacheClear($sPrefix = "")
    If Not IsObj($g_oCsvCache) Then Return False
    If $sPrefix = "" Then
        $g_oCsvCache.RemoveAll()
        Return True
    EndIf
    Local $aKeys = $g_oCsvCache.Keys
    For $i = 0 To UBound($aKeys) - 1
        If StringInStr($aKeys[$i], $sPrefix) = 1 Then $g_oCsvCache.Remove($aKeys[$i])
    Next
    Return True
EndFunc

Func _CsvLoadCached($file)
    If Not FileExists($file) Then Return SetError(1, 0, 0)
    Local $stamp = FileGetSize($file) & ":" & FileGetTime($file, 0, 1)
    Local $key = $file & "|" & $stamp
    If IsObj($g_oCsvCache) And $g_oCsvCache.Exists($key) Then Return $g_oCsvCache.Item($key)

    ; purge anciennes entrées du même fichier si le fichier a changé
    If IsObj($g_oCsvCache) Then
        Local $aKeys = $g_oCsvCache.Keys
        For $i = 0 To UBound($aKeys) - 1
            If StringInStr($aKeys[$i], $file & "|") = 1 Then $g_oCsvCache.Remove($aKeys[$i])
        Next
    EndIf

    Local $txt = FileRead($file)
    If @error Then Return SetError(2, 0, 0)
    $txt = StringReplace($txt, @CRLF, @LF)
    $txt = StringReplace($txt, @CR, @LF)
    Local $raw = StringSplit($txt, @LF, 2)
    Local $lines[0]
    For $i = 0 To UBound($raw) - 1
        If StringStripWS($raw[$i], 3) = "" Then ContinueLoop
        ReDim $lines[UBound($lines) + 1]
        $lines[UBound($lines) - 1] = $raw[$i]
    Next
    If IsObj($g_oCsvCache) Then $g_oCsvCache.Add($key, $lines)
    Return $lines
EndFunc

Func _RefreshResultsList()
    If $idResults = 0 Then Return False
    _GUICtrlListView_DeleteAllItems($idResults)
    For $i = 0 To UBound($g_aBL) - 1
        _GUICtrlListView_AddItem($idResults, "", $i)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aBL[$i], 1)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aCompany[$i], 2)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aCarrier[$i], 3)
        _GUICtrlListView_AddSubItem($idResults, $i, _ExtractDeliveryDateTime($g_aDelivery[$i]), 4)
        _GUICtrlListView_SetItemChecked($idResults, $i, False)
    Next
    Return True
EndFunc

Func _SelectedResultIndex()
    Local $aSel = _GUICtrlListView_GetSelectedIndices($idResults, True)
    If IsArray($aSel) And $aSel[0] > 0 Then
        Local $idx = Number($aSel[1])
        If $idx >= 0 And $idx < UBound($g_aBL) Then Return $idx
    EndIf
    MsgBox(48, $APP_TITLE, "Sélectionne une ligne dans l’onglet Résultats.")
    Return -1
EndFunc

Func _CheckedResultIndexes()
    Local $aIdx[0]
    For $i = 0 To UBound($g_aBL) - 1
        If _GUICtrlListView_GetItemChecked($idResults, $i) Then
            ReDim $aIdx[UBound($aIdx) + 1]
            $aIdx[UBound($aIdx) - 1] = $i
        EndIf
    Next
    Return $aIdx
EndFunc

Func _ActionIndexes()
    Local $aIdx = _CheckedResultIndexes()
    If UBound($aIdx) > 0 Then Return $aIdx
    Local $idx = _SelectedResultIndex()
    If $idx < 0 Then Return $aIdx
    ReDim $aIdx[1]
    $aIdx[0] = $idx
    Return $aIdx
EndFunc

Func _ResultsSummary()
    Return "BL générés : " & UBound($g_aBL)
EndFunc

Func _OpenPreviewEditorByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    _OpenPreviewEditor($g_aHTML[$idx], $a, $idx)
    Return True
EndFunc

Func _RebuildPdfByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _HtmlToPdf($g_aHTML[$idx], $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
    If FileExists($pdf) Then
        $g_aPDF[$idx] = $pdf
        ShellExecute($pdf)
        _Status("PDF recréé depuis HTML sélectionné :" & @CRLF & $pdf)
        Return True
    EndIf
    Return False
EndFunc

Func _CreateMailForResult($idx)
    If $idx < 0 Or $idx >= UBound($g_aPDF) Then Return ""
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _EnsureLatestPdfByIndex($idx)
    If FileExists($pdf) Then
        _CreateMailSP($pdf, $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
        _Status("Mail créé avec PDF :" & @CRLF & $g_aBL[$idx])
        Return $pdf
    EndIf
    MsgBox(48, $APP_TITLE, "PDF introuvable pour ce BL." & @CRLF & $g_aBL[$idx])
    Return ""
EndFunc

Func _UploadEdocForResult($idx)
    If $idx < 0 Or $idx >= UBound($g_aPDF) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = $g_aPDF[$idx]
    If Not FileExists($pdf) Then $pdf = _EnsureLatestPdfByIndex($idx)
    If FileExists($pdf) Then
        Local $ok = _PrintPdfToEdoc($pdf, $a)
        If $ok Then
            _Status("EDOC upload OK :" & @CRLF & $g_aBL[$idx])
            Return True
        EndIf
    EndIf
    _Status("EDOC upload KO :" & @CRLF & $g_aBL[$idx])
    Return False
EndFunc

Func _SendMailOnlyResult($idx)
    Return (_CreateMailForResult($idx) <> "")
EndFunc

Func _SendOneResult($idx)
    ; Compatibilité : mail d'abord, puis EDOC pour ce BL.
    If _CreateMailForResult($idx) = "" Then Return False
    Return _UploadEdocForResult($idx)
EndFunc

Func _SendMailOnlySmart()
    Local $aIdx = _ActionIndexes()
    If UBound($aIdx) = 0 Then Return False
    For $i = 0 To UBound($aIdx) - 1
        _CreateMailForResult($aIdx[$i])
    Next
    _Status("Mails créés avec PDF : " & UBound($aIdx))
    Return True
EndFunc

Func _UploadEdocOnlySmart()
    Local $aIdx = _ActionIndexes()
    If UBound($aIdx) = 0 Then Return False
    Local $okCount = 0
    For $i = 0 To UBound($aIdx) - 1
        If _UploadEdocForResult($aIdx[$i]) Then $okCount += 1
        Sleep(250)
    Next
    _Status("Upload EDOC seul terminé : " & $okCount & "/" & UBound($aIdx))
    Return ($okCount > 0)
EndFunc

Func _ResultsSetAllChecks($b)
    If $idResults = 0 Then Return False
    Local $count = _GUICtrlListView_GetItemCount($idResults)
    For $i = 0 To $count - 1
        _GUICtrlListView_SetItemChecked($idResults, $i, $b)
    Next
    Return True
EndFunc

Func _ResultsDetailedSummary()
    If UBound($g_aBL) = 0 Then Return "Aucun résultat BL."
    Local $s = "Résultats BL : " & UBound($g_aBL) & @CRLF & @CRLF
    For $i = 0 To UBound($g_aBL) - 1
        $s &= "#" & ($i + 1) & " - " & $g_aBL[$i] & @CRLF
        $s &= "Entreprise : " & $g_aCompany[$i] & @CRLF
        $s &= "SP : " & $g_aCarrier[$i] & @CRLF
        $s &= "Livraison : " & _ExtractDeliveryDateTime($g_aDelivery[$i]) & @CRLF
        $s &= "PDF : " & $g_aPDF[$i] & @CRLF & @CRLF
    Next
    Return $s
EndFunc

Func _SendEdocMailSmart()
    Local $aIdx = _ActionIndexes()
    If UBound($aIdx) = 0 Then Return False
    ; Étape 1 : créer tous les mails avec le dernier PDF.
    For $i = 0 To UBound($aIdx) - 1
        _CreateMailForResult($aIdx[$i])
        Sleep(150)
    Next
    ; Étape 2 : uploader les PDFs un par un dans EDOC, avec scanner anti-conflit.
    For $i = 0 To UBound($aIdx) - 1
        _UploadEdocForResult($aIdx[$i])
        Sleep(250)
    Next
    _Status("EDOC + mails terminé. Mails créés d’abord, puis uploads EDOC un par un : " & UBound($aIdx))
    Return True
EndFunc

Func _SendAllResults()
    If UBound($g_aPDF) = 0 Then
        MsgBox(48, $APP_TITLE, "Aucun BL généré.")
        Return False
    EndIf
    For $i = 0 To UBound($g_aPDF) - 1
        _SendOneResult($i)
    Next
    _Status("EDOC + mails lancés pour tous les BL générés.")
    Return True
EndFunc

Func _StringToArrayBL($s)
    Local $parts = StringSplit(StringReplace($s, " / ", @LF), @LF, 2)
    Local $a[0]
    For $i = 0 To UBound($parts) - 1
        Local $j = _CleanJ($parts[$i])
        If $j = "" Then ContinueLoop
        ReDim $a[UBound($a) + 1]
        $a[UBound($a) - 1] = $j
    Next
    Return $a
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

Func _QueueGroupsFromText($txt)
    $txt = StringReplace($txt, @CRLF, @LF)
    $txt = StringReplace($txt, @CR, @LF)
    Local $raw = StringSplit($txt, @LF, 2)
    Local $groups[0]
    Local $cur = ""
    For $i = 0 To UBound($raw) - 1
        Local $line = StringStripWS($raw[$i], 3)
        If $line = "" Or $line = "---" Then
            If StringStripWS($cur, 3) <> "" Then
                ReDim $groups[UBound($groups) + 1]
                $groups[UBound($groups) - 1] = $cur
                $cur = ""
            EndIf
        Else
            If $cur <> "" Then $cur &= @LF
            $cur &= $line
        EndIf
    Next
    If StringStripWS($cur, 3) <> "" Then
        ReDim $groups[UBound($groups) + 1]
        $groups[UBound($groups) - 1] = $cur
    EndIf
    Return $groups
EndFunc

Func _QueueGroupToArray($s)
    Local $norm = StringReplace(StringReplace($s, " / ", @LF), ";", @LF)
    Local $raw = StringSplit($norm, @LF, 2)
    Local $a[0]
    For $i = 0 To UBound($raw) - 1
        Local $j = _CleanJ($raw[$i])
        If $j = "" Then ContinueLoop
        ReDim $a[UBound($a) + 1]
        $a[UBound($a) - 1] = $j
    Next
    Return $a
EndFunc

Func _QueuePreviewText()
    Local $g = _QueueGroups()
    Local $s = ""
    If UBound($g) = 0 Then Return "Queue vide. Colle des J puis clique Ajouter BL ou Import auto."
    For $i = 0 To UBound($g) - 1
        Local $a = _QueueGroupToArray($g[$i])
        $s &= "BL " & ($i + 1) & " : " & _JoinArray($a, " / ") & @CRLF
    Next
    Return $s
EndFunc

; ============================================================
; SP
; ============================================================
Func _Enc($s)
    $s = StringReplace($s, @CRLF, "{BR}")
    $s = StringReplace($s, @CR, "{BR}")
    $s = StringReplace($s, @LF, "{BR}")
    Return $s
EndFunc

Func _Dec($s)
    Return StringReplace($s, "{BR}", @CRLF)
EndFunc

Func _SP_InitDefaultsIfNeeded()
    If IniRead($INI_PATH, "SP:CARTAGE FLEX TRANSPORT", "TO", "") <> "" Then Return
    _SP_WriteRule("CARTAGE FLEX TRANSPORT", "sensible@flex-transport.com", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Hello{NL}{NL}Pour le {DELIVERY_SHORT}", "Regards,{NL}{NL}{OPERATOR}")
    _SP_WriteRule("European Fret Distribution Services", "exploitation95@efds.eu", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance.", "Cordialement,{NL}{NL}{OPERATOR}")
    _SP_WriteRule("Transports Groussard SA", "expeditors@tps-groussard.fr", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance.", "Cordialement,{NL}{NL}{OPERATOR}")
    _SP_WriteRule("Autre", "", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance.", "Cordialement,{NL}{NL}{OPERATOR}")
EndFunc

Func _SP_Sec($carrier)
    Return "SP:" & StringUpper(StringStripWS($carrier, 3))
EndFunc

Func _SP_WriteRule($carrier, $to, $cc, $bcc, $subject, $pdf, $body, $signature)
    Local $sec = _SP_Sec($carrier)
    IniWrite($INI_PATH, $sec, "CARRIER", $carrier)
    IniWrite($INI_PATH, $sec, "TO", $to)
    IniWrite($INI_PATH, $sec, "CC", $cc)
    IniWrite($INI_PATH, $sec, "BCC", $bcc)
    IniWrite($INI_PATH, $sec, "SUBJECT", $subject)
    IniWrite($INI_PATH, $sec, "PDF", $pdf)
    IniWrite($INI_PATH, $sec, "BODY", _Enc($body))
    IniWrite($INI_PATH, $sec, "SIGNATURE", _Enc($signature))
EndFunc

Func _SP_ComboData()
    Local $sections = IniReadSectionNames($INI_PATH)
    If @error Then Return "CARTAGE FLEX TRANSPORT|European Fret Distribution Services|Transports Groussard SA|Autre"
    Local $s = ""
    For $i = 1 To $sections[0]
        If StringLeft($sections[$i], 3) = "SP:" Then
            Local $c = IniRead($INI_PATH, $sections[$i], "CARRIER", StringTrimLeft($sections[$i], 3))
            If $s <> "" Then $s &= "|"
            $s &= $c
        EndIf
    Next
    Return $s
EndFunc

Func _SP_LoadToControls($carrier)
    Local $sec = _SP_Sec($carrier)
    GUICtrlSetData($idSPCarrier, IniRead($INI_PATH, $sec, "CARRIER", $carrier))
    GUICtrlSetData($idSPTo, IniRead($INI_PATH, $sec, "TO", ""))
    GUICtrlSetData($idSPCC, IniRead($INI_PATH, $sec, "CC", ""))
    GUICtrlSetData($idSPBCC, IniRead($INI_PATH, $sec, "BCC", ""))
    GUICtrlSetData($idSPSubject, IniRead($INI_PATH, $sec, "SUBJECT", "{J_SUBJECT}"))
    GUICtrlSetData($idSPPdf, IniRead($INI_PATH, $sec, "PDF", "{J_FILE}_Delivery Order"))
    GUICtrlSetData($idSPBody, _Dec(IniRead($INI_PATH, $sec, "BODY", _Enc("Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance."))))
    GUICtrlSetData($idSPSign, _Dec(IniRead($INI_PATH, $sec, "SIGNATURE", _Enc("Cordialement,{NL}{NL}{OPERATOR}"))))
EndFunc

Func _SP_SaveFromControls()
    Local $carrier = StringStripWS(GUICtrlRead($idSPCarrier), 3)
    If $carrier = "" Then Return False
    _SP_WriteRule($carrier, GUICtrlRead($idSPTo), GUICtrlRead($idSPCC), GUICtrlRead($idSPBCC), GUICtrlRead($idSPSubject), GUICtrlRead($idSPPdf), GUICtrlRead($idSPBody), GUICtrlRead($idSPSign))
    Return True
EndFunc

Func _SP_Delete($carrier)
    IniDelete($INI_PATH, _SP_Sec($carrier))
EndFunc

Func _SP_UpdateKnownLabel()
    GUICtrlSetData($idSPKnown, "SP connus : " & StringReplace(_SP_ComboData(), "|", " / "))
EndFunc

Func _SP_ReadRule($carrier)
    Local $sec = _SP_MatchSection($carrier)
    Local $r = ObjCreate("Scripting.Dictionary")
    $r.Add("CARRIER", IniRead($INI_PATH, $sec, "CARRIER", $carrier))
    $r.Add("TO", IniRead($INI_PATH, $sec, "TO", ""))
    $r.Add("CC", IniRead($INI_PATH, $sec, "CC", ""))
    $r.Add("BCC", IniRead($INI_PATH, $sec, "BCC", ""))
    $r.Add("SUBJECT", IniRead($INI_PATH, $sec, "SUBJECT", "{J_SUBJECT}"))
    $r.Add("PDF", IniRead($INI_PATH, $sec, "PDF", "{J_FILE}_Delivery Order"))
    $r.Add("BODY", _Dec(IniRead($INI_PATH, $sec, "BODY", "")))
    $r.Add("SIGNATURE", _Dec(IniRead($INI_PATH, $sec, "SIGNATURE", "")))
    Return $r
EndFunc

Func _SP_MatchSection($carrier)
    Local $c = StringUpper(StringStripWS($carrier, 3))
    Local $sections = IniReadSectionNames($INI_PATH)
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 3) = "SP:" Then
                Local $ruleCarrier = IniRead($INI_PATH, $sections[$i], "CARRIER", "")
                If $ruleCarrier <> "" And StringInStr($c, StringUpper($ruleCarrier)) Then Return $sections[$i]
            EndIf
        Next
    EndIf
    Return "SP:AUTRE"
EndFunc

Func _SP_GetRule($carrier)
    Return _SP_ReadRule($carrier)
EndFunc

Func _SP_BuildVariables(ByRef $aNums, $delivery, $carrier)
    Local $v = ObjCreate("Scripting.Dictionary")
    $v.Add("J_SUBJECT", _JoinArray($aNums, " - "))
    $v.Add("J_FILE", _JoinArray($aNums, "_"))
    $v.Add("DELIVERY", $delivery)
    $v.Add("DELIVERY_SHORT", _ExtractDeliveryDateTime($delivery))
    $v.Add("DATE_JOUR", @MDAY & "/" & @MON & "/" & @YEAR)
    $v.Add("OPERATOR", $g_sOperator)
    $v.Add("CARRIER", $carrier)
    $v.Add("NL", @CRLF)
    Return $v
EndFunc

Func _SP_ApplyTemplate($tpl, ByRef $vars)
    Local $s = $tpl
    For $k In $vars.Keys
        $s = StringReplace($s, "{" & $k & "}", $vars.Item($k))
    Next
    Return $s
EndFunc

; ============================================================
; INPUT / ETMS / LIVE OUTPUT
; ============================================================
Func _NowText()
    Return @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc

Func _LastBL_Write(ByRef $aNums)
    If Not FileExists($INPUT_PATH) Then DirCreate($INPUT_PATH)
    IniWrite($LASTBL_MARKER, "LAST", "Date", _NowText())
    IniWrite($LASTBL_MARKER, "LAST", "Operator", $g_sOperator)
    IniWrite($LASTBL_MARKER, "LAST", "WindowsUser", @UserName)
    IniWrite($LASTBL_MARKER, "LAST", "Machine", @ComputerName)
    IniWrite($LASTBL_MARKER, "LAST", "BL", _JoinArray($aNums, " / "))
    IniWrite($LASTBL_MARKER, "LAST", "InputCSV", $INPUT_CSV)
    IniWrite($LASTBL_MARKER, "LAST", "InputGEN", $INPUT_GEN)
    _LastBL_Refresh()
    Return True
EndFunc

Func _LastBL_Text()
    If Not FileExists($LASTBL_MARKER) Then Return "Dernier BL : aucune info"
    Local $d = IniRead($LASTBL_MARKER, "LAST", "Date", "")
    Local $op = IniRead($LASTBL_MARKER, "LAST", "Operator", "")
    Local $wu = IniRead($LASTBL_MARKER, "LAST", "WindowsUser", "")
    Local $pc = IniRead($LASTBL_MARKER, "LAST", "Machine", "")
    Local $bl = IniRead($LASTBL_MARKER, "LAST", "BL", "")
    If $op = "" Then $op = $wu
    If $pc <> "" Then $pc = " / " & $pc
    If $bl = "" Then Return "Dernier BL : " & $op & $pc & " - " & $d
    Return "Dernier BL : " & $bl & @CRLF & "Par " & $op & $pc & " le " & $d
EndFunc

Func _LastBL_Refresh()
    If $idLastBL <> 0 Then GUICtrlSetData($idLastBL, _LastBL_Text())
    Return True
EndFunc

Func _ResetFiles()
    _CsvCacheClear() ; PRO CLEAN : évite toute donnée CSV ancienne en mémoire
    If FileExists($INPUT_CSV) Then FileDelete($INPUT_CSV)
    If FileExists($INPUT_GEN) Then FileDelete($INPUT_GEN)
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        If FileExists($OUTPUT_PATH & $a[$i]) Then FileDelete($OUTPUT_PATH & $a[$i])
    Next
EndFunc

Func _CreateInputFiles(ByRef $aNums)
    Local $first = _CleanJ($aNums[0])
    If $first = "" Then
        _Status("INPUT KO : premier J vide.")
        Return False
    EndIf

    Local $tracking = $first & $g_sOperator
    Local $csv = ""
    For $i = 0 To UBound($aNums) - 1
        Local $sJ = _CleanJ($aNums[$i])
        If $sJ <> "" Then $csv &= $sJ & "," & $tracking & @CRLF
    Next

    If StringStripWS($csv, 3) = "" Then
        _Status("INPUT KO : aucun J valide à écrire.")
        Return False
    EndIf

    If Not _WriteAnsi($INPUT_CSV, $csv) Then
        _Status("INPUT KO : impossible de créer" & @CRLF & $INPUT_CSV)
        Return False
    EndIf

    If Not _WriteAnsi($INPUT_GEN, $first & "," & $tracking & @CRLF) Then
        _Status("INPUT KO : impossible de créer" & @CRLF & $INPUT_GEN)
        Return False
    EndIf

    _LastBL_Write($aNums)
    _Status("INPUT OK :" & @CRLF & $INPUT_CSV & @CRLF & $INPUT_GEN & @CRLF & _LastBL_Text())
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

; ============================================================
; HTML / PDF FROM SNAPSHOT
; ============================================================
Func _OpenPreviewEditor($htmlFile, ByRef $aNums, $idx)
    If Not FileExists($htmlFile) Then Return False
    Local $hPrev = GUICreate("Modifier BL - " & _JoinArray($aNums, " / "), 1120, 820, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_SIZEBOX))
    GUICtrlCreateLabel("Modifie le BL puis clique Enregistrer. Enregistrer sauvegarde le HTML ET régénère le PDF utilisé pour le mail.", 10, 8, 970, 22)
    Local $bSave = GUICtrlCreateButton("Enregistrer", 10, 35, 150, 32)
    Local $bPdfOpen = GUICtrlCreateButton("Ouvrir PDF", 170, 35, 140, 32)
    Local $bHtml = GUICtrlCreateButton("Ouvrir HTML", 320, 35, 140, 32)
    Local $bClose = GUICtrlCreateButton("Fermer", 970, 35, 130, 32)
    Local $oIE = ObjCreate("Shell.Explorer.2")
    If Not IsObj($oIE) Then
        MsgBox(16+262144, $APP_TITLE, "Contrôle aperçu HTML indisponible.")
        GUIDelete($hPrev)
        Return False
    EndIf
    GUICtrlCreateObj($oIE, 10, 75, 1090, 730)
    $oIE.Silent = True
    $oIE.Navigate("file:///" & StringReplace($htmlFile, "\", "/"))
    While $oIE.Busy Or $oIE.ReadyState <> 4
        Sleep(100)
    WEnd
    GUISetState(@SW_SHOW, $hPrev)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $bClose
                ExitLoop
            Case $bSave
                Local $pdfSaved = _SavePreviewAndRebuildPdf($oIE, $aNums, $idx, True)
                If FileExists($pdfSaved) Then
                    MsgBox(64, $APP_TITLE, "Modifications enregistrées et PDF régénéré." & @CRLF & @CRLF & $pdfSaved)
                Else
                    MsgBox(16, $APP_TITLE, "Impossible d’enregistrer ou de régénérer le PDF.")
                EndIf
            Case $bHtml
                ShellExecute($htmlFile)
            Case $bPdfOpen
                If $idx >= 0 Then
                    _EnsureLatestPdfByIndex($idx)
                    If $idx < UBound($g_aPDF) And FileExists($g_aPDF[$idx]) Then ShellExecute($g_aPDF[$idx])
                EndIf
        EndSwitch
    WEnd
    GUIDelete($hPrev)
    Return True
EndFunc

Func _SaveEditedPreview(ByRef $oIE, ByRef $aNums)
    Local $html = ""
    If IsObj($oIE) And IsObj($oIE.Document) Then $html = "<!doctype html>" & $oIE.Document.documentElement.outerHTML
    If $html = "" Then Return ""
    Local $name = "Printed_Delivery_Order_" & _Safe(_JoinArray($aNums, "_")) & "_MODIFIE_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".html"
    Local $file = $HTML_PATH & $name
    If Not _WriteUtf8($file, $html) Then
        Local $fallback = @ScriptDir & "\HTML\"
        If Not FileExists($fallback) Then DirCreate($fallback)
        $file = $fallback & $name
        If Not _WriteUtf8($file, $html) Then Return ""
    EndIf
    Return $file
EndFunc

Func _SavePreviewAndRebuildPdf(ByRef $oIE, ByRef $aNums, $idx, $bOpenPdf = False)
    Local $edited = _SaveEditedPreview($oIE, $aNums)
    If Not FileExists($edited) Then Return ""
    Local $carrier = ""
    Local $delivery = ""
    If $idx >= 0 And $idx < UBound($g_aCarrier) Then
        $carrier = $g_aCarrier[$idx]
        $delivery = $g_aDelivery[$idx]
    EndIf
    Local $pdf = _HtmlToPdf($edited, $aNums, $carrier, $delivery)
    If FileExists($pdf) Then
        If $idx >= 0 And $idx < UBound($g_aPDF) Then
            $g_aHTML[$idx] = $edited
            $g_aPDF[$idx] = $pdf
        EndIf
        _Status("Modifications enregistrées + PDF régénéré :" & @CRLF & $pdf)
        If $bOpenPdf Then ShellExecute($pdf)
        Return $pdf
    EndIf
    Return ""
EndFunc

Func _EnsureLatestPdfByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return ""
    If Not FileExists($g_aHTML[$idx]) Then Return ""
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _HtmlToPdf($g_aHTML[$idx], $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
    If FileExists($pdf) Then
        $g_aPDF[$idx] = $pdf
        Return $pdf
    EndIf
    Return $g_aPDF[$idx]
EndFunc

Func _GenerateHtmlFromSnapshot(ByRef $aNums, $snap)
    Local $rows = "", $totalPcs = 0, $totalWeight = 0, $totalVol = 0
    For $r = 0 To 39
        Local $j = _SnapCsvGet($snap, "outputDIMS.csv", $r, 0)
        If $j = "" Then ContinueLoop
        Local $dim = _SnapCsvGet($snap, "outputDIMS.csv", $r, 1)
        Local $pcs = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 1)
        Local $wgt = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 2)
        Local $vol = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 3)
        Local $packid = _JoinUniqueColsSnap($snap, "outputPACKID.csv", $r, 2, 12, " - ")
        Local $ref = _JoinUniqueColsSnap($snap, "outputREFS.csv", $r, 2, 12, " - ")
        $rows &= "<tr><td contenteditable='true'>" & _H($j) & "</td><td contenteditable='true'>" & _H($ref) & "</td><td contenteditable='true'>" & _H($packid) & "</td><td contenteditable='true'>" & _H($pcs) & "</td><td contenteditable='true'>" & _H($wgt) & "</td><td contenteditable='true'>" & _H($dim) & "</td></tr>"
        $totalPcs += Number(_CleanNum($pcs))
        $totalWeight += Number(_CleanNum($wgt))
        $totalVol += Number(_CleanNum($vol))
    Next
    Local $deliver = _FormatDeliverTo(_DefaultDeliverToSnap($snap))
    Local $ready = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 1)
    Local $delivDate = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 5)
    Local $delivTime = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 6)
    Local $carrier = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 7)
    Local $instr = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 8)
    If _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9) <> "" Then $instr &= " " & _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9)
    $instr = _FormatDeliveryInstructions($instr)
    Local $first = _CleanJ($aNums[0])
    Local $tracking = $first & $g_sOperator
    Local $html = _HtmlTemplate($tracking, $g_sOperator, $deliver, $ready, $delivDate, $delivTime, $carrier, $instr, $rows, _TrimNum($totalPcs), _TrimNum($totalWeight), _TrimNum($totalVol))
    Local $name = "Printed_Delivery_Order_" & _Safe(_JoinArray($aNums, "_")) & "_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".html"
    Local $file = $HTML_PATH & $name
    If Not _WriteUtf8($file, $html) Then
        Local $fallback = @ScriptDir & "\HTML\"
        If Not FileExists($fallback) Then DirCreate($fallback)
        $file = $fallback & $name
        If Not _WriteUtf8($file, $html) Then Return ""
    EndIf
    Return $file
EndFunc

Func _DefaultDeliverToSnap($snap)
    Local $s = _SnapCsvGet($snap, "outputGEN.csv", 0, 1)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 2) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 2)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 3) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 3)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 4) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 4)
    Return $s
EndFunc

Func _SnapGetCompany($snap)
    Local $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 1)
    If $company = "" Then $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 2)
    If $company = "" Then $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 3)
    $company = StringReplace($company, @CRLF, " ")
    $company = StringReplace($company, @CR, " ")
    $company = StringReplace($company, @LF, " ")
    While StringInStr($company, "  ")
        $company = StringReplace($company, "  ", " ")
    WEnd
    Return StringStripWS($company, 3)
EndFunc

Func _SnapGetDelivery($snap)
    Local $d = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 8)
    If _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9) <> "" Then $d &= " " & _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9)
    Return $d
EndFunc

Func _FormatDeliveryInstructions($s)
    Local $x = StringStripWS($s, 3)
    Local $rep = "$1" & @CRLF
    $x = StringRegExpReplace($x, "(?i)(entre\s+[0-9]{1,2}h(?:[0-9]{0,2})?\s+et\s+[0-9]{1,2}h(?:[0-9]{0,2})?)\s+", $rep)
    Return $x
EndFunc

Func _FormatDeliverTo($s)
    $s = StringReplace($s, " Tél", @CRLF & "Tél")
    $s = StringReplace($s, " Tel", @CRLF & "Tel")
    $s = StringReplace($s, " TEL", @CRLF & "TEL")
    $s = StringReplace($s, " Phone", @CRLF & "Phone")
    $s = StringReplace($s, " Mobile", @CRLF & "Mobile")
    $s = StringReplace($s, " +33", @CRLF & "+33")
    Return _H($s)
EndFunc

Func _HtmlTemplate($tracking, $operator, $deliver, $ready, $delivDate, $delivTime, $carrier, $instr, $rows, $totalPcs, $totalWeight, $totalVol)
    Local $date = @MDAY & "/" & @MON & "/" & @YEAR
    Local $s = "<!doctype html><html lang='fr'><head><meta charset='utf-8'><title>DELIVERY ORDER - " & _H($tracking) & "</title>" & _Css() & "</head><body>"
    $s &= "<main class='page'>"
    $s &= "<section class='header'><div class='logo'><img src='" & _LogoDataUri() & "'></div><div class='head-right'><div class='doc-title'>DELIVERY ORDER</div><div class='datebox'><span class='label'>Date :</span> " & _H($date) & "</div></div></section>"
    $s &= "<section class='topline'><div class='line'><span class='label'>Tracking# :</span><span class='value'>" & _H($tracking) & "</span></div><div class='line'><span class='label'>Operator :</span><span class='value'>" & _H($operator) & "</span></div></section>"
    $s &= "<section class='boxgrid'><div class='box'><div class='boxtitle'>DELIVER TO:</div><div class='addr' contenteditable='true'>" & $deliver & "</div></div><div class='box'><div class='boxtitle'>Pick Up From:</div><div contenteditable='true'>Expeditors International France SAS<br>La Porte des Champs<br>Batiment B - Quais 21 a 27<br>95470 Survilliers CEDEX</div></div></section>"
    $s &= "<section class='sectiongrid'><div class='info'><span class='label'>DELIVERING CARRIER:</span><br><span contenteditable='true'>" & _H($carrier) & "</span></div><div class='info'><span class='label'>Commodity:</span><br><span contenteditable='true'>IT spare parts</span></div></section>"
    $s &= "<section class='schedule'><div><span class='label'>Ready:</span><br><span contenteditable='true'>" & _H($ready) & "</span></div><div><span class='label'>Delivery date:</span><br><span contenteditable='true'>" & _H($delivDate) & "</span></div><div><span class='label'>Delivery Time:</span><br><span contenteditable='true'>" & _H($delivTime) & "</span></div></section>"
    $s &= "<section class='instructions'><span class='label'>Delivery Instructions:</span><br><div contenteditable='true'>" & _H($instr) & "</div></section>"
    $s &= "<table><thead><tr><th>Expeditors Ref</th><th>W/o &amp; S/o Ref</th><th>Pack ID</th><th>Pcs count</th><th>Weight</th><th>Dims</th></tr></thead><tbody>" & $rows
    $s &= "<tr class='total'><td>Total:</td><td></td><td></td><td>" & _H($totalPcs) & "</td><td>" & _H($totalWeight) & "</td><td>" & _H($totalVol) & "</td></tr></tbody></table>"
    $s &= "<section class='signatures'><div><div class='sigline'>X</div><div class='sublabel'>(signature)</div><div class='sigline second'>X</div><div class='sublabel'>(name)</div></div><div class='received'><div class='received-row'><span>DATE RECEIVED:</span><span class='received-line'></span></div><div class='received-row'><span>TIME RECEIVED:</span><span class='received-line'></span></div></div></section>"
    $s &= "</main></body></html>"
    Return $s
EndFunc

Func _Css()
    Return "<style>*{box-sizing:border-box}body{margin:0;background:#e9e9e9;color:#111;font-family:Calibri,Arial,sans-serif;font-size:11pt}.page{width:210mm;min-height:297mm;margin:0 auto;background:#fff;padding:12mm}.header{display:grid;grid-template-columns:62mm 1fr 58mm;gap:8mm;margin-bottom:8px}.logo img{width:58mm}.head-right{grid-column:3;padding-top:2mm}.doc-title{font-size:16pt;font-weight:700;margin:0 0 7mm 0}.topline{display:grid;grid-template-columns:1fr 1fr;gap:18mm;margin-top:6px}.line{display:flex;gap:8px;margin:4px 0}.label{font-weight:700;white-space:nowrap}.boxgrid,.sectiongrid{display:grid;grid-template-columns:1fr 1fr;gap:12mm;margin-top:10px}.box{border:1.5px solid #111;min-height:34mm;padding:7px}.boxtitle{font-weight:700;margin-bottom:6px;text-transform:uppercase}.addr{white-space:normal;word-break:break-word}.info,.schedule,.instructions{border:1px solid #111;padding:7px}.schedule{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:10px}.instructions{margin-top:10px;min-height:14mm}table{width:100%;border-collapse:collapse;margin-top:12px;table-layout:fixed}th,td{border:1px solid #111;padding:5px 4px;vertical-align:top;word-wrap:break-word;font-size:11pt}th{background:#f2f2f2;text-align:left}th:nth-child(4),th:nth-child(5),th:nth-child(6),td:nth-child(4),td:nth-child(5),td:nth-child(6){text-align:center}.total td{font-weight:700;background:#fafafa}.signatures{display:grid;grid-template-columns:1fr 1fr;gap:14mm;margin-top:18mm}.sigline{border-bottom:1px solid #111;height:15mm;font-size:18pt;padding-left:8px}.second{margin-top:8mm}.sublabel{text-align:center}.received-row{display:flex;align-items:flex-end;gap:8px;margin-bottom:13mm;font-weight:700}.received-line{border-bottom:1px solid #111;height:8mm;flex:1}[contenteditable='true']:focus{outline:1px dashed #777;background:#fffbe6}@media print{body{background:#fff}.page{margin:0;width:auto;padding:10mm}@page{size:A4 portrait;margin:8mm}}</style>"
EndFunc

Func _HtmlToPdf($sHtml, ByRef $aNums, $carrier, $delivery)
    If Not FileExists($sHtml) Then Return ""
    DirCreate($PDF_PATH)
    If $carrier = "" Then $carrier = "Autre"
    Local $rule = _SP_GetRule($carrier)
    Local $vars = _SP_BuildVariables($aNums, $delivery, $carrier)
    Local $pdfBase = _SP_ApplyTemplate($rule.Item("PDF"), $vars)
    Local $pdf = $PDF_PATH & _Safe($pdfBase) & ".pdf"
    If FileExists($pdf) Then FileDelete($pdf)
    Local $browser = _FindBrowser()
    If $browser = "" Then Return ""
    Local $url = "file:///" & StringReplace($sHtml, "\", "/")
    Local $cmd = '"' & $browser & '" --headless --disable-gpu --print-to-pdf="' & $pdf & '" "' & $url & '"'
    RunWait($cmd, @ScriptDir, @SW_HIDE)
    If FileExists($pdf) And FileGetSize($pdf) > 0 Then Return $pdf
    Return ""
EndFunc

Func _FindBrowser()
    If $g_sBrowserCache <> "" And FileExists($g_sBrowserCache) Then Return $g_sBrowserCache
    Local $a[6] = [@ProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe", @ProgramFilesDir & " (x86)\Microsoft\Edge\Application\msedge.exe", @LocalAppDataDir & "\Microsoft\Edge\Application\msedge.exe", @ProgramFilesDir & "\Google\Chrome\Application\chrome.exe", @ProgramFilesDir & " (x86)\Google\Chrome\Application\chrome.exe", @LocalAppDataDir & "\Google\Chrome\Application\chrome.exe"]
    For $i = 0 To UBound($a) - 1
        If FileExists($a[$i]) Then
            $g_sBrowserCache = $a[$i]
            Return $g_sBrowserCache
        EndIf
    Next
    Return ""
EndFunc
; ============================================================
; MODERN UX HELPERS - HOVER SOFT / NON BLOQUANT
; ============================================================
Func _CreateHoverButton($sText, $x, $y, $w, $h, $normalColor = 0xF8FAFC, $hoverColor = 0xE2E8F0, $textColor = 0x111827)
    Local $idButton = GUICtrlCreateButton($sText, $x, $y, $w, $h)
    GUICtrlSetFont($idButton, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($idButton, $normalColor)
    GUICtrlSetColor($idButton, $textColor)
    _RegisterHoverButton($idButton, $normalColor, $hoverColor)
    Return $idButton
EndFunc
Func _RegisterHoverButton($idButton, $normalColor, $hoverColor)
    Local $n = UBound($g_aHoverIds)
    ReDim $g_aHoverIds[$n + 1]
    ReDim $g_aHoverNormal[$n + 1]
    ReDim $g_aHoverHover[$n + 1]
    ReDim $g_aHoverIsHover[$n + 1]
    $g_aHoverIds[$n] = $idButton
    $g_aHoverNormal[$n] = $normalColor
    $g_aHoverHover[$n] = $hoverColor
    $g_aHoverIsHover[$n] = False
EndFunc
Func _UpdateHoverButtons()
    If TimerDiff($g_iLastHoverTick) < $HOVER_THROTTLE_MS Then Return
    $g_iLastHoverTick = TimerInit()
    If UBound($g_aHoverIds) = 0 Then Return
    Local $aCur = GUIGetCursorInfo()
    Local $hoverId = 0
    If IsArray($aCur) And UBound($aCur) >= 5 Then $hoverId = $aCur[4]
    For $i = 0 To UBound($g_aHoverIds) - 1
        If $hoverId = $g_aHoverIds[$i] Then
            If Not $g_aHoverIsHover[$i] Then
                GUICtrlSetBkColor($g_aHoverIds[$i], $g_aHoverHover[$i])
                $g_aHoverIsHover[$i] = True
            EndIf
        Else
            If $g_aHoverIsHover[$i] Then
                GUICtrlSetBkColor($g_aHoverIds[$i], $g_aHoverNormal[$i])
                $g_aHoverIsHover[$i] = False
            EndIf
        EndIf
    Next
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
    MsgBox(64, $sTitle, $sMessage)
EndFunc
Func _ModernErrorChoice($sTitle, $sMessage, $sBtn1, $sBtn2, $sBtn3)
    Local $r = MsgBox(51 + 262144, $sTitle, $sMessage & @CRLF & @CRLF & "Oui = " & $sBtn1 & @CRLF & "Non = " & $sBtn2 & @CRLF & "Annuler = " & $sBtn3)
    If $r = 6 Then Return 1
    If $r = 7 Then Return 2
    Return 3
EndFunc

; ============================================================
; CSV / SNAP / MAIL / EDOC / UTILS
; ============================================================
Func _SnapCsvGet($snap, $fileName, $rowIndex, $colIndex)
    Return _CsvGet($snap & $fileName, $rowIndex, $colIndex)
EndFunc

Func _CsvGet($file, $rowIndex, $colIndex)
    Local $lines = _CsvLoadCached($file)
    If Not IsArray($lines) Then Return ""
    If $rowIndex < 0 Or $rowIndex >= UBound($lines) Then Return ""
    Local $a = _ParseCsvLine($lines[$rowIndex])
    If $colIndex < 0 Or $colIndex >= UBound($a) Then Return ""
    Return _CleanJ($a[$colIndex])
EndFunc

Func _GetLine($file, $rowIndex)
    Local $lines = _CsvLoadCached($file)
    If Not IsArray($lines) Then Return ""
    If $rowIndex < 0 Or $rowIndex >= UBound($lines) Then Return ""
    Return $lines[$rowIndex]
EndFunc

Func _ParseCsvLine($line)
    Local $delim = ","
    If StringInStr($line, ";") And Not StringInStr($line, ",") Then $delim = ";"
    If StringInStr($line, @TAB) Then $delim = @TAB
    Local $a[0], $field = "", $inQ = False
    For $i = 1 To StringLen($line)
        Local $ch = StringMid($line, $i, 1)
        If $ch = '"' Then
            If $inQ And $i < StringLen($line) And StringMid($line, $i + 1, 1) = '"' Then
                $field &= '"'
                $i += 1
            Else
                $inQ = Not $inQ
            EndIf
        ElseIf $ch = $delim And Not $inQ Then
            ReDim $a[UBound($a) + 1]
            $a[UBound($a) - 1] = $field
            $field = ""
        Else
            $field &= $ch
        EndIf
    Next
    ReDim $a[UBound($a) + 1]
    $a[UBound($a) - 1] = $field
    Return $a
EndFunc

Func _JoinUniqueColsSnap($snap, $fileName, $row, $c1, $c2, $sep)
    Local $d = ObjCreate("Scripting.Dictionary")
    Local $res = ""
    For $c = $c1 To $c2
        Local $v = _SnapCsvGet($snap, $fileName, $row, $c)
        If $v = "" Then ContinueLoop
        If Not $d.Exists($v) Then
            $d.Add($v, 1)
            If $res <> "" Then $res &= $sep
            $res &= $v
        EndIf
    Next
    Return $res
EndFunc

Func _CreateMailSP($sPdf, ByRef $aNums, $carrierOverride = "", $deliveryOverride = "")
    Local $carrier = $carrierOverride
    If $carrier = "" Then $carrier = "Autre"
    Local $delivery = $deliveryOverride
    Local $rule = _SP_GetRule($carrier)
    Local $vars = _SP_BuildVariables($aNums, $delivery, $carrier)
    Local $subject = _SP_ApplyTemplate($rule.Item("SUBJECT"), $vars)
    Local $body = _SP_ApplyTemplate($rule.Item("BODY"), $vars)
    Local $signatureAdd = _SP_ApplyTemplate($rule.Item("SIGNATURE"), $vars)
    If $signatureAdd <> "" Then $body &= @CRLF & @CRLF & $signatureAdd
    Local $oOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($oOutlook) Then Return False
    Local $mail = $oOutlook.CreateItem(0)
    $mail.To = $rule.Item("TO")
    $mail.CC = $rule.Item("CC")
    $mail.BCC = $rule.Item("BCC")
    $mail.Subject = $subject
    If FileExists($sPdf) Then $mail.Attachments.Add($sPdf)
    $mail.Display
    Sleep(700)
    Local $sig = $mail.HTMLBody
    $mail.HTMLBody = _TextToHtml($body) & "<br><br>" & $sig
    If $MAIL_AUTOSEND Then $mail.Send
    Return True
EndFunc

Func _TextToHtml($s)
    $s = _H($s)
    $s = StringReplace($s, @CRLF, "<br>")
    $s = StringReplace($s, @CR, "<br>")
    $s = StringReplace($s, @LF, "<br>")
    Return "<div style='font-family:Calibri,Arial;font-size:11pt'>" & $s & "</div>"
EndFunc

Func _ExtractDeliveryDateTime($s)
    Local $x = StringStripWS($s, 3)
    $x = StringReplace($x, "Livraison HPE", "")
    $x = StringReplace($x, "LIVRAISON HPE", "")
    $x = StringReplace($x, "livraison HPE", "")
    $x = StringReplace($x, "Livraison", "")
    $x = StringReplace($x, "livraison", "")
    $x = StringReplace($x, "Delivery", "")
    $x = StringReplace($x, "delivery", "")
    $x = StringReplace($x, "HPE", "")
    $x = StringStripWS($x, 7)
    $x = StringRegExpReplace($x, "(?i)^\s*le\s*", "")
    While StringInStr($x, "  ")
        $x = StringReplace($x, "  ", " ")
    WEnd
    Return StringStripWS($x, 3)
EndFunc

Func _FindEdocPrinter()
    Local $wmi = ObjGet("winmgmts:\\.\root\cimv2")
    If Not IsObj($wmi) Then Return ""
    Local $printers = $wmi.ExecQuery("SELECT Name FROM Win32_Printer")
    For $p In $printers
        Local $name = String($p.Name)
        If StringLower($name) = "edoc upload" Then Return $name
    Next
    For $p In $printers
        Local $name = String($p.Name)
        If StringInStr(StringUpper($name), "EDOC") Then Return $name
    Next
    Return ""
EndFunc

Func _PrintPdfToEdoc($sPdf, ByRef $aNums)
    If Not FileExists($sPdf) Then
        _Status("EDOC KO : PDF introuvable" & @CRLF & $sPdf)
        Return False
    EndIf
    If Not _EdocOpenFirstDossier($aNums) Then
        _Status("EDOC KO : impossible d'ouvrir le dossier dans edoc Viewer CDG" & @CRLF & _JoinArray($aNums, " / "))
        Return False
    EndIf
    Local $printer = _FindEdocPrinter()
    If $printer = "" Then
        _Status("EDOC KO : imprimante edoc Upload introuvable.")
        Return False
    EndIf
    Local $net = ObjCreate("WScript.Network")
    If IsObj($net) Then
        $net.SetDefaultPrinter($printer)
        Sleep(700)
    EndIf
    ShellExecute($sPdf, "", "", "print", @SW_HIDE)
    Return _EdocValidateUploadWindow("Delivery Order", $aNums)
EndFunc


Func _EdocOpenFirstDossier(ByRef $aNums)
    If UBound($aNums) < 1 Then Return False
    If Not WinExists("edoc Viewer CDG") Then Return False
    Local $sFirst = _CleanJ($aNums[0])
    If $sFirst = "" Then Return False
    WinActivate("edoc Viewer CDG")
    If Not WinWaitActive("edoc Viewer CDG", "", 5) Then Return False
    ControlSetText("edoc Viewer CDG", "", "Edit1", "")
    Sleep(600)
    ControlSetText("edoc Viewer CDG", "", "Edit1", $sFirst)
    Sleep(600)
    ControlSend("edoc Viewer CDG", "", "Edit1", "{ENTER}")
    Sleep(3500)
    Return True
EndFunc
Func _EdocValidateUploadWindow($sDocType, ByRef $aNums)
    If Not WinWait("Upload Documents CDG", "", 35) Then
        _Status("EDOC KO : la fenêtre 'Upload Documents CDG' n'est pas apparue.")
        Return False
    EndIf
    WinActivate("Upload Documents CDG")
    If Not WinWaitActive("Upload Documents CDG", "", 8) Then Return False
    Sleep(3500)

    ; 1 seul dossier : TAB -> Delivery Order -> ENTER.
    ; Plusieurs dossiers : TAB -> Delivery Order -> TAB 1 -> TAB 2 -> collage numéros -> validation.
    Send("{TAB}")
    Sleep(900)
    Send($sDocType)
    Sleep(1200)

    If UBound($aNums) <= 1 Then
        Send("{ENTER}")
    Else
        Send("{TAB}")
        Sleep(900)
        Send("{TAB 2}")
        Sleep(1000)
        ClipPut(_ArrayNumsClipboard($aNums))
        Sleep(400)
        Send("^v")
        Sleep(1200)
        Local $hUpload = ControlGetHandle("Upload Documents CDG", "", "TButton2")
        If $hUpload <> "" Then
            ControlClick("Upload Documents CDG", "", "TButton2")
        Else
            Send("{ENTER}")
        EndIf
    EndIf

    If WinWaitClose("Upload Documents CDG", "", 45) Then
        Sleep(300)
        Return True
    EndIf
    _Status("EDOC attention : validation envoyée mais la fenêtre ne s'est pas fermée.")
    Return False
EndFunc
Func _ArrayNumsClipboard(ByRef $aNums)
    Local $s = ""
    For $i = 0 To UBound($aNums) - 1
        Local $j = _CleanJ($aNums[$i])
        If $j <> "" Then $s &= $j & @CRLF
    Next
    Return $s
EndFunc

Func _FillEdocWindow(ByRef $aNums)
    Local $title = $EDOC_UPLOAD_TITLE
    If Not WinExists($title) Then Return False
    WinActivate($title)
    WinWaitActive($title, "", 5)
    Sleep(500)
    Send("{TAB}")
    Sleep(220)
    Send("Delivery Order")
    Sleep(280)
    Send("{TAB}")
    Sleep(150)
    Send("{TAB}")
    Sleep(150)
    ClipPut(_JoinArray($aNums, @CRLF))
    Sleep(180)
    Send("^a")
    Sleep(100)
    Send("^v")
    Sleep(350)
    Local $hUpload = ControlGetHandle($title, "", $EDOC_UPLOAD_BTN)
    If $hUpload <> "" Then
        ControlClick($title, "", $EDOC_UPLOAD_BTN)
    Else
        Send("{TAB}")
        Sleep(180)
        Send("^u")
    EndIf
    ; Anti-conflit EDOC : on ne lance pas le PDF suivant tant que la fenêtre upload n’est pas fermée.
    If WinWaitClose($title, "", 35) Then
        Sleep(200)
        Return True
    EndIf
    Return False
EndFunc

Func _ScanWaitUploadReady($timeoutSec)
    Local $t = TimerInit()
    While TimerDiff($t) < ($timeoutSec * 1000)
        If WinExists($EDOC_UPLOAD_TITLE) Then
            Local $state = WinGetState($EDOC_UPLOAD_TITLE)
            If BitAND($state, 2) Then
                WinActivate($EDOC_UPLOAD_TITLE)
                WinWaitActive($EDOC_UPLOAD_TITLE, "", 5)
                Local $hUpload = ControlGetHandle($EDOC_UPLOAD_TITLE, "", $EDOC_UPLOAD_BTN)
                If $hUpload <> "" Then
                    Local $pos = WinGetPos($EDOC_UPLOAD_TITLE)
                    If IsArray($pos) Then
                        Local $chk1 = PixelChecksum($pos[0] + 10, $pos[1] + 10, $pos[0] + 180, $pos[1] + 70)
                        Sleep(180)
                        Local $chk2 = PixelChecksum($pos[0] + 10, $pos[1] + 10, $pos[0] + 180, $pos[1] + 70)
                        If $chk1 = $chk2 Then
                            Sleep(350)
                            Return True
                        EndIf
                    Else
                        Sleep(350)
                        Return True
                    EndIf
                EndIf
            EndIf
        EndIf
        Sleep(180)
    WEnd
    Return False
EndFunc

Func _CleanJ($s)
    $s = String($s)
    $s = StringReplace($s, Chr(239) & Chr(187) & Chr(191), "")
    $s = StringReplace($s, "ï»¿", "")
    $s = StringReplace($s, "Ï»¿", "")
    $s = StringReplace($s, ChrW(65279), "")
    Return StringStripWS($s, 3)
EndFunc

Func _CleanNum($s)
    Return StringReplace($s, ",", ".")
EndFunc

Func _TrimNum($n)
    Local $s = StringFormat("%.3f", $n)
    While StringRight($s, 1) = "0"
        $s = StringTrimRight($s, 1)
    WEnd
    If StringRight($s, 1) = "." Then $s = StringTrimRight($s, 1)
    Return $s
EndFunc

Func _JoinArray(ByRef $a, $sep)
    Local $s = ""
    For $i = 0 To UBound($a) - 1
        If $s <> "" Then $s &= $sep
        $s &= _CleanJ($a[$i])
    Next
    Return $s
EndFunc

Func _H($s)
    $s = String($s)
    $s = StringReplace($s, "&", "&amp;")
    $s = StringReplace($s, "<", "&lt;")
    $s = StringReplace($s, ">", "&gt;")
    $s = StringReplace($s, '"', "&quot;")
    $s = StringReplace($s, @CRLF, "<br>")
    $s = StringReplace($s, @CR, "<br>")
    $s = StringReplace($s, @LF, "<br>")
    Return $s
EndFunc

Func _WriteAnsi($file, $txt)
    Local $sDir = StringLeft($file, StringInStr($file, "\", 0, -1))
    If $sDir <> "" And Not FileExists($sDir) Then DirCreate($sDir)
    Local $h = FileOpen($file, 2)
    If $h = -1 Then Return False
    FileWrite($h, $txt)
    FileClose($h)
    Return FileExists($file)
EndFunc

Func _WriteUtf8($file, $txt)
    Local $sDir = StringLeft($file, StringInStr($file, "\", 0, -1))
    If $sDir <> "" And Not FileExists($sDir) Then DirCreate($sDir)
    Local $o = ObjCreate("ADODB.Stream")
    If Not IsObj($o) Then Return False
    $o.Type = 2
    $o.Charset = "utf-8"
    $o.Open
    $o.WriteText($txt)
    $o.SaveToFile($file, 2)
    $o.Close
    If Not FileExists($file) Then Return False
    If FileGetSize($file) <= 0 Then Return False
    Return True
EndFunc

Func _Safe($s)
    Local $a[12] = ["\", "/", ":", "*", "?", '"', "<", ">", "|", @CR, @LF, @TAB]
    For $i = 0 To UBound($a) - 1
        $s = StringReplace($s, $a[$i], "_")
    Next
    Return StringStripWS($s, 3)
EndFunc

Func _LogoDataUri()
    Local $s = ""
    $s &= "iVBORw0KGgoAAAANSUhEUgAAAQUAAABhCAYAAADfhLRpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAACbVSURBVHhe7Z0JnI3VG8epkH1nGDN2svwrZSdkCZU1a4k2lUhSISkkUvyzVkryV6mUiopKpZJISouKKMqaMGPfef6f7zOd671n7szcOzN3zOX8Pp/zmbnvPe/7nuV5fuc5z3nOuVnEwcHBwYMs9gUHB4dzG44UHBwc/OBIwcHBwQ+OFBwcHPzgSMHBwcEPjhQcHBz84EjBwcHBD44UHBwc/OBIwSHT4uTGv+TYh5/KsU8+z3zpo0/l+LJv7CKfFXCk4JApceyzpRIfe7HER1WR+BJVM1cqVlniil8kR1+daxf7rIAjBYdMh+NLlkv8RXUlLneMxOUvk7lSzpISF/sfOTxxmsiJE3bRzwo4UnDIVDg6+02JK1JR4vLGSlzBcpkr5YmR+Ao15diCRXaxzyo4UnDINDj89AyJK1lN4vKWTqyQZzrliZU9F9WVYx9/Zhf7rIMjBYczj+PH5ciU6ap4u7OXSJg22OlMWg55YtR6OfruB3bJz0o4UnA44zgybabsbdJG9rXuIvuu7po4te8hextdI3GFK0hcgbKJlTacKXeMxMdeIkf+95pd7LMWjhQczhhOHTosh4aPlX3NOogcOSqn9u5LnI4eleNffyt7G10rcfkyeFqRM1r2Nu8gx7/9wS76WQ1HCg5nBidPysG7h8juLPkTSCEJnFj9q+xt0DrBhLeVNpwpZ7RaLsRKnGtwpOCQ4Th15Ijs73KLjvy7c0XLvuYd7SyK48tXyp5azSQuV6nEShvOlLuU7GvTXU5u2W4X6ZyAIwWHjMXx"
    $s &= "43LwgeE6EsflL5skKRxbvET2VKsncbmiEyttuBL+ihwlZE/1BnJi/Qa7SOcMHCk4ZBhO7dsvBx96TOIKl08IBCpYLjEpHD8ux5Z+LfFlLpXdWQrJ7mzFT6ccJRIrcnql/GUlLl8ZnTKcWLvOW+xzDo4UHDIEJ9aul32tOieM/P8SAskmhVM7d8vh/06Vg/cPl0NDHzudHnlcDg4YqoqbSKHTmihP3li1YE7tP+BX7nMRjhQcMgQsO+7OUjCRQtqkkBzUykhvhyNThgJl5eDDY+TU4SP2K89JOFJwyBAcmfGKTgFspQyKFA4fkcNjJiSY+OkZp8ASZ65ScnDYaF0NcUiAIwWHDEFqSeHkjp1yoFdfiSdwyTPtSHMilLpoJTk84Vk3ZbDgSMEhQ5AaUji5ZZvsrdVMdmcrlqDEWArpkfLESnypanJk5mz7lQ6OFBxSg0ULF0mX1l3kxg43+qXubbrL4LsH29kVoZLCyc1bZW+t5hJXpIKO6OmWClfQlY2jc96xX+nwL8JGCpv+2iRzX50rc16eIwvmLbC/Dgnbt26X12a9Jm/OftOX/joHI80yA3b+s1NK5CghhbIUkqjsUVLywpK+VPS8otL40sb2LYpQSeFU/B45+uY8Ofreh+mb5i+UY59/Zb/OwYOwkMJniz6Thv9pKPmy5JOC5xWUMgXKyC1dbpGbO98ccrqt221yTaNrpMh5RaR49uKaCmctLL069pJ9e/fZr3YIM44fPy7vv/O+jBo6SmpWqCnlCpbzpZjcMXJV3avsWxShkoLDmUNYSOGPdX/I5Ccmy6C+g6RxjcaqxPaoYlLxbMWlQJYCAVOhrIUSyCBbcSmTv4yULVBWhS82b6zUqlhLflz1o/1qhwzEgQMHpN/N/aRkzpKOFM4i"
    $s &= "hIUUvFi/dr289dpbcsXFV0iZfGX8RhY+X3nZlfL+2+/LJx98Ih8v/NiXsDamT50uA+8YqILXoHoDKV+ovI8U6lap60ghE+Dp/z4txS4o5kjhLELYScFg9szZOuJ7SYHP7wTh8Dl16pScPHlSZr84Wy2O0nlLO1LIJEgXUmjSxs7ucAaRYaTw+qzXA5LC/Dfn21mTxIkTJ+SBvg+oEDpSyBxIEykQd3BhSTk04gk7u0MIYMA8duyYfTnViChSACu+WiGXlL5Ealas6UghEyDVpEA0YdGKcqDP/XrAikPoOHLkiCz5dImu8j391NPy0L0PyW9rfrOzhYyIIwXQs2NPubTMpfLjd44UzjRSRQoEIhUqL4cmPKMnKzmkDt8s+0bKFy4vF0VdJG2vbCsTx06Uzq06y58b/rSzhoSIJIW9e/dq7MPunbvtrxwyGCGTQr4yEl+ymhx+fpadxSFE3Hv7vfL4I49rm7es11Ju7Xqr3N/nfnngrgfsrCEhIknBIfMgJFLIWlh254mVwxOm2V8HjT1xe2TxosWy6c9N9lfnFDb/tVluaHeDrPl5jdzR4w6pXbm2dG7dWT56/yNpclkTdc6nFhFFCgcPHJQ98Xvsyw5nEEGTwrMzZXfWInJw8Ej7q5AwdfxUickbo1Gt5zqG3jNU3n3rXWnbpK1OH5rWbCrXNr5Wbup0k501JEQMKaxbs04ri3mUGSMZN23cJMMfGC6vzHjF/irsMEu23pRRCJYUTqxZJ0fnLRQ5kbhslDeYkY14lsvKXyal85VON1Kw2y2YciSFnTt2ysP3Pywvv/CynDju/5NyqXm+XTa7X4ntaVarmbbJgN4DpFX9VtKmcRuJj4v3yxcqIoYUHh3yqOTJkkeFDtMpEPA1bN+2Xf7e/neS"
    $s &= "ie937dzlu4dlzl27diV7H9/t3hXYf3Hq5CnZ8fcOnc8VzFJQTbeMAntCvv/2e/nv6P/qikztSrU19Lh5reay7Itl8sf6P0ISwtQgWFIIhMOHD2twW52L6sjzk5+3v/YDCle/an2JzhWt73rh6Re0T+y+Iv2z4x/7dj/s2L5DndRMQ+pXqy+1KtXStkO5+t3ST7764iv14ttKmBywYjHj82bJK1dfcbXKhF4/eFAWzl+oz768wuXSok4L3cdDfhv0FdHAy79crhvLLi9/uU4LSKy2Nbq0kcx8bqb89utvvn5FF54Y8YTK38A7B8qvq3+1HxsyIoIUaECsBEKeafAtm7bYWRRTxk3RsOpWDVrpfolACWbt1LKTNj745+9/pM+NfaRJjSbSumHrRPlbN2itSsYzibL0YtGCRcrQbZq00ShLlIJODzcYhYYPGq7OJdoQpURZCOoixeaJVYJCEJ956hn79pDw3YrvVECHDhgqQ+4eIls3b/X7PrWkgMLRX4Srs7lq/KjxdhYFwj9vzjw1kbEQjOw0rN5QTWW7v5AP0tuvvW0/SjHvjXnSsn5LqVikosoTIdoxeWK03fjL3hraDoV8+L6H1UINBihm0fOLSqncpaRD8w4StztO/vzjT+nRrofKBony85f2QuZs0FcQAOUyfUqijGwJ4N4SF5aQqiWqyqznwueoPbOkkL24vP164M7zYsE7C+Q/Mf/R+++88c4kAzWen/K8KmjVklWlVK5SKqTehLJccckVyqqGWA4fOqybezC9qkdX97uPTq5YtKI0r91crm97vaz6ZpVaBs9OeFauu+o6XQ4iwpK81IdOGzlkpI5EKNOw+4ap4jISMsp5U51KdeS6FtfJjGdmyPcrv5dDhw7Z1QkIdiminLQduxWx"
    $s &= "EPr07KPTFnamMkJt3rRZvdJ8F5UtSu6+5e6gHHNHDh+R1T+sllnTZ0n3a7vr6FmucDm10KqVqibReaOlQtEK0viSxvLgPQ/Kpx9+qu1g6p8SKWCVQSo4w65ueLXuZUGJKhSuoKHwXjBFfOOVN/R55DMWgkmqxFb/msS+mdt73u73vI1/bJS+N/WVsvnLah/TduzoxEJZsWyFxMXF6RLfyMEjtc/M3pxqJarJ6GGjA5rkWDorl6+UEYNHSOVilbWcyBiyhBxcVu4yicoRJeUKnC63kROsEwOsCVYMjCwh62wgnPX8LLUMIF76g3KbwZTP4cIZJQU6tl3TdqqUjz74aKLEdUZEYhJgWe5n3p4SEDoa3mygMh1R96K68kMSv/YDOdzW/TYVUjZfITRPjnxSln621M+HgWCjZBAGz/TWh/dhjja6pJFE54zWTuZ5JCNkMD0jkloWeWL0Gvf27t5bpwPJgSXYZjWbqaAhfAj1siXLVDgDARO5SlQVJRAEdc3qNXYWH3DgsscE4aadqR/k+uIzL8r/pv1Pfvr+J23Xmc/OlP639tfRnbJ7R++USGHyk5OVaBgFTf+g7B0D7H3AochGOr6nrt52Nm1NP9mJ57Ju/+XiL33PWvr5Ut21y4hr3tn16q5Jbr9f+8taNfXVgshXWmUBJd22dZtfvulTpqsC0w5eWeN/2pD3UCb63ysrlIOBwmD196uVOIzcYUHZIFBpypNTfLKUVmdicjijpMAGJ5QHc43RiMScjM8IHQnBMIIXLCmAhfMWaoeZe2nwatHV1FsbCCuWrlDBQRDIP/6x8XI0icCaQwcPqQK/OvNV3yhphIGRqEf7HjLh8QlqqmI1MG9+bvJzusHrrp53SdmCZbXe3vtQlG7XdJO/NgQWVEZx"
    $s &= "TGiEhnsQDkaSlMCIjEAiSFgPNjDP58+drxvWaF+IGtMaS2DD74F/+4Dt0y9MfSGhvh5lSIkUnpv0nE7fKhWrpP3BPUmRQvzueLXmtmzeotMX6mvair+Y2rSt5rGSl1yZehLsZqY4tAVTg6QIwYBpLe1gdudyP33qxYvPvig1ytRINDjwmWkk7frzjz/LsIHDfNMCEu//YP7pH6sdO2KsygP3MaCwzBgI1AsZpZ/OWlKgsTGx7rjhDnUcwbyTnpgkt99wu9zQ9gbp0aGHnuZjOiYUUsAku+e2e3wjsQptnhidXuzb4796geDA3JhnjEpMD4IBTh2vWcvz7+l9j53ND1gknEeAoJqRxNxfLFsxad+svezds9e+TRWK55uRDi/z/n377WyJAEEZq4SdprYjCoca+0h8o2jOaJk2Mbg4AgTTKKvWPwVSABAtIyEWCXVPihS8YNcsVpeRGe795adf7GwBwRQHS8nbR1g7wQD5MSY7gwtOXDta8N257/qRHIMC5WNK6MUvP/4iT49/WpMdnv/I/Y8kTH3yxEiL2i2SJGOA5czgyUlX4cIZJQU6+uUZL9tZ5eiRo3Jw/0EdkVFYRi0zfSC+O1js379fOjTr4EcMCFXX1l39FOqj9z5S5w0szvw8WK+zTQpKKBODIxSmIa/+71UVIi8x8Lw3X0m83Eb0mmk/FBFTPxAgHZxcnHXAPB1rTEfXQuW0ft7yUU+eq9bUv/Ne5tNYJcGAKYWZS6vCBUEKgDl8jbI19L2pIYVqJasF7YMZ3G+wz/rjb7druwX0/AcCBMC7TP0giGsaXqOWkhfImJED2gDCTmkq6AWKzn28hzbBN5YUGDBYVcJKChfOKCnwOZjVh61btmpDY8IxF0xqeTAQWF1A0I3ZTcMzKg7q"
    $s &= "N0g7l7k08eKU5cb2N8rPP/1sPyJJBCKFQPPBpIDpjtcaJTeCZ8oGaRiwBIXZaMxUFL1Diw7S67peMubhMer74Dl48HGmYQ3hAKweU12fjaCRIEfugTgAzkljfZCPNqatgwVth5IZayFYUsAPYvxE4SQF2teQD/fizxk3apydLUnQB/SFmepB3hcVv0hJzYtOV3XyyQFt0b5p+5CUllUOY5FQP8rJlCOYOoYDEUEKgIbDx8AcGe96KDiw/4D07NDTN2IYBbyvz33S9PKmUjJHSV1hCHYEMUgrKQCchESjMZc3ZcNqgcwMWAr1+iAYnXEC4kRk3utNCC3fkS6OvViVzySuIeRg4+8bddpgfCI4FhlVQwVKZvo1s5HC6IdG+/l8qOMbL79hZ0sWEKepH+9mpYRVES9sUsB5TrxEsMBX4T2AiPdAQDitP3z3Qzl2NPBqW7gQMaTAtIG5FKN+UnEKyWHjho2+EdM0PAKDsuHc+fKz097qYJEepAB++O4HnyWjz8kb62d+Erlm5sW8g7MrGa1Wfr1Svln+TdBp+dLlGgQEWOY1c2FS9VLV5YtPvvCUKjikJk4ho0jhrl53+e5D6aijd1UiGHA4UJHzi/j6hucRm+JFWkkBeb7y8it9x9p55ZP+ZuXD9gWFExFDCqzPYxKzbBhKg3vBkhpLYt4lLjp50thJdtagkF6k8Pe2v9UCMpYMSmpIAVMfa8aQGUKX1DHqoQBnIqcv80zey7F4qUFmJgWWTjVO4N86svJhgtaCBUuB17e53qewlIMjAr1IKymAzz/+XC0DdMJYsj49uaC4BtHhjMdfFG5EDCmkF+ylNASTaEavuR4s0osUmLZwarV37mpIYd++fRokZcqL4hHoktbwZVZ6fMqcyUkBc93rE0ot"
    $s &= "KUCu639bb2dLEXf2uNP3/nCRAmBVgjJTR2TBDAQqp3lLKzm0u7KdrAvzr2Kfk6RgBBjPP155nDx4kIMNaTVIL1IAEx+fqEuS+hzP9IHAKV2d+Hd1IL0shUghBcgPP4g3TiE1pECbsiRL3EAoIMYE095YceGYPnjBUjrLxPhDqpSoohGpRr5MXA/7III52zS1iFhSYLmShnlp+kv2V0li1cpVGvqLgOBYxBFnlgNZriOqMKkQ6kBIL1JA8AldZjRC6CmXmfvjAWf50TjMEE7K6V2dSA2Ie8DxxjOZb/PO5UuW29lSRLhJgWVT4gVM/UMhBVZizLSDRDkDLfcmBwLUTDuRKC9hzV6kJyl4wcao26+/XZ/ttRp4B/4gfE3hQMSSwrcrvtUOJzosGLAngBDhIlmLSMcWHXW9t3e33lIi+2mztEKRCrpmHKzCpRcp4C8haEkdn7mi1ZHojRVgTl3s/ATFg8QgtrSeK0Ho76VlExST5yL4bEsPFRlBCgNuH5AqS4FAN2+MCvJmrxykBE4Q964+sBfGHqXDRQoAWVz80WKNkDTWCokBhOjYpKJu04KIJQVGT+4PRpA57JWOwjzHoWcUihBadkAagUPhUOy5s+fajwiI9CIFdlviGUdJKINt/bC5ipBtY9XwHkarYMkrEDas36Cbw0zdeTdm6c8/hGZeZwgp9B6QKksB4Fz0OWlzlVLrIdglPt79YP8H/Xw9bAxb++tav3xpIQX6EJJJabv310u/1iVlIwO8h7galtvTGxFJCiMGjVCmhDmfGvOU/bUfMMPZysq7GDUohxfsCkRwvCMRChrMiBKIFIKNaDTAyYijk3spA/Pe33/73S8PAoOgGeEkb6i/kBVoWoQDzRv/gOLh5WbfQbDICFLgnABv/6Ac"
    $s &= "wSoD8S2+EbZAObUG1/wSeG+BDeJhIGPj5IVUWtZtaWdLEykQQEfsw7RJKYeWE2RnAth4D6HywbZDKIgoUmB5jn39dBKMScMsWbzEzubDti3bpOd1CUFLdCgmeqDDXtmsZCL7jHATQMR++ORgkwLPQIBDAbtBzW5DlN0mLQP2Y5DHJ6C5S6kwBjONQMlx1gU6FYr4D6+JDUnw3GCdrpw0ZDzzwZICo56ZuqRECgAT3rv6QJAWu1eDAcvQXrmj3ThzI25Xykt7BLeZ/qWsREcS9GXDJgWc1uaQlZSApQDx1KtWL8nDgwzY5m0IDnlmf1Aw+19CRYaQAhuQHnvoMT/hMxXjx10ICU2qcgg9HUhwkcb/eza3EA5qm3IG7NtnHwN5NG+OEjoyBgLzMmVhT/wCjc+uxeQ2p9ikoApboJyGHKcU+44FQz5GTAQJz3JyUw/OA0B4vPNKVeCWnTToCMW3lykZZZk6MdUiGnTs8LF+3wP2L2AZefuG/+tVrac7EVk/12efPP1s3sOqCNuq2V1pzPNgSQErjLm5Ttfyxkq9KvXk15+SDs754N0PdLcsK0W8h/fhF2IbPGWjLQl9Jxyc8ym8hIFcEWdgnIX0EeUcNzL5cGfe6d3oVPSCojLu0cT34JtiW7pvGvbvr5cFS6qQgp67kD1K7rvzvmSjarEOjaVA308dN9XOki4ICykQYIHCfLzgYw0rRcDs7aUmmYg61pBnPD1D8zNaktjQc1Wdq9Q7Tmd6lZZrNKb38BCUgD3vjCxGeMzIakakQNGQCBZLdHYZMYsRPsKMA+1ctEmBhNJCXIQis3Ua4WCDFzDHbWHd0CY8Hz8HdSFvSuDgDiMQ5n0II+2HAM+cNlPmvDRH244zDDGd+Q4lQOCSwntvv6eHfjKi"
    $s &= "ciiqUVYS7VepaCXdWciz8d5DXvgfKL+3zegfpkLUke2//DUOUw4poS3ML5J77yMwiKg9/BmByJRlQQJ7vBF/5n5Mb5PYkk572kfi8U4Ul4GBulEnnsUKTCAzn/waAv7vuRcoLPN370iO1crUk+3oXn+P9kmuUnowCwewQFS0BVu1A/mAuIYFwv0sP7K3x7b+kB9+sNls5aYeTDPZGBUOpDspEAGG8w9nUHTuaA1NpqNo2OQSTM5P13OegjfB0HZeGo8zF2x/AtMAIhaLZE/4pepE9+WI0g4YcNvpdWYEF+djwfMKJrqHlYnC5xeWqJxRevyX7YC0SYEOI+qSuR6jbVSuKKlTuY6O5ph6nJajBFkgVuvFKMuGJnOiU0rAouHgFIJnKKtu9f5XORB0ysx1RlUsA04g4pQnzgQM5FPwgqkWoyP+F/ZP8Bz12+SOUeLhF8DpD+3PC4upknEIDnX1bubhL3VmhYTYB0ZzhJyVH6wi9muw/Gu3M+1fOaqylpcR2S4vy3MckoOccI8pF4k+4BrE0KV1F1VGG1gSHGHGahV5uY/8DEZsKGMAQsE5YYs8RmbZXMbUzd7gNGbYGA0qQ2ZtuSEVyFpAv2fgMntPsNpsnDxxUo/NM4fv0Ib4dYixgIBxsjIwUVbyYFEyuK1asUrvDQfSnxQOH1GzE5MV4YUg0jsNvH2gjpr2QRmYpXiXGRXte0yiXExlDFBsphnJ3cN3rBdzAKcXNikwUnJiEMDyYGmRUYsTkthfAQlgXvOsUQ+OCnqLtg0UhqhGotuYH3uFkqg33sMSFlGSqTn5mpOHiA2459Z79Dns8ccfM6jvIOl/S389rRiwAkAdIR9DSFgKBPv06tjLdy4DS670CwLOFNBuX5P4rm+vvjp3trcnA6wI"
    $s &= "8iBXkBJtSfk4fIbTsFhmTQmM7pSD+1BYyK7weYW13fiLkkPuKB4ylhQ44Ia9FXYd7PrwLhJnMiZ1BsScl+doHSBR2pEyFcxaUC0x/ubPkl/LSmxNIL9QeiPdSeFcQiBS8PoFIEhMTnMiEP+zz+H4icQCHyqYinAqNY4vDhJhLZvE/1g/vI8Ar7QA09aUO6kwcAiEOArzbnwfW/7aEvSZDKkBdcdnZMoG6YRCsJAq93LwiWk3b/thlcTHx8spSdl6Sw9QHywZHNt2eUgEKVFWpjrHjidv8aUHHCmkASmRgoNDJMKRQhrgSMHhbIQjhTTAkYLD2QhHCmmAIwWHsxGOFNIARwoOZyMcKaQBjhQczkY4UkgDApFCqBuiHBwyGxwppAEEqHj3IvA/4brA3ofg4BApcKQQAji5if0E/KKP+Qk3wlLtRCg1m6/IN6T/kJDPKHBwOJNwpBAC2BVI/Dlx82y2Ig49UDIbirAc2G8RzG8+OjhkFjhSCAHsVmObLiG/Qaft/wR90o+DQ2aAIwUHBwc/OFJwcHDwQ1hIAc+7vWuNa+eKR37z5s0ycuRIWb489CPT0wOBth3boC+8+Uyfkbge6ECQSAY/qnP33XdL3759ZffuhCP5Vq1alWxbPffcc0HJLL9u3r9/f7nrrrtk165deu2FF14I+nDZzIawkML69etl1qxZUrlyZSlTpoyULl1aevfurZ0QTCN7gXBu27bNvpyp8Ntvv/mUCAFp166dZMmSRVq1amVnDTvi4uLkjjvukJiYGClbtqzs3Jnw+xE2fv/9d7ntttu0f2JjY6Vly5by1ltvyRtvvCEPPvigNGvWTIoUKSKDBiX8IG2kAjK4/vrrVf5GjBghw4cPl0mTJknnzp2lRo0aSSruJ598ou2H"
    $s &= "zCaFtWvXSocOHWTcuHE6CPDsgQMHSoMGDVTm+WGXSERYSAHQ2Llz51blyJkzp7z33nt2lqAwfvx46datm30502D+/PkSHR2tvx5twGhEve+9916/vBmF7du36/tJf/+d+LgxAwjswgsv1HyPPPKI33dYO1dccYX07NnT73qkgdGe+n399dd+1z/66COpUqWKHDgQ+DRkCIT77rsv6WPs6tSpI7Vr15Z//jl91gSDw2OPPaaE6kjBQlpJATP22WeflQsuuEBHscyIadOmSZ48eSRv3rx+x4chaFg3oVpF6YXUkAKjHMCc/uGHH/T/d955R3r16mXdFVkwBD12bOJDa7Eg9u5NfPYm1lWjRo30PuQPqyEQChcurMq/YMECv+tr1qyRSy+9VK2USESGkwLz7AceeEDNUsxVTrh5+OGHZfDgwfLoo4/6zN2XXnrJJ9jlypVTk3bRokW+53Pv888/L++//76yOiMbeO211/T55Of/efPm6bsQ+hUrVuhzeRcJE/rtt9+W+++/Xx566CE/HwDC8vnnn8vrr78uH3/8sZbxiSee8M1Hf/75ZylQoICWL0eOHNKvXz9ZuXKlCsSQIUM0P+/yAmV78sknZfHixToyz5w50zftmDNnjq9dyMc7GaX4TD05+9Lgu+++k6eeeko+++wzbTPyfPnl6Z9YTw0pYP4CLJ5XX31V/6cvEPhff/1V22fo0KFqudFulJ9rlAWQj7b+8MMPtY589/333/veg4lNu/Dd6tWrtS2p77Bhw+Svv/yP1WNeTv0+/fRTVWbeuWNHwhmJjL7UlWt8z3OYviUF/AjUj3qOGjVK+8iAvgrkO6G+KPV5552n9zLyB0Lx4sX1++zZs8uECRO0HMaXRn96rcdIQoaTAgLC6MN12Bjl"
    $s &= "Hj16tC8vcz8AWaCsXKtfv778+OOPKuwoZfv27aVYsWIyd+5c7cD8+fPrHBrhYoSuVKmS3leiRAkVuooVK+pnBBVFhf3NZ3wfV199tU9wvvjiC30/gpAtWzYVYt7ZqVMnzVOtWjUVJoTz5ptv1mtYC0uWLFEl4nDVAQMG6PWrrjp93PnEiRO1HczoxLtNGRCkrVu3qjnLNRxWlJO24DPp5ZcTzkWcPXu2nH/++VKzZk2t68KFC/V7RqylSxPOKEwNKfAuyAqypK5e4Kdg7kw+FIDR97LLLtPPKC1TKP6HNHgmz+BzVFSUj8ghSa7RV/QJpHfxxRfrtWuuucb3LpQWsxySZaTFPCfPDTfcoMTQsWNHlRWesWfPHu1rFPiPPwKfbAxxmrYgQeS0HQNGoJ9cY3BBvrgPXwv38L6ffvrJzqr+A/NcZAqZ5N5QreLMhgwnBYDwcL1Hjx6+a7A416677jrfNUYXrqG0BozemOtYDxs2JPwmA55e8sHW4PLLL9fPCDJAYbEmACNarly59HsEEyC4PJNrppyMYnw2fgGj6HQ+1gMwwo+ge6cPEADXDSlgoRjhYcoBIAE+Z82a1TcnRRm4hgVh0LhxY732zDPP6OfJkyfrZ5Thzz//9D2HxHcgNaQAQeM/wDncvHlzO6t618nH6AggY/oHYpoyZYp+d9NNN+l39INdJsrPZ4ia9gDUk2vUGyAzpr5YAgCfAASENUQ/8R2kiEUCxowZo9ewUgIBxceqgaBMmUwyUyYvsMogPWCeTcJ5aAOSgqyMPJkEOUA6kYqIIwWUimt0BKMIwswIi3AjQMBLCvZoQDkoD98bUsDpxGjPNVNOOhzv8i+//KKrKcYqwHrAQw+MJZMSKTCaGoExyu1V5i5duui1YEgBa4Qy"
    $s &= "MaJCikx5zHOmTk04STo1pIBlhA+EZ2LN2GDkJp8hBS+wpGgriIL7UTbz/lBIgbphiXANCw1QJiwVpk9McfgOIsXKwBGKFchnppGBgAnPvbQD8kFfmbLR52ZgMaAvaFPKwjTH5C1VqpRvOgTobxLyxVQS+cDHYPLTTmZ5MtIQcaSA4HONDkDgUGyUGgvCLF2mBymgWCxdYZHceuutWi6+Tw0peJUkECnUqlVLrwVDCoBpEyYw0zCmFeY5aSEFM2qiRExRbCRHCrQxxEc98OUw/THvD4UUsJhQcK4ZS8ELnm2eiwVH/5HwLQRSQMgAP4YX69at0yVb85xrr73W9x3tyjIl1iRWCc/FerLbFzC9YNrkBc/GoiAv9XjxxRf9vo8URBwpmCWmokWL+jnX8JozxwRpIYUPPvhA57ItWrTQz8QcADOdSIoUeL/xNtukgDPTCFYgUli2bJleS44UWIkBmLSYz4UKFZKvvvpKld48Jy2kQPsD/BvGmepFcqRg/AWM3MDUnxQKKTA6m/Lga/EC0sX3ZJ7LlMUA/04gpx6kUK9ePdm06fSviAGsQOMT8fpPIH/bOYzD0LyT2A18XQDiJL8N+pK8OClfeSX8v9EQDkQEKSBseMSZyyOwpkMJDsKMxnFHIA7KDQwp4PQJlRS49s033/gEgZEDBybzbD57ScE4mvLly6fXjDJjYXDdkALK17RpU71mSIH38BnBMh7r5EjBRNeZ56A8jGjGx0BKCyl42z0QzNJeIFIwoyPTOEZs438ghUIKwDhg8d1gvVBnVjuwEjD1mzRpot8THIR5j+LfeeedieIQAO2KA7Z169Z+QVwMHjyHqYdZPWFK1r1794DBTMbJSqIcgP5GvimjV8aYSpCPZyXX9pkZYSEFPMHMUYkI"
    $s &= "Y1WA+RjebVYQWPPF9EU4qlevrmvheM3xDyBwCAidDRhh8QATdceIzZwV0IEoB1MIlJboNBgeAcF6qFu3rjp7sDBY3TACwWiDw6pkyZLqeGK5imt4ss01BIxRByXg+cxxMdNxVKL8PPfGG29UUxdnEvdQP96F+ci8k9US6gJ5ffvtt/pulqt4N0TFfTjlsEawGACWQPny5fU+nFeMfjNmzFDB5RpzXYQPwsFKYCRCkbFgEFqIrm3btkpolJ36kMgfCCaikf4xCbLGJLZDf5lj8376rGrVqkrQ3mVElg9pKywmRl76DbMc5ejatavWkQA02or34IikjakndaPexhKCqLDACAgjf5s2bZR8kRNAGzds2FDbAFmh75MKR2a5EYcszlPkj4GHJVzeT/lMPAZ/aUNWMaZPn+57Fn3Ae+lb08/II32MXCG/yCb1R955dp8+fbSPA01nIgVhIQU7rh6YuHo6yqwNe+PsvddC2TdhnmHAc0xe/tqfvc8239nXDAhC8i5F4UwznW2Xx3zmflOeQHVhFDfk5gX3mLymXIGuAYQVLzkw1xBUrpPf69+w+8GA+7z5zOdA6/Y801tfb1kMUPyNGzf6rqPcKBgw/W5gnmfLgQ3ae8uWLYneZcD9tiXoBfUxZIN1gK8Df4KtsKbNvH1nEKiugPBnM31l2ggR82xvdGOkIiyk4ODgELlwpODg4OAHRwoODg5+cKTg4ODgB0cKDg4OfnCk4ODg4AdHCg4ODn5wpODg4OAHRwoODg5++D8bEYSTuWkTBwAAAABJRU5ErkJggg=="
    Return "data:image/png;base64," & $s
EndFunc

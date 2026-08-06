; ============================================================================
; StateService.au3
; Service etat dispatch/data/status.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _Spinner($sTxt)
    ToolTip($sTxt, 0, 0, "Robot E.TMS", 1)
EndFunc

Func _WinWaitSpinner($sClass, $sTxt)
    _Spinner($sTxt)
    Return WinWait($sClass, "", 10)
EndFunc

; ==============================================================================
; BACKUP ROTATION — garder les N dernières sauvegardes
; ==============================================================================

Func _Batch_FC($sData)
    If $sData = "" Then Return
    $bFC_Stop = False
    $bFC_Pause = False
    $bFC_Skip = False
    HotKeySet("{F9}", "_HK_FC_PauseToggle")
    HotKeySet("{ESCAPE}", "_HK_FC_Stop")

    Local $aJobs = StringSplit($sData, "|")

    ; Pré-calculer la liste complète des numéros individuels pour le tracker
    Local $aAllNums[100]
    Local $iTotal = 0
    For $i = 1 To $aJobs[0]
        Local $aDetails = StringSplit($aJobs[$i], ";")
        If $aDetails[0] >= 1 Then
            ; Séparer les groupes "J1A001 + J1A002 + J1A003"
            Local $aSubs = StringSplit(StringRegExpReplace($aDetails[1], "\s*\+\s*", "|"), "|")
            For $s = 1 To $aSubs[0]
                Local $sN = StringStripWS($aSubs[$s], 3)
                If $sN <> "" Then
                    If $iTotal >= UBound($aAllNums) Then ReDim $aAllNums[$iTotal + 20]
                    $aAllNums[$iTotal] = $sN
                    $iTotal += 1
                EndIf
            Next
        EndIf
    Next
    ReDim $aAllNums[$iTotal]
    _Tracker_Start("File Closing — Kanban col 5", $aAllNums)

    Local $iTrackIdx = 0
    Local $iDoneFC = 0, $iStoppedFC = 0
    Local $sRemainingFC = ""
    For $i = 1 To $aJobs[0]
        Local $aDetails = StringSplit($aJobs[$i], ";")
        If $aDetails[0] >= 1 Then
            Local $sFileField = $aDetails[1]
            Local $sTransp    = ""
            Local $sContact   = ""
            Local $sDateG     = ""
            Local $sHoraire   = "09h et 12h"
            Local $sDLY       = ""
            Local $sDLYNotes  = ""
            If $aDetails[0] >= 4 Then $sTransp   = $aDetails[4]
            If $aDetails[0] >= 5 Then $sContact  = $aDetails[5]
            If $aDetails[0] >= 6 Then $sDateG    = $aDetails[6]
            If $aDetails[0] >= 7 And $aDetails[7] <> "" Then $sHoraire = $aDetails[7]
            If $aDetails[0] >= 8 Then $sDLY      = $aDetails[8]
            If $aDetails[0] >= 9 Then $sDLYNotes = $aDetails[9]
            Local $sCarrier = _FC_ResolveCarrier($sTransp)

            ; Éclater le groupe en dossiers individuels
            Local $aSubs = StringSplit(StringRegExpReplace($sFileField, "\s*\+\s*", "|"), "|")
            For $s = 1 To $aSubs[0]
                Local $sNumJ = StringStripWS($aSubs[$s], 3)
                If $sNumJ = "" Then ContinueLoop
                $bFC_Skip = False
                _Tracker_Update($iTrackIdx, 1)
                _FC_WaitIfPaused()
                If $bFC_Stop Then
                    _Tracker_Update($iTrackIdx, 3)
                    $iStoppedFC = 1
                    ; Collecter les dossiers restants
                    For $rr = $iTrackIdx To $iTotal - 1
                        $sRemainingFC &= $aAllNums[$rr] & @CRLF
                    Next
                    ExitLoop 2
                EndIf
                If $bFC_Skip Then
                    _Tracker_Update($iTrackIdx, 4)
                    $bFC_Skip = False
                    $iTrackIdx += 1
                    ContinueLoop
                EndIf
                If $sCarrier = "UPS" Then
                    _Run_FileClosing_UPS($sNumJ)
                Else
                    _Run_FileClosing_Single($sNumJ, $sCarrier, $sDateG, $sHoraire, $sContact, $sDLY, $sDLYNotes)
                EndIf
                If $bFC_Stop Then
                    _Tracker_Update($iTrackIdx, 3)
                    $iStoppedFC = 1
                    For $rr = $iTrackIdx + 1 To $iTotal - 1
                        $sRemainingFC &= $aAllNums[$rr] & @CRLF
                    Next
                    ExitLoop 2
                EndIf
                If $bFC_Skip Then
                    _Tracker_Update($iTrackIdx, 4)
                    $bFC_Skip = False
                Else
                    _Tracker_Update($iTrackIdx, 2)
                    $iDoneFC += 1
                EndIf
                $iTrackIdx += 1
                _Tracker_PollButtons()
                _FC_SmartSleep(300)
            Next
        EndIf
    Next
    HotKeySet("{F9}")
    HotKeySet("{ESCAPE}")
    _Tracker_End()
    ; Bilan final FC
    If $iStoppedFC And $sRemainingFC <> "" Then
        ClipPut(StringStripWS($sRemainingFC, 2))
        MsgBox(48+262144, "FC — Arrêté", _
            $iDoneFC & " dossier(s) traité(s) sur " & $iTotal & "." & @CRLF & @CRLF & _
            "Dossiers restants (copiés dans le presse-papier) :" & @CRLF & $sRemainingFC)
    ElseIf $iStoppedFC Then
        MsgBox(48+262144, "FC — Arrêté", $iDoneFC & " dossier(s) traité(s) sur " & $iTotal & ".")
    EndIf
    $bFC_Stop = False
    $bFC_Pause = False
    $bFC_Skip = False
EndFunc

; ==============================================================================
; FILE CLOSING UPS — séquence courte (pas de Carrier ID, juste DEF)
; ==============================================================================

Func _URIDecode($sStr)
    $sStr = StringReplace($sStr, "+", " ")
    Local $aMatch
    While 1
        $aMatch = StringRegExp($sStr, "%([0-9A-Fa-f]{2})", 3)
        If Not IsArray($aMatch) Then ExitLoop
        $sStr = StringReplace($sStr, "%" & $aMatch[0], Chr(Dec($aMatch[0])), 1)
    WEnd
    Return $sStr
EndFunc

; ==============================================================================
; RÉSEAU PARTAGÉ
; ==============================================================================

Func _Net_SaveState($sPath, $sJSON)
    Local $hFile = FileOpen($sPath, 2)
    If $hFile = -1 Then
        _NotifyError("Réseau", "Impossible d'écrire : " & $sPath)
        Return False
    EndIf
    FileWrite($hFile, $sJSON)
    FileClose($hFile)
    Return True
EndFunc

Func _Net_LoadState($sPath)
    If Not FileExists($sPath) Then
        _NotifyError("Réseau", "Fichier introuvable : " & $sPath)
        Return "{}"
    EndIf
    Local $hFile = FileOpen($sPath, 0)
    If $hFile = -1 Then
        _NotifyError("Réseau", "Impossible de lire : " & $sPath)
        Return "{}"
    EndIf
    Local $sContent = FileRead($hFile)
    FileClose($hFile)
    Return $sContent
EndFunc

; ==============================================================================
; CONFIG PJ
; ==============================================================================

Func _AttachPJIfExists($sFile, $sTransp)
    Local $aCfg  = _GetPJConfig()
    Local $sBase = $aCfg[1]
    If $sBase = "" Then Return ""
    Local $sSubDir = ""
    If StringInStr($sTransp, "UPS") Then $sSubDir = $aCfg[3] & "\"
    Local $sDir  = $sBase & $sSubDir
    Local $aExts = StringSplit($aCfg[2], ",", 1)
    For $i = 1 To $aExts[0]
        Local $sFilePath = $sDir & $sFile & "." & StringStripWS($aExts[$i], 8)
        If FileExists($sFilePath) Then Return $sFilePath
    Next
    Return ""
EndFunc

; ==============================================================================
; DIAGNOSTIC COMPLET — Benchmark système + E.TMS + fichiers + contrôles
; ==============================================================================

Func _NotifyError($sSource, $sMsg)
    TrayTip("Dispatch — Erreur " & $sSource, $sMsg, 10, 3)
    _AuditLog("ERREUR", $sSource & " : " & $sMsg)
EndFunc

; ==============================================================================
; AUDIT LOG — Journal d'erreurs en arrière-plan
; ==============================================================================

Func _SanitizeString($s)
    Local $sResult = StringRegExpReplace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    Return $sResult
EndFunc

; ── /api/save-patch — Sauvegarde différentielle ──

Func _API_SavePatch($sBody)
    Local $sId = _GetJsonValue($sBody, "id")
    Local $sVersion = _GetJsonValue($sBody, "v")
    Local $sUpdatedAt = _GetJsonValue($sBody, "updatedAt")
    Local $sUpdatedBy = _GetJsonValue($sBody, "updatedBy")
    Local $sChanges = _GetJsonValue($sBody, "changes")

    If Not _ValidateId($sId) Then
        Return '{"success":false,"reason":"invalid_id"}'
    EndIf

    $sId = _SanitizeString($sId)
    $sUpdatedBy = _SanitizeString($sUpdatedBy)

    If Not FileExists($g_sDataFile) Then
        Return '{"success":false,"reason":"no_data_file"}'
    EndIf

    Local $hRead = FileOpen($g_sDataFile, 256)
    If $hRead = -1 Then
        Return '{"success":false,"reason":"cannot_read_file"}'
    EndIf
    Local $sData = FileRead($hRead)
    FileClose($hRead)

    Local $sSearchPattern = '"file"\s*:\s*"' & StringRegExpReplace($sId, '([\.\+\*\?\[\]\(\)\{\}\^\$\\])', '\\\1') & '"'
    Local $iPos = StringRegExp($sData, $sSearchPattern)

    If Not $iPos Then
        $sSearchPattern = '"id"\s*:\s*"' & StringRegExpReplace($sId, '([\.\+\*\?\[\]\(\)\{\}\^\$\\])', '\\\1') & '"'
        $iPos = StringRegExp($sData, $sSearchPattern)
    EndIf

    If Not $iPos Then
        Return '{"success":false,"reason":"dossier_not_found","id":"' & $sId & '"}'
    EndIf

    Local $iNewV = Number($sVersion)
    If $iNewV <= 0 Then $iNewV = 1

    If $sChanges <> "" And $sChanges <> "{}" Then
        Local $aKeys = StringRegExp($sChanges, '"([^"]+)"\s*:', 3)
        If IsArray($aKeys) Then
            For $k = 0 To UBound($aKeys) - 1
                Local $sKey = $aKeys[$k]
                Local $sVal = _GetJsonValue($sChanges, $sKey)
                $sVal = _SanitizeString($sVal)
                _AuditLog("PATCH", "Dossier=" & $sId & " Clé=" & $sKey & " Val=" & $sVal)
            Next
        EndIf
    EndIf

    _AuditLog("PATCH", "Patch reçu pour " & $sId & " v" & $iNewV & " par " & $sUpdatedBy)
    Return '{"success":true,"v":' & $iNewV & ',"id":"' & $sId & '"}'
EndFunc

; ── /api/net-check — Vérification légère réseau ──

Func _API_NetCheck($sPath)
    If Not _ValidateNetPath($sPath) Then
        Return '{"error":"invalid_path","changed":false}'
    EndIf

    Local $sDir = StringRegExpReplace($sPath, "\\[^\\]*$", "")
    Local $sGlob = StringRegExpReplace($sPath, "^.*\\", "")
    $sGlob = StringReplace($sGlob, ".json", "_*.json")

    Local $hSearch = FileFindFirstFile($sDir & "\" & $sGlob)
    If $hSearch = -1 Then
        Return '{"lastModified":"","checksum":"none","changed":false,"files":0}'
    EndIf

    Local $sCheckData = ""
    Local $iFiles = 0
    Local $sLatest = ""

    While True
        Local $sFound = FileFindNextFile($hSearch)
        If @error Then ExitLoop
        Local $sFullPath = $sDir & "\" & $sFound
        Local $sSize = FileGetSize($sFullPath)
        Local $sTime = FileGetTime($sFullPath, 0, 1)
        $sCheckData &= $sFound & ":" & $sSize & ":" & $sTime & "|"
        If $sTime > $sLatest Then $sLatest = $sTime
        $iFiles += 1
    WEnd
    FileClose($hSearch)

    Local $iSum = 0
    For $c = 1 To StringLen($sCheckData)
        $iSum += Asc(StringMid($sCheckData, $c, 1))
    Next
    Local $sChecksum = Hex(Mod($iSum, 16777216), 6)

    Local $bChanged = ($sChecksum <> $g_sLastNetChecksum)
    $g_sLastNetChecksum = $sChecksum

    Local $sISO = ""
    If StringLen($sLatest) >= 14 Then
        $sISO = StringLeft($sLatest, 4) & "-" & StringMid($sLatest, 5, 2) & "-" & StringMid($sLatest, 7, 2) & "T" & _
                StringMid($sLatest, 9, 2) & ":" & StringMid($sLatest, 11, 2) & ":" & StringMid($sLatest, 13, 2) & "Z"
    EndIf

    Return '{"lastModified":"' & $sISO & '","checksum":"' & $sChecksum & '","changed":' & _
           StringLower(String($bChanged)) & ',"files":' & $iFiles & '}'
EndFunc

; ── /api/job-status — Statut d'un job E.TMS en cours ──

Func _API_JobStatus()
    If Not IsDeclared("g_bJobRunning") Or Not $g_bJobRunning Then
        Return '{"jobId":"","status":"idle","progress":0,"total":0,"current":"","done":true}'
    EndIf

    Local $sStatus = "running"
    If IsDeclared("g_bJobPaused") And $g_bJobPaused Then $sStatus = "paused"

    Local $sJobId = ""
    If IsDeclared("g_sCurrentJobId") Then $sJobId = $g_sCurrentJobId

    Local $iProgress = 0
    If IsDeclared("g_iJobProgress") Then $iProgress = $g_iJobProgress

    Local $iTotal = 0
    If IsDeclared("g_iJobTotal") Then $iTotal = $g_iJobTotal

    Local $sCurrent = ""
    If IsDeclared("g_sJobCurrent") Then $sCurrent = _SanitizeString($g_sJobCurrent)

    Local $sType = ""
    If IsDeclared("g_sJobType") Then $sType = $g_sJobType

    Local $bDone = ($iProgress >= $iTotal And $iTotal > 0)

    Return '{"jobId":"' & $sJobId & '","status":"' & $sStatus & '","progress":' & $iProgress & _
           ',"total":' & $iTotal & ',"current":"' & $sCurrent & '","type":"' & $sType & _
           '","done":' & StringLower(String($bDone)) & '}'
EndFunc

; ── Sécurité réseau — Remplacements sécurisés ──

Func _Net_SaveState_Secure($sPath, $sJSON)
    If Not _ValidateNetPath($sPath) Then
        _AuditLog("SECURITY", "Chemin réseau rejeté (save) : " & $sPath)
        Return False
    EndIf
    If StringLen($sJSON) > $MAX_BODY_SIZE Then
        _AuditLog("SECURITY", "Body trop large pour save : " & StringLen($sJSON) & " bytes")
        Return False
    EndIf
    Local $hFile = FileOpen($sPath, 2 + 256)
    If $hFile = -1 Then
        _NotifyError("Réseau", "Impossible d'écrire : " & $sPath)
        Return False
    EndIf
    FileWrite($hFile, $sJSON)
    FileClose($hFile)
    _AuditLog("NET", "Sauvegardé : " & $sPath & " (" & StringLen($sJSON) & " bytes)")
    Return True
EndFunc

Func _Net_LoadState_Secure($sPath)
    If Not _ValidateNetPath($sPath) Then
        _AuditLog("SECURITY", "Chemin réseau rejeté (load) : " & $sPath)
        Return "{}"
    EndIf
    If Not FileExists($sPath) Then Return "{}"
    Local $iSize = FileGetSize($sPath)
    If $iSize > $MAX_BODY_SIZE Then
        _AuditLog("SECURITY", "Fichier trop gros : " & $sPath & " (" & $iSize & " bytes)")
        Return "{}"
    EndIf
    Local $hFile = FileOpen($sPath, 256)
    If $hFile = -1 Then
        _NotifyError("Réseau", "Impossible de lire : " & $sPath)
        Return "{}"
    EndIf
    Local $sContent = FileRead($hFile)
    FileClose($hFile)
    _AuditLog("NET", "Chargé : " & $sPath & " (" & StringLen($sContent) & " bytes)")
    Return $sContent
EndFunc

; CMR_FULL_FUSION_MODULE_APPENDED


Opt("GUIOnEventMode", 0)

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


; (Le module CMR complet a ete deplace dans src/features/cmr/ :
; CMR_Engine.au3, CMR_Results.au3, CMR_PDF.au3, CMR_SP.au3, CMR_UI.au3.
; Les constantes/variables globales ci-dessus restent ici -- partagees
; par plusieurs domaines (ETMS/EDOC en utilisent aussi certaines).)

Func _MainProcess($sCompte, $sProf, $sChoixTemps, $sDateStartManual, $sDateEndManual)
    If $sProf = "" Then Return MsgBox(16, "Erreur", "Veuillez sélectionner un profil.")
    Local $bAny = False
    For $i = 0 To $g_iDynCount - 1
        If GUICtrlRead($g_aDynCtrls[$i][0]) = $GUI_CHECKED Then $bAny = True
    Next
    If Not $bAny Then Return MsgBox(48, "Sélection", "Aucune action cochée.")
    Local $oStoreTarget = Null
    For $oStore In $g_oNamespace.Stores
        If $oStore.DisplayName = $sCompte Then $oStoreTarget = $oStore
    Next
    If $oStoreTarget = Null Then Return MsgBox(16, "Outlook", "Boîte mail introuvable.")
    $g_iGroupedCount = 0
    For $actIdx = 0 To $g_iDynCount - 1
        If GUICtrlRead($g_aDynCtrls[$actIdx][0]) <> $GUI_CHECKED Then ContinueLoop
        Local $sAct = $g_aDynCtrls[$actIdx][3]
        Local $bDoMail = (GUICtrlRead($g_aDynCtrls[$actIdx][1]) = $GUI_CHECKED)
        Local $bDoPJ = (GUICtrlRead($g_aDynCtrls[$actIdx][2]) = $GUI_CHECKED)
        If Not $bDoMail And Not $bDoPJ Then ContinueLoop
        Local $sSec = $sProf & "_" & $sAct
        Local $sStart, $sEnd
        If $sChoixTemps = "Dernier upload (Automatique)" Then
            $sStart = _FmtDate(IniRead($CFG_FILE, $sSec, "LastUpload_" & $sCompte, _DateAdd('h', -1, _NowCalc())))
            $sEnd = _FmtDate(_NowCalc())
        ElseIf $sChoixTemps = "Personnalisé" Then
            $sStart = _FmtDate($sDateStartManual)
            $sEnd = _FmtDate($sDateEndManual)
        Else
            $sStart = _FmtDate($sDateStartManual)
            $sEnd = _FmtDate(_NowCalc())
        EndIf
        Local $folderID = Int(IniRead($CFG_FILE, $sSec, "Folder", "5"))
        Local $keys = IniRead($CFG_FILE, $sSec, "Keywords", "")
        Local $prefix = IniRead($CFG_FILE, $sSec, "Prefix", "J")
        Local $len = IniRead($CFG_FILE, $sSec, "Length", "9")
        Local $senderFilter = IniRead($CFG_FILE, $sSec, "Sender", "")
        Local $docType = IniRead($CFG_FILE, $sSec, "DocType", "Document")
        Local $regex = "(?i)(" & $prefix & "[A-Za-z0-9]{" & $len & "})"
        If $len = "" Or $len = "0" Then $regex = "(?i)(" & $prefix & "[A-Za-z0-9]+)"
        Local $aKeys = StringSplit($keys, ",")
        Local $oFolder = $oStoreTarget.GetDefaultFolder($folderID)
        Local $oItems = $oFolder.Items
        If $folderID = 5 Then
            $oItems.Sort("[SentOn]", True)
        Else
            $oItems.Sort("[ReceivedTime]", True)
        EndIf
        Local $seen = "|"
        For $oItem In $oItems
            If $g_iGroupedCount >= 799 Then ExitLoop 2
            If $oItem.Class <> 43 Then ContinueLoop
            Local $itemDate
            If $folderID = 5 Then
                $itemDate = _FmtDate($oItem.SentOn)
            Else
                $itemDate = _FmtDate($oItem.ReceivedTime)
            EndIf
            If $itemDate > $sEnd Then ContinueLoop
            If $itemDate < $sStart Then ExitLoop
            Local $subj = $oItem.Subject
            If StringRegExp($subj, "(?i)^(RE|TR|FW|FWD)\s*:") Then ContinueLoop
            If $senderFilter <> "" Then
                If Not StringInStr($oItem.SenderEmailAddress, $senderFilter) And Not StringInStr($oItem.SenderName, $senderFilter) Then ContinueLoop
            EndIf
            Local $ok = True
            For $k = 1 To $aKeys[0]
                Local $key = StringStripWS($aKeys[$k], 3)
                If $key <> "" And Not StringRegExp($subj, "(?i)" & $key) Then
                    $ok = False
                    ExitLoop
                EndIf
            Next
            If Not $ok Then ContinueLoop
            Local $aJ = StringRegExp($subj, $regex, 3)
            If @error Then ContinueLoop
            Local $nums = "", $mem = ""
            For $j = 0 To UBound($aJ) - 1
                Local $num = StringUpper($aJ[$j])
                Local $code = $num & "-" & $sSec
                If StringInStr($g_sHistorique, "|" & $code & "|") > 0 Then ContinueLoop
                If StringInStr($seen, "|" & $num & "|") > 0 Then ContinueLoop
                If StringInStr("|" & $nums, "|" & $num & "|") > 0 Then ContinueLoop
                $nums &= $num & "|"
                $seen &= $num & "|"
                $mem &= $code & "|"
            Next
            If $nums = "" Then ContinueLoop
            Local $attSel = _DefaultAttSelection($oItem)
            If (Not $bDoMail) And $bDoPJ And $attSel = "" Then ContinueLoop
            Local $storeID = ""
            If IsObj($oItem.Parent) Then $storeID = $oItem.Parent.StoreID
            $g_aGrouped[$g_iGroupedCount][0] = $oItem
            $g_aGrouped[$g_iGroupedCount][1] = $nums
            $g_aGrouped[$g_iGroupedCount][2] = $docType
            $g_aGrouped[$g_iGroupedCount][3] = $bDoMail
            $g_aGrouped[$g_iGroupedCount][4] = $bDoPJ
            $g_aGrouped[$g_iGroupedCount][5] = $sAct
            $g_aGrouped[$g_iGroupedCount][6] = $sSec
            $g_aGrouped[$g_iGroupedCount][7] = $subj
            $g_aGrouped[$g_iGroupedCount][8] = $itemDate
            $g_aGrouped[$g_iGroupedCount][9] = $oItem.SenderName
            $g_aGrouped[$g_iGroupedCount][10] = $attSel
            $g_aGrouped[$g_iGroupedCount][11] = False
            $g_aGrouped[$g_iGroupedCount][12] = $mem
            $g_aGrouped[$g_iGroupedCount][13] = $oItem.EntryID
            $g_aGrouped[$g_iGroupedCount][14] = $storeID
            $g_iGroupedCount += 1
        Next
    Next
    If $g_iGroupedCount = 0 Then Return MsgBox(64, "Scan", "Rien à traiter pour les actions sélectionnées.")
    If Not _SelectionGUI() Then Return
    If _CountSelected() = 0 Then Return MsgBox(64, "Sélection", "Aucune ligne cochée.")
    If Not WinExists("edoc Viewer CDG") Then Return MsgBox(16, "EDOC", "La fenêtre 'edoc Viewer CDG' est introuvable.")
    _RunUpload($sCompte)
EndFunc

Func _SelectionGUI()
    Local $h = GUICreate("Sélection uploads EDOC", 1180, 700, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX))
    GUISetBkColor($C_BG)
    GUISetFont(9, 400, 0, "Segoe UI")
    GUICtrlCreateLabel("Sélection des mails / pièces jointes à uploader", 24, 18, 760, 30)
    GUICtrlSetFont(-1, 17, 800)
    GUICtrlSetColor(-1, $C_TEXT)
    GUICtrlSetBkColor(-1, $C_BG)
    GUICtrlCreateLabel("1 ligne = 1 mail source. Mail + PJ restent sur la même ligne. Double-clic = ouvrir mail Outlook.", 26, 52, 1080, 22)
    GUICtrlSetColor(-1, $C_MUTED)
    GUICtrlSetBkColor(-1, $C_BG)
    $g_idSelList = GUICtrlCreateListView("|Action|Date|Dossiers|Mail|PJ|Fichiers PJ|Objet", 24, 92, 1130, 490, BitOR($LVS_SHOWSELALWAYS, $LVS_SINGLESEL))
    $g_hSelList = GUICtrlGetHandle($g_idSelList)
    _GUICtrlListView_SetExtendedListViewStyle($g_hSelList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES, $LVS_EX_DOUBLEBUFFER))
    _GUICtrlListView_SetColumnWidth($g_hSelList, 0, 38)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 1, 120)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 2, 135)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 3, 190)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 4, 70)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 5, 70)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 6, 290)
    _GUICtrlListView_SetColumnWidth($g_hSelList, 7, 520)
    For $i = 0 To $g_iGroupedCount - 1
        GUICtrlCreateListViewItem(_BuildRow($i), $g_idSelList)
        _GUICtrlListView_SetItemChecked($g_hSelList, $i, False)
    Next
    Local $idOpen = GUICtrlCreateButton("Ouvrir mail", 24, 606, 115, 34)
    Local $idToggleMail = GUICtrlCreateButton("Mail oui/non", 150, 606, 110, 34)
    Local $idChoosePJ = GUICtrlCreateButton("Choisir PJ", 270, 606, 110, 34)
    Local $idAll = GUICtrlCreateButton("Tout cocher", 395, 606, 105, 34)
    Local $idNone = GUICtrlCreateButton("Tout décocher", 510, 606, 115, 34)
    Local $idMailOnly = GUICtrlCreateButton("Cocher mails", 640, 606, 110, 34)
    Local $idPJOnly = GUICtrlCreateButton("Cocher PJ", 760, 606, 95, 34)
    Local $idCancel = GUICtrlCreateButton("Annuler", 900, 600, 120, 44)
    Local $idRun = GUICtrlCreateButton("Lancer upload", 1035, 600, 120, 44)
    GUICtrlSetBkColor($idRun, $C_ACCENT)
    GUICtrlSetFont($idRun, 10, 800)
    Local $idInfo = GUICtrlCreateLabel("", 24, 660, 1110, 22)
    GUICtrlSetColor($idInfo, $C_MUTED)
    GUICtrlSetBkColor($idInfo, $C_BG)
    GUIRegisterMsg($EDOC_WM_NOTIFY, "_Notify")
    GUISetState(@SW_SHOW, $h)
    _UpdateSelInfo($idInfo)
    While 1
        Local $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE, $idCancel
                GUIRegisterMsg($EDOC_WM_NOTIFY, "")
                GUIDelete($h)
                Return False
            Case $idRun
                _ApplySelection()
                GUIRegisterMsg($EDOC_WM_NOTIFY, "")
                GUIDelete($h)
                Return True
            Case $idOpen
                _OpenSelectedMail()
            Case $idToggleMail
                Local $idx = _SelIndex()
                If $idx >= 0 Then
                    $g_aGrouped[$idx][3] = Not $g_aGrouped[$idx][3]
                    _RefreshRow($idx)
                EndIf
            Case $idChoosePJ
                Local $idx2 = _SelIndex()
                If $idx2 >= 0 Then
                    _PJPicker($idx2)
                    _RefreshRow($idx2)
                EndIf
            Case $idAll
                _SetAllChecks(True)
            Case $idNone
                _SetAllChecks(False)
            Case $idMailOnly
                _CheckRowsByMode("MAIL")
            Case $idPJOnly
                _CheckRowsByMode("PJ")
        EndSwitch
        _UpdateSelInfo($idInfo)
    WEnd
EndFunc

Func _BuildRow($i)
    Local $mail = "Non"
    If $g_aGrouped[$i][3] Then $mail = "Oui"
    Local $pj = "Non"
    If $g_aGrouped[$i][4] Then $pj = _AttSummary($g_aGrouped[$i][10])
    Return "|" & $g_aGrouped[$i][5] & "|" & $g_aGrouped[$i][8] & "|" & _PipeToComma(StringTrimRight($g_aGrouped[$i][1], 1)) & "|" & $mail & "|" & $pj & "|" & _AttNames($i) & "|" & StringReplace($g_aGrouped[$i][7], "|", "/")
EndFunc

Func _RefreshRow($idx)
    Local $a = StringSplit(_BuildRow($idx), "|")
    For $c = 1 To $a[0]
        _GUICtrlListView_SetItemText($g_hSelList, $idx, $a[$c], $c - 1)
    Next
EndFunc

Func _SelIndex()
    Return _GUICtrlListView_GetNextItem($g_hSelList, -1, $LVNI_SELECTED)
EndFunc

Func _SetAllChecks($checked)
    For $i = 0 To $g_iGroupedCount - 1
        _GUICtrlListView_SetItemChecked($g_hSelList, $i, $checked)
    Next
EndFunc

Func _CheckRowsByMode($mode)
    For $i = 0 To $g_iGroupedCount - 1
        If $mode = "MAIL" Then _GUICtrlListView_SetItemChecked($g_hSelList, $i, $g_aGrouped[$i][3])
        If $mode = "PJ" Then _GUICtrlListView_SetItemChecked($g_hSelList, $i, ($g_aGrouped[$i][4] And $g_aGrouped[$i][10] <> ""))
    Next
EndFunc

Func _ApplySelection()
    For $i = 0 To $g_iGroupedCount - 1
        If _GUICtrlListView_GetItemChecked($g_hSelList, $i) And ($g_aGrouped[$i][3] Or ($g_aGrouped[$i][4] And $g_aGrouped[$i][10] <> "")) Then
            $g_aGrouped[$i][11] = True
        Else
            $g_aGrouped[$i][11] = False
        EndIf
    Next
EndFunc

Func _Notify($hWnd, $iMsg, $wParam, $lParam)
    If $g_hSelList = 0 Then Return $GUI_RUNDEFMSG
    Local $t = DllStructCreate("hwnd hWndFrom;uint_ptr IDFrom;int Code;int Item;int SubItem", $lParam)
    If @error Then Return $GUI_RUNDEFMSG
    If DllStructGetData($t, "hWndFrom") = $g_hSelList And DllStructGetData($t, "Code") = $EDOC_NM_DBLCLK Then
        Local $idx = DllStructGetData($t, "Item")
        If $idx < 0 Then $idx = _SelIndex()
        If $idx >= 0 Then _OpenMail($idx)
        Return 0
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func _OpenFirstDossier($nums)
    Local $a = StringSplit(StringTrimRight($nums, 1), "|")
    If $a[0] < 1 Then Return False
    If Not WinExists("edoc Viewer CDG") Then Return False
    WinActivate("edoc Viewer CDG")
    WinWaitActive("edoc Viewer CDG", "", 5)
    ControlSetText("edoc Viewer CDG", "", "Edit1", "")
    Sleep(600)
    ControlSetText("edoc Viewer CDG", "", "Edit1", $a[1])
    Sleep(600)
    ControlSend("edoc Viewer CDG", "", "Edit1", "{ENTER}")
    Sleep(3500)
    Return True
EndFunc

Func _CountSelected()
    Local $n = 0
    For $i = 0 To $g_iGroupedCount - 1
        If $g_aGrouped[$i][11] Then $n += 1
    Next
    Return $n
EndFunc

Func _CountNums($nums)
    Local $clean = StringTrimRight($nums, 1)
    If $clean = "" Then Return 0
    Local $a = StringSplit($clean, "|")
    Return $a[0]
EndFunc

Func _DefaultAttSelection($oMail)
    Local $s = ""
    For $i = 1 To $oMail.Attachments.Count
        If _IsPrintable($oMail.Attachments.Item($i).FileName) Then $s &= $i & "|"
    Next
    Return $s
EndFunc

Func _AttIndexSelected($sel, $idx)
    Return StringInStr("|" & $sel, "|" & $idx & "|") > 0
EndFunc

Func _AttSummary($sel)
    If $sel = "" Then Return "Non"
    Local $a = StringSplit(StringTrimRight($sel, 1), "|")
    Return $a[0] & " PJ"
EndFunc

Func _AttNames($idx)
    If $g_aGrouped[$idx][10] = "" Then Return ""
    Local $oMail = $g_aGrouped[$idx][0]
    Local $a = StringSplit(StringTrimRight($g_aGrouped[$idx][10], 1), "|")
    Local $s = ""
    For $i = 1 To $a[0]
        Local $n = Int($a[$i])
        If $n >= 1 And $n <= $oMail.Attachments.Count Then
            If $s <> "" Then $s &= "; "
            $s &= $oMail.Attachments.Item($n).FileName
        EndIf
    Next
    If StringLen($s) > 90 Then $s = StringLeft($s, 87) & "..."
    Return StringReplace($s, "|", "/")
EndFunc

Func _IsPrintable($file)
    Local $e = StringLower(_Ext($file))
    Switch $e
        Case "pdf", "doc", "docx", "xls", "xlsx"
            Return True
    EndSwitch
    Return False
EndFunc

Func _NumbersClipboard($nums)
    Local $a = StringSplit(StringTrimRight($nums, 1), "|")
    Local $s = ""
    For $i = 1 To $a[0]
        If StringStripWS($a[$i], 3) <> "" Then $s &= $a[$i] & @CRLF
    Next
    Return $s
EndFunc

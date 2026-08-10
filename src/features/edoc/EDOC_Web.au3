; ============================================================================
; EDOC_Web.au3
; Web EDOC.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

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

Func _EDOC_WebInit()
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","reason":"outlook_not_open","stores":[],"profiles":[]}'
    Local $stores = _GetStoresList($g_oNamespace)
    Return '{"status":"ok","stores":' & _EDOC_JsonArrayFromPipe($stores) & ',"profiles":' & _EDOC_ProfilesJSON() & '}'
EndFunc

; Variante sans dependance a Outlook, pour l'onglet "Clients EDOC" (edition
; des profils/regles) qui n'a besoin ni des boites mail ni d'une session
; Outlook active -- juste de l'ini de config.
Func _EDOC_ProfilesOnlyJSON()
    Return '{"status":"ok","profiles":' & _EDOC_ProfilesJSON() & '}'
EndFunc

Func _EDOC_WebScan($sBody)
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","reason":"outlook_not_open","items":[]}'
    Local $sCompte = _GetJsonValue($sBody, "mailbox")
    Local $sProf = _GetJsonValue($sBody, "profile")
    Local $sChoixTemps = _GetJsonValue($sBody, "period")
    Local $sDateStartManual = _GetJsonValue($sBody, "start")
    Local $sDateEndManual = _GetJsonValue($sBody, "end")
    Local $sSpec = _GetJsonValue($sBody, "actions")
    If $sProf = "" Then Return '{"status":"error","reason":"missing_profile","items":[]}'
    Local $oStoreTarget = Null
    For $oStore In $g_oNamespace.Stores
        If $oStore.DisplayName = $sCompte Then $oStoreTarget = $oStore
    Next
    If $oStoreTarget = Null Then Return '{"status":"error","reason":"mailbox_not_found","items":[]}'
    $g_iGroupedCount = 0
    Local $sActs = IniRead($CFG_FILE, $sProf & "_Actions", "List", "")
    Local $aActs = StringSplit($sActs, "|")
    For $actIdx = 1 To $aActs[0]
        Local $sAct = $aActs[$actIdx]
        If $sAct = "" Then ContinueLoop
        Local $bDoMail = _EDOC_ActionEnabled($sSpec, $sAct, "mail")
        Local $bDoPJ = _EDOC_ActionEnabled($sSpec, $sAct, "pj")
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
    Return '{"status":"ok","count":' & $g_iGroupedCount & ',"items":' & _EDOC_GroupedJSON() & '}'
EndFunc

Func _EDOC_WebUpload($sBody)
    Local $sCompte = _GetJsonValue($sBody, "mailbox")
    Local $sIdx = _GetJsonValue($sBody, "indexes") ; pipe separated indexes
    If $g_iGroupedCount = 0 Then Return '{"status":"error","reason":"nothing_scanned"}'
    For $i = 0 To $g_iGroupedCount - 1
        $g_aGrouped[$i][11] = False
    Next
    Local $a = StringSplit($sIdx, "|")
    For $k = 1 To $a[0]
        Local $idx = Int($a[$k])
        If $idx >= 0 And $idx < $g_iGroupedCount Then $g_aGrouped[$idx][11] = True
    Next
    Local $n = _CountSelected()
    If $n = 0 Then Return '{"status":"error","reason":"nothing_selected"}'
    If Not WinExists("edoc Viewer CDG") Then Return '{"status":"error","reason":"edoc_viewer_not_found"}'
    _RunUpload($sCompte)
    Return '{"status":"ok","uploaded":' & $n & '}'
EndFunc

Func _EDOC_WebRuleSave($sBody)
    Local $prof = _GetJsonValue($sBody, "profile")
    Local $act = _GetJsonValue($sBody, "actionName")
    If $prof = "" Or $act = "" Then Return '{"status":"error","reason":"missing_profile_or_action"}'
    Local $sec = $prof & "_" & $act
    IniWrite($CFG_FILE, $sec, "Folder", _GetJsonValue($sBody, "folder"))
    IniWrite($CFG_FILE, $sec, "Keywords", _GetJsonValue($sBody, "keywords"))
    IniWrite($CFG_FILE, $sec, "Prefix", _GetJsonValue($sBody, "prefix"))
    IniWrite($CFG_FILE, $sec, "Length", _GetJsonValue($sBody, "length"))
    IniWrite($CFG_FILE, $sec, "Sender", _GetJsonValue($sBody, "sender"))
    IniWrite($CFG_FILE, $sec, "DocType", _GetJsonValue($sBody, "docType"))
    Return '{"status":"ok"}'
EndFunc

; ============================================================================
; PROFILS CLIENTS + ACTIONS -- creation/suppression, memes regles que les
; boutons "Nouveau"/"Supprimer" de l'ecran natif _RulesGUI() (Config.au3),
; ici exposees en JSON pour l'onglet web "EDOC Clients". Modifier les champs
; d'une action existante reste _EDOC_WebRuleSave() ci-dessus.
; ============================================================================

Func _EDOC_ProfileAddJSON($sBody)
    Local $sProf = StringStripWS(_GetJsonValue($sBody, "profile"), 3)
    If $sProf = "" Then Return '{"status":"error","message":"nom_vide"}'
    Local $sCurr = IniRead($CFG_FILE, "System", "ProfilesList", "")
    If _PipeContains($sCurr, $sProf) Then Return '{"status":"error","message":"deja_existant"}'
    If $sCurr = "" Then
        $sCurr = $sProf
    Else
        $sCurr &= "|" & $sProf
    EndIf
    IniWrite($CFG_FILE, "System", "ProfilesList", $sCurr)
    IniWrite($CFG_FILE, $sProf & "_Actions", "List", "")
    Return '{"status":"ok"}'
EndFunc

Func _EDOC_ProfileDeleteJSON($sBody)
    Local $sProf = StringStripWS(_GetJsonValue($sBody, "profile"), 3)
    If $sProf = "" Then Return '{"status":"error","message":"nom_vide"}'
    Local $sActs = IniRead($CFG_FILE, $sProf & "_Actions", "List", "")
    Local $aActs = StringSplit($sActs, "|")
    For $i = 1 To $aActs[0]
        If $aActs[$i] <> "" Then IniDelete($CFG_FILE, $sProf & "_" & $aActs[$i])
    Next
    IniDelete($CFG_FILE, $sProf & "_Actions")
    Local $sNewList = _RemoveFromPipeList(IniRead($CFG_FILE, "System", "ProfilesList", ""), $sProf)
    IniWrite($CFG_FILE, "System", "ProfilesList", $sNewList)
    Return '{"status":"ok"}'
EndFunc

Func _EDOC_ActionAddJSON($sBody)
    Local $sProf = StringStripWS(_GetJsonValue($sBody, "profile"), 3)
    Local $sAct = StringStripWS(_GetJsonValue($sBody, "actionName"), 3)
    If $sProf = "" Or $sAct = "" Then Return '{"status":"error","message":"champs_vides"}'
    Local $sCurr = IniRead($CFG_FILE, $sProf & "_Actions", "List", "")
    If _PipeContains($sCurr, $sAct) Then Return '{"status":"error","message":"deja_existant"}'
    If $sCurr = "" Then
        $sCurr = $sAct
    Else
        $sCurr &= "|" & $sAct
    EndIf
    IniWrite($CFG_FILE, $sProf & "_Actions", "List", $sCurr)
    Return '{"status":"ok"}'
EndFunc

Func _EDOC_ActionDeleteJSON($sBody)
    Local $sProf = StringStripWS(_GetJsonValue($sBody, "profile"), 3)
    Local $sAct = StringStripWS(_GetJsonValue($sBody, "actionName"), 3)
    If $sProf = "" Or $sAct = "" Then Return '{"status":"error","message":"champs_vides"}'
    Local $sNewList = _RemoveFromPipeList(IniRead($CFG_FILE, $sProf & "_Actions", "List", ""), $sAct)
    IniWrite($CFG_FILE, $sProf & "_Actions", "List", $sNewList)
    IniDelete($CFG_FILE, $sProf & "_" & $sAct)
    Return '{"status":"ok"}'
EndFunc

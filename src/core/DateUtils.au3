; ============================================================================
; DateUtils.au3
; Fonctions dates/jours ouvrables.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _AddWorkingDays($iDaysToAdd)
    Local $sDate = _NowCalcDate()
    While $iDaysToAdd > 0
        $sDate = _DateAdd('D', 1, $sDate)
        Local $iWDay = _DateToDayOfWeek(StringLeft($sDate, 4), StringMid($sDate, 6, 2), StringRight($sDate, 2))
        If $iWDay <> 1 And $iWDay <> 7 Then $iDaysToAdd -= 1
    WEnd
    Return StringRight($sDate, 2) & "/" & StringMid($sDate, 6, 2) & "/" & StringLeft($sDate, 4)
EndFunc

Func _FC_WorkDay($iDay, $iMon, $iYear, $iJours)
    Local $d = $iDay
    Local $m = $iMon
    Local $y = $iYear
    Local $iCount = 0
    While $iCount < $iJours
        $d += 1
        If $d > _FC_DaysInMonth($m, $y) Then
            $d = 1
            $m += 1
            If $m > 12 Then
                $m = 1
                $y += 1
            EndIf
        EndIf
        Local $iWDay = _FC_DayOfWeek($d, $m, $y)
        If $iWDay <> 1 And $iWDay <> 7 Then $iCount += 1
    WEnd
    Local $aResult[3]
    $aResult[0] = $d
    $aResult[1] = $m
    $aResult[2] = $y
    Return $aResult
EndFunc

Func _FC_DaysInMonth($m, $y)
    Local $aDays[13]
    $aDays[0]=0
    $aDays[1]=31
    $aDays[2]=28
    $aDays[3]=31
    $aDays[4]=30
    $aDays[5]=31
    $aDays[6]=30
    $aDays[7]=31
    $aDays[8]=31
    $aDays[9]=30
    $aDays[10]=31
    $aDays[11]=30
    $aDays[12]=31
    If $m = 2 And (Mod($y,4)=0 And (Mod($y,100)<>0 Or Mod($y,400)=0)) Then Return 29
    Return $aDays[$m]
EndFunc

Func _FC_DayOfWeek($d, $m, $y)
    If $m < 3 Then
        $m += 12
        $y -= 1
    EndIf
    Local $k = Mod($y, 100)
    Local $j = Int($y / 100)
    Local $h = Mod($d + Int(13*($m+1)/5) + $k + Int($k/4) + Int($j/4) - 2*$j, 7)
    Return Mod($h + 6, 7) + 1
EndFunc

; ==============================================================================
; MAILS RDV (Colonne 2)
; ==============================================================================

Func _Tracker_Update($iIndex, $iStatus)
    If $g_hTracker = 0 Then Return

    ; Vérifier les boutons à chaque update (polling GUI)
    _Tracker_PollButtons()

    Local $sText = "Attente"
    Local $iColor = 0xAAAAAA
    Switch $iStatus
        Case 1
            $sText = "En cours"
            $iColor = 0xFFCC00
            GUICtrlSetData($g_idTrackLbl, "Traitement : " & ($iIndex+1) & " / " & $g_iTrackCount)
            GUICtrlSetData($g_idTrackProg, ($iIndex / $g_iTrackCount) * 100)
            ; Afficher le numéro de dossier en cours
            Local $sJ = GUICtrlRead($g_aTrackIDs[$iIndex])
            $sJ = StringMid($sJ, StringInStr($sJ, "|") + 1)
            GUICtrlSetData($g_idBatchInfo, "→ " & $sJ)
        Case 2
            $sText = "OK"
            $iColor = 0x00CC55
            GUICtrlSetData($g_idTrackProg, (($iIndex+1) / $g_iTrackCount) * 100)
        Case 3
            $sText = "Stop/Err"
            $iColor = 0xFF4444
        Case 4
            $sText = "Passé"
            $iColor = 0x3399FF
    EndSwitch
    Local $sJ2 = GUICtrlRead($g_aTrackIDs[$iIndex])
    $sJ2 = StringMid($sJ2, StringInStr($sJ2, "|") + 1)
    GUICtrlSetData($g_aTrackIDs[$iIndex], $sText & "|" & $sJ2)
    GUICtrlSetColor($g_aTrackIDs[$iIndex], $iColor)
EndFunc

; Polling des boutons GUI — appelé dans les boucles d'attente et entre chaque dossier

Func _ValidateNetPath($sPath)
    If StringLen($sPath) > $MAX_PATH_LENGTH Or StringLen($sPath) = 0 Then Return False
    If StringInStr($sPath, "..") Then Return False
    If StringInStr($sPath, "%2e%2e") Then Return False
    If StringInStr($sPath, "%2E%2E") Then Return False
    If StringLeft($sPath, StringLen($ALLOWED_NET_PREFIX)) <> $ALLOWED_NET_PREFIX Then Return False
    If StringRegExp($sPath, '[<>"|*?]') Then Return False
    If StringRight(StringLower($sPath), 5) <> ".json" Then Return False
    Return True
EndFunc

; ── Validation d'un ID de dossier ──

Func _ValidateId($sId)
    If StringLen($sId) = 0 Or StringLen($sId) > 200 Then Return False
    Return StringRegExp($sId, $ID_PATTERN)
EndFunc

; ── Sanitisation d'une chaîne ──

Func _SP_UpdateKnownLabel()
    GUICtrlSetData($idSPKnown, "SP connus : " & StringReplace(_SP_ComboData(), "|", " / "))
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

Func _EdocValidateUploadWindow($sDocType, ByRef $aNums)
    If Not WinWait("Upload Documents CDG", "", 35) Then
        _Status("EDOC KO : la fenêtre 'Upload Documents CDG' n'est pas apparue.")
        Return False
    EndIf
    WinActivate("Upload Documents CDG")
    If Not WinWaitActive("Upload Documents CDG", "", 8) Then Return False
    Sleep(3500)

    ; 1 seul dossier : TAB -> type de document -> ENTER (le numero est deja
    ; pre-rempli par edoc au debut de l'upload, pas besoin d'aller le
    ; toucher). Plusieurs dossiers : TAB -> type de document -> TAB x2 ->
    ; collage numeros -> validation -- exactement 2 tabs, PAS 3 : l'ancien
    ; code faisait Send("{TAB}") PUIS Send("{TAB 2}"), et "{TAB 2}" en
    ; AutoIt envoie le Tab 2 FOIS (et non "le 2e Tab"), donc 1+2 = 3 tabs au
    ; lieu des 2 documentes ci-dessus -- d'ou le decalage de champ au collage.
    Send("{TAB}")
    Sleep(900)
    Send($sDocType)
    Sleep(1200)

    If UBound($aNums) <= 1 Then
        Send("{ENTER}")
    Else
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

Func _UpdateDates($idT, $idStart, $idLblEnd, $idEnd)
    Local $s = GUICtrlRead($idT)
    If $s = "Personnalisé" Then
        GUICtrlSetState($idStart, $GUI_ENABLE)
        GUICtrlSetState($idLblEnd, $GUI_SHOW)
        GUICtrlSetState($idEnd, $GUI_SHOW)
        If GUICtrlRead($idStart) = "" Or StringInStr(GUICtrlRead($idStart), "Calculé") Then GUICtrlSetData($idStart, @YEAR & "/" & @MON & "/" & @MDAY & " 00:00:00")
    Else
        GUICtrlSetState($idLblEnd, $GUI_HIDE)
        GUICtrlSetState($idEnd, $GUI_HIDE)
        If $s = "Dernier upload (Automatique)" Then
            GUICtrlSetData($idStart, "Calculé individuellement par action...")
            GUICtrlSetState($idStart, $GUI_DISABLE)
        Else
            GUICtrlSetState($idStart, $GUI_ENABLE)
            If $s = "Aujourd'hui" Then GUICtrlSetData($idStart, @YEAR & "/" & @MON & "/" & @MDAY & " 00:00:00")
            If $s = "Il y a 1h" Then GUICtrlSetData($idStart, _DateAdd('h', -1, _NowCalc()))
        EndIf
    EndIf
EndFunc

; ==================================================================================================
; GESTION DES RÈGLES
; ==================================================================================================

Func _UpdateSelInfo($idInfo)
    Local $rows = 0, $mails = 0, $pjs = 0
    For $i = 0 To $g_iGroupedCount - 1
        If _GUICtrlListView_GetItemChecked($g_hSelList, $i) Then
            $rows += 1
            If $g_aGrouped[$i][3] Then $mails += 1
            If $g_aGrouped[$i][4] And $g_aGrouped[$i][10] <> "" Then $pjs += 1
        EndIf
    Next
    GUICtrlSetData($idInfo, "Coché : " & $rows & " ligne(s) | uploads Mail : " & $mails & " | lignes avec PJ : " & $pjs & " | Total affiché : " & $g_iGroupedCount)
EndFunc

Func _FmtDate($s)
    $s = StringStripWS($s, 3)
    If StringRegExp($s, "^\d{14}$") Then Return StringMid($s,1,4) & "/" & StringMid($s,5,2) & "/" & StringMid($s,7,2) & " " & StringMid($s,9,2) & ":" & StringMid($s,11,2) & ":" & StringMid($s,13,2)
    Local $a = StringRegExp($s, "\d+", 3)
    If UBound($a) >= 3 Then
        Local $y, $m, $d, $h = 0, $min = 0, $sec = 0
        If StringLen($a[0]) = 4 Then
            $y = $a[0]
            $m = $a[1]
            $d = $a[2]
        Else
            $d = $a[0]
            $m = $a[1]
            $y = $a[2]
        EndIf
        If Number($m) > 12 Then
            Local $tmp = $m
            $m = $d
            $d = $tmp
        EndIf
        If UBound($a) > 3 Then $h = $a[3]
        If UBound($a) > 4 Then $min = $a[4]
        If UBound($a) > 5 Then $sec = $a[5]
        If StringInStr($s, "PM") And Number($h) < 12 Then $h = Number($h) + 12
        If StringInStr($s, "AM") And Number($h) = 12 Then $h = 0
        Return StringFormat("%04d/%02d/%02d %02d:%02d:%02d", $y, $m, $d, $h, $min, $sec)
    EndIf
    Return "1970/01/01 00:00:00"
EndFunc

; ============================================================================
; EDOC_HTML_BRIDGE_V13 - Interface EDOC 100% HTML avec pont AutoIt
; ============================================================================

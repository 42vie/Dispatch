; ============================================================================
; Action_EDOC.au3
; Action EDOC.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _ActionEDOC($sNumDossier)
    Local $hWndEdoc = WinGetHandle("[CLASS:TfmEdocViewerMainDlg]")
    If Not WinExists($hWndEdoc) Then Return False
    $sNumDossier = StringStripWS($sNumDossier, 8)
    ; Mode arrière-plan — pas de WinActivate
    ControlSetText($hWndEdoc, "", "[CLASS:Edit; INSTANCE:1]", $sNumDossier)
    ControlSend($hWndEdoc, "", "[CLASS:Edit; INSTANCE:1]", "{ENTER}")
    Return True
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

Func _PJPicker($idx)
    Local $oMail = $g_aGrouped[$idx][0]
    If Not IsObj($oMail) Or $oMail.Attachments.Count = 0 Then Return MsgBox(64, "PJ", "Aucune pièce jointe.")
    Local $h = GUICreate("Choisir les pièces jointes", 720, 470, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_EX_TOPMOST))
    GUISetBkColor($C_BG)
    GUISetFont(9, 400, 0, "Segoe UI")
    GUICtrlCreateLabel("Pièces jointes imprimables", 22, 18, 330, 26)
    GUICtrlSetFont(-1, 15, 800)
    GUICtrlSetColor(-1, $C_TEXT)
    GUICtrlSetBkColor(-1, $C_BG)
    Local $idList = GUICtrlCreateListView("|Nom fichier|Type", 22, 60, 675, 320, BitOR($LVS_SHOWSELALWAYS, $LVS_SINGLESEL))
    Local $hList = GUICtrlGetHandle($idList)
    _GUICtrlListView_SetExtendedListViewStyle($hList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES, $LVS_EX_DOUBLEBUFFER))
    _GUICtrlListView_SetColumnWidth($hList, 0, 38)
    _GUICtrlListView_SetColumnWidth($hList, 1, 520)
    _GUICtrlListView_SetColumnWidth($hList, 2, 90)
    Local $aMap[1], $visible = 0
    For $i = 1 To $oMail.Attachments.Count
        Local $name = $oMail.Attachments.Item($i).FileName
        If _IsPrintable($name) Then
            $visible += 1
            ReDim $aMap[$visible + 1]
            $aMap[$visible] = $i
            GUICtrlCreateListViewItem("|" & StringReplace($name, "|", "/") & "|" & _Ext($name), $idList)
            _GUICtrlListView_SetItemChecked($hList, $visible - 1, _AttIndexSelected($g_aGrouped[$idx][10], $i))
        EndIf
    Next
    If $visible = 0 Then
        GUIDelete($h)
        Return MsgBox(64, "PJ", "Aucune PJ imprimable détectée.")
    EndIf
    Local $idAll = GUICtrlCreateButton("Tout cocher", 190, 405, 105, 34)
    Local $idNone = GUICtrlCreateButton("Tout décocher", 305, 405, 120, 34)
    Local $idOK = GUICtrlCreateButton("Valider", 445, 400, 110, 42)
    GUICtrlSetBkColor($idOK, $C_ACCENT)
    Local $idCancel = GUICtrlCreateButton("Annuler", 565, 400, 100, 42)
    GUISetState(@SW_SHOW, $h)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idCancel
                GUIDelete($h)
                Return
            Case $idAll
                For $i = 0 To $visible - 1
                    _GUICtrlListView_SetItemChecked($hList, $i, True)
                Next
            Case $idNone
                For $i = 0 To $visible - 1
                    _GUICtrlListView_SetItemChecked($hList, $i, False)
                Next
            Case $idOK
                Local $sel = ""
                For $i = 0 To $visible - 1
                    If _GUICtrlListView_GetItemChecked($hList, $i) Then $sel &= $aMap[$i + 1] & "|"
                Next
                $g_aGrouped[$idx][10] = $sel
                $g_aGrouped[$idx][4] = ($sel <> "")
                GUIDelete($h)
                Return
        EndSwitch
    WEnd
EndFunc

; ==================================================================================================
; UPLOAD EDOC
; ==================================================================================================

Func _RunUpload($sCompte)
    Local $oNet = ObjCreate("WScript.Network")
    Local $total = _CountSelected(), $done = 0
    ProgressOn($EDOC_APP_TITLE, "Upload EDOC", "", -1, -1, 16)
    For $i = 0 To $g_iGroupedCount - 1
        If Not $g_aGrouped[$i][11] Then ContinueLoop
        $done += 1
        Local $oMail = $g_aGrouped[$i][0], $nums = $g_aGrouped[$i][1], $doc = $g_aGrouped[$i][2], $sec = $g_aGrouped[$i][6]
        ProgressSet(Int(($done / $total) * 100), _PipeToComma(StringTrimRight($nums, 1)), "Source " & $done & "/" & $total)
        If $g_aGrouped[$i][3] Then
            If _OpenFirstDossier($nums) Then
                _PrintMail($oMail, $oNet)
                _ValiderMultiDansEDOC($doc, $nums, $sec, $sCompte)
            EndIf
        EndIf
        If $g_aGrouped[$i][4] And $g_aGrouped[$i][10] <> "" Then
            Local $aSel = StringSplit(StringTrimRight($g_aGrouped[$i][10], 1), "|")
            For $p = 1 To $aSel[0]
                Local $attIndex = Int($aSel[$p])
                If $attIndex < 1 Or $attIndex > $oMail.Attachments.Count Then ContinueLoop
                Local $oAtt = $oMail.Attachments.Item($attIndex)
                If Not _IsPrintable($oAtt.FileName) Then ContinueLoop
                If _OpenFirstDossier($nums) Then
                    _PrintAttachment($oAtt, $oNet)
                    _ValiderMultiDansEDOC($doc, $nums, $sec, $sCompte)
                EndIf
            Next
        EndIf
        $g_sHistorique &= $g_aGrouped[$i][12]
    Next
    ProgressOff()
    MsgBox(64, "Terminé", "Upload terminé pour " & $total & " mail(s) source.")
EndFunc

Func _PrintAttachment($oAtt, $oNet)
    Local $path = $TEMP_DIR & "\" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_" & _SafeName($oAtt.FileName)
    $oAtt.SaveAsFile($path)
    Sleep(1200)
    $oNet.SetDefaultPrinter("edoc Upload")
    Sleep(700)
    ShellExecute($path, "", "", "print", @SW_HIDE)
    Sleep(3500)
    FileDelete($path)
EndFunc

Func _ValiderMultiDansEDOC($sDocType, $sListeJ, $sSec, $sCompte)
    If Not WinWait("Upload Documents CDG", "", 35) Then
        MsgBox(48, "EDOC", "La fenêtre 'Upload Documents CDG' n'est pas apparue.")
        Return False
    EndIf
    WinActivate("Upload Documents CDG")
    WinWaitActive("Upload Documents CDG", "", 8)
    Sleep(3500)
    Send("{TAB}")
    Sleep(900)
    Send($sDocType)
    Sleep(1200)
    ; TAB après Delivery Order uniquement si plusieurs numéros dans le mail.
    If _CountNums($sListeJ) > 1 Then
        Send("{TAB}")
        Sleep(900)
    EndIf
    Send("{TAB 2}")
    Sleep(1000)
    ClipPut(_NumbersClipboard($sListeJ))
    Sleep(400)
    Send("^v")
    Sleep(1200)
    ControlClick("Upload Documents CDG", "", "TButton2")
    Sleep(3000)
    IniWrite($CFG_FILE, $sSec, "LastUpload_" & $sCompte, _NowCalc())
    Return True
EndFunc

; ==================================================================================================
; UTILS
; ==================================================================================================

Func _EDOC_EnsureOutlook()
    If IsObj($g_oOutlook) And IsObj($g_oNamespace) Then Return True
    _InitConfig()
    $g_oOutlook = ObjGet("", "Outlook.Application")
    If @error Or Not IsObj($g_oOutlook) Then Return False
    $g_oNamespace = $g_oOutlook.GetNamespace("MAPI")
    Return IsObj($g_oNamespace)
EndFunc

Func _EDOC_ActionEnabled($sSpec, $sAction, $sKind)
    ; sSpec format: Action~mail~pj|Action2~mail~pj
    Local $a = StringSplit($sSpec, "|")
    For $i = 1 To $a[0]
        Local $p = StringSplit($a[$i], "~")
        If $p[0] >= 3 And $p[1] = $sAction Then
            If $sKind = "mail" Then Return ($p[2] = "1" Or StringLower($p[2]) = "true")
            If $sKind = "pj" Then Return ($p[3] = "1" Or StringLower($p[3]) = "true")
        EndIf
    Next
    Return False
EndFunc

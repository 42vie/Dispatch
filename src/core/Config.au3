; ============================================================================
; Config.au3
; Configuration et initialisation INI.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
#include-once

Func _GetPJConfig()
    Local $sIni  = @ScriptDir & "\dispatch_config.ini"
    Local $sPath = IniRead($sIni, "PJ", "Path",       "")
    Local $sRDV  = IniRead($sIni, "PJ", "RDV_Ext",    "pdf")
    Local $sUPS  = IniRead($sIni, "PJ", "UPS_Folder", "UPS")
    Local $sDGS  = IniRead($sIni, "PJ", "DGS_Folder", "DGS")
    Return StringSplit($sPath & "|" & $sRDV & "|" & $sUPS & "|" & $sDGS, "|", 1)
EndFunc

Func _InitConfig()
    If FileExists($CFG_FILE) Then Return
    IniWrite($CFG_FILE, "System", "ProfilesList", "HPE")
    IniWrite($CFG_FILE, "HPE_Actions", "List", "Pre-Alertes|Rendez-Vous")
    IniWrite($CFG_FILE, "HPE_Pre-Alertes", "Folder", "5")
    IniWrite($CFG_FILE, "HPE_Pre-Alertes", "Keywords", "Livraison,Commande")
    IniWrite($CFG_FILE, "HPE_Pre-Alertes", "Sender", "")
    IniWrite($CFG_FILE, "HPE_Pre-Alertes", "Prefix", "J")
    IniWrite($CFG_FILE, "HPE_Pre-Alertes", "Length", "9")
    IniWrite($CFG_FILE, "HPE_Pre-Alertes", "DocType", "Delivery Order")
    IniWrite($CFG_FILE, "HPE_Rendez-Vous", "Folder", "5")
    IniWrite($CFG_FILE, "HPE_Rendez-Vous", "Keywords", "RDV,HPE")
    IniWrite($CFG_FILE, "HPE_Rendez-Vous", "Sender", "")
    IniWrite($CFG_FILE, "HPE_Rendez-Vous", "Prefix", "J")
    IniWrite($CFG_FILE, "HPE_Rendez-Vous", "Length", "9")
    IniWrite($CFG_FILE, "HPE_Rendez-Vous", "DocType", "Appointment Requested")
EndFunc

Func _LoadActions($sProf, $sMail)
    For $i = 0 To $g_iDynCount - 1
        For $j = 0 To 4
            If $g_aDynCtrls[$i][$j] <> 0 Then GUICtrlDelete($g_aDynCtrls[$i][$j])
        Next
    Next
    $g_iDynCount = 0
    If $sProf = "" Then Return
    Local $sList = IniRead($CFG_FILE, $sProf & "_Actions", "List", "")
    If $sList = "" Then Return
    Local $a = StringSplit($sList, "|")
    Local $y = 338
    For $i = 1 To $a[0]
        If $i > 8 Then ExitLoop
        Local $sSec = $sProf & "_" & $a[$i]
        Local $sLast = IniRead($CFG_FILE, $sSec, "LastUpload_" & $sMail, "Jamais")
        If $sLast <> "Jamais" Then $sLast = StringRight($sLast, 8)
        $g_aDynCtrls[$g_iDynCount][0] = GUICtrlCreateCheckbox($a[$i], 55, $y, 175, 22)
        GUICtrlSetState(-1, $GUI_UNCHECKED)
        GUICtrlSetBkColor(-1, $C_CARD)
        $g_aDynCtrls[$g_iDynCount][1] = GUICtrlCreateCheckbox("", 270, $y, 22, 22)
        GUICtrlSetState(-1, $GUI_UNCHECKED)
        GUICtrlSetBkColor(-1, $C_CARD)
        $g_aDynCtrls[$g_iDynCount][2] = GUICtrlCreateCheckbox("", 360, $y, 22, 22)
        GUICtrlSetState(-1, $GUI_UNCHECKED)
        GUICtrlSetBkColor(-1, $C_CARD)
        $g_aDynCtrls[$g_iDynCount][3] = $a[$i]
        $g_aDynCtrls[$g_iDynCount][4] = GUICtrlCreateLabel("Dernier : " & $sLast, 485, $y + 3, 180, 18)
        GUICtrlSetColor(-1, $C_MUTED)
        GUICtrlSetBkColor(-1, $C_CARD)
        $y += 25
        $g_iDynCount += 1
    Next
EndFunc

Func _RulesGUI()
    Local $h = GUICreate("Règles EDOC", 585, 570, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_EX_TOPMOST))
    GUISetBkColor($C_BG)
    GUISetFont(9, 400, 0, "Segoe UI")
    GUICtrlCreateLabel("Règles de scan", 22, 18, 260, 28)
    GUICtrlSetFont(-1, 16, 800)
    GUICtrlSetColor(-1, $C_TEXT)
    GUICtrlSetBkColor(-1, $C_BG)
    GUICtrlCreateLabel("Profil", 32, 66, 100, 20)
    Local $idP = GUICtrlCreateCombo("", 32, 88, 150, 26)
    Local $idNewP = GUICtrlCreateButton("Nouveau", 190, 87, 80, 28)
    Local $idDelP = GUICtrlCreateButton("Supprimer", 278, 87, 80, 28)
    GUICtrlCreateLabel("Action", 370, 66, 100, 20)
    Local $idA = GUICtrlCreateCombo("", 370, 88, 120, 26)
    Local $idNewA = GUICtrlCreateButton("Nouvelle", 496, 87, 78, 28)
    GUICtrlCreateLabel("Dossier Outlook", 32, 135, 180, 20)
    Local $idFolder = GUICtrlCreateCombo("", 32, 157, 525, 26)
    GUICtrlSetData($idFolder, "Éléments envoyés|Boîte de réception", "Éléments envoyés")
    GUICtrlCreateLabel("Mots-clés objet, séparés par virgules", 32, 200, 300, 20)
    Local $idKeys = GUICtrlCreateInput("", 32, 222, 525, 26)
    GUICtrlCreateLabel("Préfixe", 32, 265, 80, 20)
    Local $idPrefix = GUICtrlCreateInput("", 100, 261, 90, 26)
    GUICtrlCreateLabel("Longueur après préfixe", 220, 265, 160, 20)
    Local $idLen = GUICtrlCreateInput("", 390, 261, 70, 26)
    GUICtrlCreateLabel("Filtre expéditeur optionnel", 32, 310, 220, 20)
    Local $idSender = GUICtrlCreateInput("", 32, 332, 525, 26)
    GUICtrlCreateLabel("Nom exact document EDOC", 32, 375, 220, 20)
    Local $idDoc = GUICtrlCreateInput("", 32, 397, 525, 26)
    Local $idDelA = GUICtrlCreateButton("Supprimer action", 85, 480, 135, 34)
    Local $idSave = GUICtrlCreateButton("Sauvegarder", 235, 480, 130, 34)
    GUICtrlSetBkColor($idSave, $C_ACCENT)
    Local $idClose = GUICtrlCreateButton("Fermer", 380, 480, 100, 34)
    Local $sPList = IniRead($CFG_FILE, "System", "ProfilesList", "")
    If $sPList <> "" Then GUICtrlSetData($idP, "|" & $sPList, StringSplit($sPList, "|")[1])
    _ReloadActionsCombo($idP, $idA)
    _LoadActionData($idP, $idA, $idFolder, $idKeys, $idPrefix, $idLen, $idSender, $idDoc)
    GUISetState(@SW_SHOW, $h)
    Local $oldP = GUICtrlRead($idP), $oldA = GUICtrlRead($idA)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idClose
                GUIDelete($h)
                Return
            Case $idNewP
                Local $sNewP = InputBox("Nouveau profil", "Nom du profil :")
                If $sNewP <> "" Then
                    Local $sCurrP = IniRead($CFG_FILE, "System", "ProfilesList", "")
                    If Not _PipeContains($sCurrP, $sNewP) Then
                        If $sCurrP = "" Then
                            $sCurrP = $sNewP
                        Else
                            $sCurrP &= "|" & $sNewP
                        EndIf
                        IniWrite($CFG_FILE, "System", "ProfilesList", $sCurrP)
                        IniWrite($CFG_FILE, $sNewP & "_Actions", "List", "")
                        GUICtrlSetData($idP, "|" & $sCurrP, $sNewP)
                    EndIf
                EndIf
            Case $idDelP
                Local $sP = GUICtrlRead($idP)
                If $sP <> "" And MsgBox($MB_YESNO + $MB_ICONWARNING, "Confirmer", "Supprimer le profil " & $sP & " ?") = $IDYES Then
                    Local $sActs = IniRead($CFG_FILE, $sP & "_Actions", "List", "")
                    Local $aActs = StringSplit($sActs, "|")
                    For $i = 1 To $aActs[0]
                        If $aActs[$i] <> "" Then IniDelete($CFG_FILE, $sP & "_" & $aActs[$i])
                    Next
                    IniDelete($CFG_FILE, $sP & "_Actions")
                    Local $sNewList = _RemoveFromPipeList(IniRead($CFG_FILE, "System", "ProfilesList", ""), $sP)
                    IniWrite($CFG_FILE, "System", "ProfilesList", $sNewList)
                    GUICtrlSetData($idP, "|" & $sNewList, "")
                    _ReloadActionsCombo($idP, $idA)
                EndIf
            Case $idNewA
                Local $sP2 = GUICtrlRead($idP)
                If $sP2 = "" Then ContinueLoop
                Local $sNewA = InputBox("Nouvelle action", "Nom de l'action :")
                If $sNewA <> "" Then
                    Local $sCurrA = IniRead($CFG_FILE, $sP2 & "_Actions", "List", "")
                    If Not _PipeContains($sCurrA, $sNewA) Then
                        If $sCurrA = "" Then
                            $sCurrA = $sNewA
                        Else
                            $sCurrA &= "|" & $sNewA
                        EndIf
                        IniWrite($CFG_FILE, $sP2 & "_Actions", "List", $sCurrA)
                        GUICtrlSetData($idA, "|" & $sCurrA, $sNewA)
                    EndIf
                EndIf
            Case $idDelA
                Local $sP3 = GUICtrlRead($idP), $sA3 = GUICtrlRead($idA)
                If $sP3 <> "" And $sA3 <> "" And MsgBox($MB_YESNO, "Confirmer", "Supprimer l'action " & $sA3 & " ?") = $IDYES Then
                    Local $sNewAList = _RemoveFromPipeList(IniRead($CFG_FILE, $sP3 & "_Actions", "List", ""), $sA3)
                    IniWrite($CFG_FILE, $sP3 & "_Actions", "List", $sNewAList)
                    IniDelete($CFG_FILE, $sP3 & "_" & $sA3)
                    _ReloadActionsCombo($idP, $idA)
                    _LoadActionData($idP, $idA, $idFolder, $idKeys, $idPrefix, $idLen, $idSender, $idDoc)
                EndIf
            Case $idSave
                Local $sP4 = GUICtrlRead($idP), $sA4 = GUICtrlRead($idA)
                If $sP4 <> "" And $sA4 <> "" Then
                    Local $sSec = $sP4 & "_" & $sA4
                    Local $sFolderID = "5"
                    If GUICtrlRead($idFolder) = "Boîte de réception" Then $sFolderID = "6"
                    IniWrite($CFG_FILE, $sSec, "Folder", $sFolderID)
                    IniWrite($CFG_FILE, $sSec, "Keywords", GUICtrlRead($idKeys))
                    IniWrite($CFG_FILE, $sSec, "Prefix", GUICtrlRead($idPrefix))
                    IniWrite($CFG_FILE, $sSec, "Length", GUICtrlRead($idLen))
                    IniWrite($CFG_FILE, $sSec, "Sender", GUICtrlRead($idSender))
                    IniWrite($CFG_FILE, $sSec, "DocType", GUICtrlRead($idDoc))
                    MsgBox(64, "OK", "Règle sauvegardée.", 2, $h)
                EndIf
        EndSwitch
        If GUICtrlRead($idP) <> $oldP Then
            $oldP = GUICtrlRead($idP)
            _ReloadActionsCombo($idP, $idA)
            _LoadActionData($idP, $idA, $idFolder, $idKeys, $idPrefix, $idLen, $idSender, $idDoc)
        EndIf
        If GUICtrlRead($idA) <> $oldA Then
            $oldA = GUICtrlRead($idA)
            _LoadActionData($idP, $idA, $idFolder, $idKeys, $idPrefix, $idLen, $idSender, $idDoc)
        EndIf
    WEnd
EndFunc

Func _ReloadActionsCombo($idP, $idA)
    Local $s = IniRead($CFG_FILE, GUICtrlRead($idP) & "_Actions", "List", "")
    Local $first = ""
    If $s <> "" Then $first = StringSplit($s, "|")[1]
    GUICtrlSetData($idA, "|" & $s, $first)
EndFunc

Func _LoadActionData($idP, $idA, $idF, $idK, $idPre, $idLen, $idS, $idD)
    Local $sSec = GUICtrlRead($idP) & "_" & GUICtrlRead($idA)
    If IniRead($CFG_FILE, $sSec, "Folder", "5") = "6" Then
        GUICtrlSetData($idF, "Boîte de réception")
    Else
        GUICtrlSetData($idF, "Éléments envoyés")
    EndIf
    GUICtrlSetData($idK, IniRead($CFG_FILE, $sSec, "Keywords", ""))
    GUICtrlSetData($idPre, IniRead($CFG_FILE, $sSec, "Prefix", ""))
    GUICtrlSetData($idLen, IniRead($CFG_FILE, $sSec, "Length", ""))
    GUICtrlSetData($idS, IniRead($CFG_FILE, $sSec, "Sender", ""))
    GUICtrlSetData($idD, IniRead($CFG_FILE, $sSec, "DocType", ""))
EndFunc

; ==================================================================================================
; SCAN OUTLOOK + SELECTION
; ==================================================================================================

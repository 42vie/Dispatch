#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>

; ==================================================================================================
; EDOC MASTER BOT V15.6 FULL
; - Garde robot_v15_config.ini
; - Design blocs nets : bordures visibles + titres sur fond bloc
; - Accueil : aucune action / Mail / PJ cochée par défaut
; - Sélection : 1 ligne = 1 mail source, Mail + PJ sur la même ligne
; - EDOC : TAB après le type document uniquement si plusieurs numéros dans le mail
; ==================================================================================================

Global $APP_TITLE = "EDOC Master Bot V15.6"
Global $CFG_FILE = @ScriptDir & "\robot_v15_config.ini"
Global $TEMP_DIR = @TempDir & "\Temp_EDOC_Robot_V15"
Global $EDOC_WM_NOTIFY = 0x004E, $EDOC_NM_DBLCLK = -3
Global $C_BG = 0xF3F4F6, $C_CARD = 0xFFFFFF, $C_BORDER = 0xCBD5E1, $C_TEXT = 0x111827, $C_MUTED = 0x6B7280, $C_ACCENT = 0xDCFCE7, $C_BUTTON = 0xE5E7EB
Global $g_oOutlook = 0, $g_oNamespace = 0, $g_sHistorique = "|"
Global $g_aDynCtrls[30][6], $g_iDynCount = 0
Global $g_aGrouped[800][15], $g_iGroupedCount = 0
Global $g_idSelList = 0, $g_hSelList = 0

ObjEvent("AutoIt.Error", "_ComErr")
If Not FileExists($TEMP_DIR) Then DirCreate($TEMP_DIR)
_InitConfig()
$g_oOutlook = ObjGet("", "Outlook.Application")
If @error Or Not IsObj($g_oOutlook) Then
    MsgBox($MB_ICONERROR, "Outlook", "Veuillez ouvrir Outlook avant de lancer le script.")
    Exit
EndIf
$g_oNamespace = $g_oOutlook.GetNamespace("MAPI")
Local $sStoresList = _GetStoresList($g_oNamespace)
If $sStoresList = "" Then
    MsgBox($MB_ICONERROR, "Outlook", "Aucune boîte mail Outlook détectée.")
    Exit
EndIf

; ========================= INTERFACE PRINCIPALE =========================
Local $hMain = GUICreate($APP_TITLE, 790, 665)
GUISetBkColor($C_BG)
GUISetFont(9, 400, 0, "Segoe UI")
GUICtrlCreateLabel("EDOC Master Bot", 24, 18, 360, 34)
GUICtrlSetFont(-1, 20, 800, 0, "Segoe UI")
GUICtrlSetColor(-1, $C_TEXT)
GUICtrlSetBkColor(-1, $C_BG)
GUICtrlCreateLabel("V15.6 - blocs nets - multi-dossiers EDOC", 28, 54, 650, 20)
GUICtrlSetColor(-1, $C_MUTED)
GUICtrlSetBkColor(-1, $C_BG)
Local $idBtnRules = GUICtrlCreateButton("Règles", 655, 24, 100, 34)
GUICtrlSetBkColor($idBtnRules, $C_BUTTON)

_Box(20, 92, 750, 90)
_SectionTitle("1. Source Outlook", 40, 108, 260)
_MutedLabel("Boîte mail", 40, 138, 110)
Local $idComboMail = GUICtrlCreateCombo("", 155, 134, 585, 26)
GUICtrlSetData($idComboMail, $sStoresList, StringSplit($sStoresList, "|")[1])

_Box(20, 198, 750, 228)
_SectionTitle("2. Profil et actions", 40, 214, 260)
_MutedLabel("Profil client", 40, 244, 100)
Local $idComboProfile = GUICtrlCreateCombo("", 155, 240, 260, 26)
_MutedLabel("Rien n'est coché par défaut : coche l'action puis Mail et/ou PJ.", 40, 278, 660)
_HeaderLabel("Action", 55, 313, 150)
_HeaderLabel("Mail", 270, 313, 70)
_HeaderLabel("PJ", 360, 313, 70)
_HeaderLabel("Dernier upload", 485, 313, 160)

_Box(20, 442, 750, 116)
_SectionTitle("3. Période", 40, 458, 180)
_MutedLabel("Période", 40, 492, 80)
Local $idComboTemps = GUICtrlCreateCombo("", 130, 488, 245, 26)
GUICtrlSetData($idComboTemps, "Dernier upload (Automatique)|Aujourd'hui|Il y a 1h|Personnalisé", "Dernier upload (Automatique)")
_MutedLabel("Du", 398, 492, 30)
Local $idInputStart = GUICtrlCreateInput("", 432, 488, 300, 26)
Local $idLblAu = _MutedLabel("Au", 398, 526, 30)
Local $idInputFin = GUICtrlCreateInput(@YEAR & "/" & @MON & "/" & @MDAY & " 23:59:59", 432, 522, 300, 26)

Local $idBtnStart = GUICtrlCreateButton("Scanner puis sélectionner les uploads", 215, 586, 360, 46)
GUICtrlSetFont(-1, 11, 800)
GUICtrlSetBkColor(-1, $C_ACCENT)
Local $idStatus = GUICtrlCreateLabel("Prêt.", 24, 642, 730, 18)
GUICtrlSetColor($idStatus, $C_MUTED)
GUICtrlSetBkColor($idStatus, $C_BG)

_LoadProfiles($idComboProfile, GUICtrlRead($idComboMail))
_UpdateDates($idComboTemps, $idInputStart, $idLblAu, $idInputFin)
GUISetState(@SW_SHOW, $hMain)

Local $sLastMail = GUICtrlRead($idComboMail), $sLastProf = GUICtrlRead($idComboProfile), $sLastTime = GUICtrlRead($idComboTemps)
While 1
    Local $msg = GUIGetMsg()
    Switch $msg
        Case $GUI_EVENT_CLOSE
            Exit
        Case $idBtnRules
            GUISetState(@SW_DISABLE, $hMain)
            _RulesGUI()
            _LoadProfiles($idComboProfile, GUICtrlRead($idComboMail))
            GUISetState(@SW_ENABLE, $hMain)
            GUISetState(@SW_RESTORE, $hMain)
        Case $idBtnStart
            GUICtrlSetState($idBtnStart, $GUI_DISABLE)
            GUICtrlSetData($idStatus, "Scan Outlook en cours...")
            _MainProcess(GUICtrlRead($idComboMail), GUICtrlRead($idComboProfile), GUICtrlRead($idComboTemps), GUICtrlRead($idInputStart), GUICtrlRead($idInputFin))
            GUICtrlSetData($idStatus, "Prêt.")
            GUICtrlSetState($idBtnStart, $GUI_ENABLE)
    EndSwitch
    If GUICtrlRead($idComboProfile) <> $sLastProf Or GUICtrlRead($idComboMail) <> $sLastMail Then
        $sLastProf = GUICtrlRead($idComboProfile)
        $sLastMail = GUICtrlRead($idComboMail)
        _LoadActions($sLastProf, $sLastMail)
    EndIf
    If GUICtrlRead($idComboTemps) <> $sLastTime Then
        $sLastTime = GUICtrlRead($idComboTemps)
        _UpdateDates($idComboTemps, $idInputStart, $idLblAu, $idInputFin)
    EndIf
WEnd

; ========================= DESIGN / CONFIG =========================
Func _ComErr()
    Return
EndFunc
Func _Box($x, $y, $w, $h)
    Local $idCard = GUICtrlCreateLabel("", $x, $y, $w, $h)
    GUICtrlSetBkColor($idCard, $C_CARD)
    GUICtrlSetState($idCard, $GUI_DISABLE)
    Local $idTop = GUICtrlCreateLabel("", $x, $y, $w, 1)
    GUICtrlSetBkColor($idTop, $C_BORDER)
    GUICtrlSetState($idTop, $GUI_DISABLE)
    Local $idBottom = GUICtrlCreateLabel("", $x, $y + $h - 1, $w, 1)
    GUICtrlSetBkColor($idBottom, $C_BORDER)
    GUICtrlSetState($idBottom, $GUI_DISABLE)
    Local $idLeft = GUICtrlCreateLabel("", $x, $y, 1, $h)
    GUICtrlSetBkColor($idLeft, $C_BORDER)
    GUICtrlSetState($idLeft, $GUI_DISABLE)
    Local $idRight = GUICtrlCreateLabel("", $x + $w - 1, $y, 1, $h)
    GUICtrlSetBkColor($idRight, $C_BORDER)
    GUICtrlSetState($idRight, $GUI_DISABLE)
    Return $idCard
EndFunc
Func _SectionTitle($sText, $x, $y, $w)
    Local $id = GUICtrlCreateLabel($sText, $x, $y, $w, 20)
    GUICtrlSetFont($id, 10, 700, 0, "Segoe UI")
    GUICtrlSetColor($id, $C_TEXT)
    GUICtrlSetBkColor($id, $C_CARD)
    Return $id
EndFunc
Func _MutedLabel($sText, $x, $y, $w, $h = 20)
    Local $id = GUICtrlCreateLabel($sText, $x, $y, $w, $h)
    GUICtrlSetColor($id, $C_MUTED)
    GUICtrlSetBkColor($id, $C_CARD)
    Return $id
EndFunc
Func _HeaderLabel($sText, $x, $y, $w)
    Local $id = GUICtrlCreateLabel($sText, $x, $y, $w, 18)
    GUICtrlSetFont($id, 8, 700, 0, "Segoe UI")
    GUICtrlSetColor($id, $C_TEXT)
    GUICtrlSetBkColor($id, $C_CARD)
    Return $id
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
Func _GetStoresList($oNamespace)
    Local $s = ""
    For $oStore In $oNamespace.Stores
        $s &= $oStore.DisplayName & "|"
    Next
    If StringRight($s, 1) = "|" Then $s = StringTrimRight($s, 1)
    Return $s
EndFunc
Func _LoadProfiles($idCombo, $sMail)
    Local $sList = IniRead($CFG_FILE, "System", "ProfilesList", "")
    If $sList = "" Then Return
    GUICtrlSetData($idCombo, "|" & $sList, StringSplit($sList, "|")[1])
    _LoadActions(GUICtrlRead($idCombo), $sMail)
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
Func _ApplySelection()
    For $i = 0 To $g_iGroupedCount - 1
        If _GUICtrlListView_GetItemChecked($g_hSelList, $i) And ($g_aGrouped[$i][3] Or ($g_aGrouped[$i][4] And $g_aGrouped[$i][10] <> "")) Then
            $g_aGrouped[$i][11] = True
        Else
            $g_aGrouped[$i][11] = False
        EndIf
    Next
EndFunc
Func _OpenSelectedMail()
    Local $idx = _SelIndex()
    If $idx < 0 Or $idx >= $g_iGroupedCount Then Return MsgBox(48, "Outlook", "Sélectionne d'abord une ligne.")
    _OpenMail($idx)
EndFunc
Func _OpenMail($idx)
    Local $oMail = 0, $entry = $g_aGrouped[$idx][13], $store = $g_aGrouped[$idx][14]
    If IsObj($g_oNamespace) And $entry <> "" Then
        If $store <> "" Then
            $oMail = $g_oNamespace.GetItemFromID($entry, $store)
        Else
            $oMail = $g_oNamespace.GetItemFromID($entry)
        EndIf
    EndIf
    If Not IsObj($oMail) Then $oMail = $g_aGrouped[$idx][0]
    If IsObj($oMail) Then
        $oMail.Display()
    Else
        MsgBox(48, "Outlook", "Impossible d'ouvrir le mail source.")
    EndIf
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
    ProgressOn($APP_TITLE, "Upload EDOC", "", -1, -1, 16)
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
Func _PrintMail($oMail, $oNet)
    Local $oT = $oMail.Copy()
    $oT.BodyFormat = 2
    $oT.HTMLBody = "<style>body,p,td,div{font-size:8.5pt !important;line-height:1.0 !important;margin:0 !important;padding:0 !important;}img{max-width:40% !important;}</style>" & $oT.HTMLBody
    $oT.Save()
    Local $oW = $oT.GetInspector.WordEditor
    If IsObj($oW) Then
        $oW.PageSetup.TopMargin = 10
        $oW.PageSetup.BottomMargin = 10
    EndIf
    $oNet.SetDefaultPrinter("edoc Upload")
    Sleep(700)
    $oT.PrintOut()
    Sleep(1800)
    $oT.Close(1)
    $oT.Delete()
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
Func _Ext($file)
    Local $p = StringInStr($file, ".", 0, -1)
    If $p = 0 Then Return ""
    Return StringLower(StringTrimLeft($file, $p))
EndFunc
Func _SafeName($s)
    Local $bad = '<>:"/\|?*'
    For $i = 1 To StringLen($bad)
        $s = StringReplace($s, StringMid($bad, $i, 1), "_")
    Next
    Return $s
EndFunc
Func _NumbersClipboard($nums)
    Local $a = StringSplit(StringTrimRight($nums, 1), "|")
    Local $s = ""
    For $i = 1 To $a[0]
        If StringStripWS($a[$i], 3) <> "" Then $s &= $a[$i] & @CRLF
    Next
    Return $s
EndFunc
Func _PipeToComma($s)
    Return StringReplace($s, "|", ", ")
EndFunc
Func _PipeContains($list, $item)
    Return StringInStr("|" & $list & "|", "|" & $item & "|") > 0
EndFunc
Func _RemoveFromPipeList($list, $item)
    Local $s = "|" & $list & "|"
    $s = StringReplace($s, "|" & $item & "|", "|")
    $s = StringReplace($s, "||", "|")
    If StringLeft($s, 1) = "|" Then $s = StringTrimLeft($s, 1)
    If StringRight($s, 1) = "|" Then $s = StringTrimRight($s, 1)
    Return $s
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

; ============================================================================
; EDOC_Master.au3
; EDOC Master GUI/automation.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _EDOCMasterGUI()
ObjEvent("AutoIt.Error", "_ComErr")
If Not FileExists($TEMP_DIR) Then DirCreate($TEMP_DIR)
_InitConfig()
$g_oOutlook = ObjGet("", "Outlook.Application")
If @error Or Not IsObj($g_oOutlook) Then
    MsgBox($MB_ICONERROR, "Outlook", "Veuillez ouvrir Outlook avant de lancer le script.")
    Return False
EndIf
$g_oNamespace = $g_oOutlook.GetNamespace("MAPI")
Local $sStoresList = _GetStoresList($g_oNamespace)
If $sStoresList = "" Then
    MsgBox($MB_ICONERROR, "Outlook", "Aucune boîte mail Outlook détectée.")
    Return False
EndIf
Local $hMain = GUICreate($EDOC_APP_TITLE, 790, 665)
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
            Return False
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
    Return True
EndFunc

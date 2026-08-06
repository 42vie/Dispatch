; ============================================================================
; Action_FC.au3
; File Closing unitaire.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _FC_SmartSleep($iMs)
    Local $iSlept = 0
    While $iSlept < $iMs
        _Tracker_PollButtons()
        If $bFC_Stop Or $bFC_Skip Then Return
        _FC_WaitIfPaused()
        If $bFC_Stop Or $bFC_Skip Then Return
        Sleep(100)
        $iSlept += 100
    WEnd
EndFunc

Func _FC_ResolveCarrier($sTransp)
    ; UPS → retourner marqueur spécial
    If StringInStr($sTransp, "UPS") Then Return "UPS"
    ; EFDS / Groussard → numéro direct dans transp (ex: "13")
    If StringInStr($sTransp, "EFDS") Or StringInStr($sTransp, "Groussard") Then
        ; Extraire le numéro entre parenthèses ex: "EFDS (1)" → 1, "Groussard (12)" → 12
        Local $aM2 = StringRegExp($sTransp, "\((\d+)\)", 3)
        If IsArray($aM2) Then Return $aM2[0]
        ; Fallback si pas de parenthèses : EFDS=1, Groussard=12
        If StringInStr($sTransp, "EFDS") And Not StringInStr($sTransp, "Groussard") Then Return "1"
        Return "12"
    EndIf
    ; Autres : extraire numéro entre parenthèses ex: "DGS (8)"
    Local $aM = StringRegExp($sTransp, "\((\d+)\)", 3)
    If IsArray($aM) Then Return $aM[0]
    Return "13" ; DGS par défaut
EndFunc

; ==============================================================================
; FILE CLOSING — BATCH (Colonne 5)
; ==============================================================================

Func _Run_FileClosing_UPS($Num)
    $Num = StringStripWS($Num, 8)
    If $Num = "" Then Return

    Local Const $sFC_LOG      = "[CLASS:TEIEdit; INSTANCE:91]"
    Local Const $sFC_TOOLBAR  = "[CLASS:TRzToolbar; INSTANCE:1]"
    Local Const $sFC_FILEOPEN = "[CLASS:TRzShellOpenSaveForm]"
    Local Const $sFC_MENU     = "[CLASS:TEIInputQueryForm; REGEXPTITLE:(?i).*MENU SELECTION.*]"
    Local Const $sFC_INPUT    = "[CLASS:TInputQueryForm]"
    Local Const $sFC_EDS      = "F:\Scripting\Export\EXPORT_HPE_FILECLOSING_001\EXPORT_HPE_FILECLOSING_031.eds"

    _FC_AuditInit("FC-UPS | Num=" & $Num)
    _FC_AuditFileCheck($sFC_EDS)
    Local $tTotal = TimerInit()

    Local $hWnd = _GetWindowETMS()
    If $hWnd = 0 Then
        _FC_AuditLog("*** ERREUR : E.TMS introuvable (hWnd=0) ***")
        _FC_AuditShow($Num)
        Return
    EndIf
    _FC_AuditLog("E.TMS hwnd=" & $hWnd)
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    $bFC_Stop = False
    $bFC_Pause = False

    ; ── 1. LOG J ─────────────────────────────────────────────────────────────
    $iFC_StepCurrent = 1
    _FC_AuditStep(1, "LOG J")
    Local $t1 = TimerInit()
    _Spinner("FC-UPS [" & $Num & "] 1/5 - LOG J...")
    _FC_AuditCtrl($hWnd, $sFC_LOG, "LOG avant")
    ControlSetText($hWnd, "", $sFC_LOG, "LOG " & $Num)
    _FC_SmartSleep(100)
    _FC_AuditCtrl($hWnd, $sFC_LOG, "LOG apres")
    ControlSend($hWnd, "", $sFC_LOG, "{F8}")
    _FC_SmartSleep(1500)
    _FC_AuditTiming("Step1-LOGJ", TimerDiff($t1))
    If $bFC_Stop Or $bFC_Skip Then
        _FC_AuditLog("*** STOP/SKIP par utilisateur Step 1 ***")
        _FC_AuditShow($Num)
        Return
    EndIf

    ; ── 2. Toolbar EDS ───────────────────────────────────────────────────────
    $iFC_StepCurrent = 2
    _FC_AuditStep(2, "Toolbar EDS click")
    Local $t2 = TimerInit()
    _Spinner("FC-UPS [" & $Num & "] 2/5 - Lancement EDS...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    _FC_SmartSleep(200)
    _FC_AuditWinState("[CLASS:TfmBrowser]", "ETMS avant click toolbar")
    ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
    _FC_AuditTiming("Step2-ToolbarClick", TimerDiff($t2))
    _FC_WaitIfPaused()
    If $bFC_Stop Then
        _FC_AuditLog("*** STOP par utilisateur Step 2 ***")
        _FC_AuditShow($Num)
        Return
    EndIf

    ; ── 3. FileOpen (retry x3) ────────────────────────────────────────────────
    $iFC_StepCurrent = 3
    _FC_AuditStep(3, "FileOpen dialog")
    Local $bFileOK = False
    Local $iTentative = 0
    While Not $bFileOK And $iTentative < 3
        $iTentative += 1
        Local $t3 = TimerInit()
        _FC_AuditLog("  Tentative " & $iTentative & "/3")
        _Spinner("FC-UPS [" & $Num & "] 3/5 - FileOpen (essai " & $iTentative & "/3)...")
        If $iTentative > 1 Then
            If WinExists($sFC_FILEOPEN) Then WinClose($sFC_FILEOPEN)
            WinWaitClose($sFC_FILEOPEN, "", 3)
            WinActivate($hWnd)
            WinWaitActive($hWnd, "", 3)
            Sleep(500)
            ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
        EndIf
        Local $iTimer = TimerInit()
        While Not WinExists($sFC_FILEOPEN)
            Sleep(100)
            If _IsPressed("1B") Then
                _FC_AuditLog("*** ECHAP par utilisateur pendant attente FileOpen ***")
                _FC_AuditShow($Num)
                Return
            EndIf
            If TimerDiff($iTimer) > 10000 Then
                _FC_AuditLog("  TIMEOUT 10s : FileOpen ne s'ouvre pas")
                ExitLoop
            EndIf
        WEnd
        _FC_AuditTiming("Attente apparition FileOpen", TimerDiff($iTimer))
        If Not WinExists($sFC_FILEOPEN) Then
            _FC_AuditLog("  FileOpen toujours absent apres timeout")
            _FC_AuditWinState($sFC_FILEOPEN, "FileOpen")
            ContinueLoop
        EndIf
        WinActivate($sFC_FILEOPEN)
        WinWaitActive($sFC_FILEOPEN, "", 3)
        _FC_AuditWinState($sFC_FILEOPEN, "FileOpen ouvert")
        Sleep(100)
        ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", "")
        Sleep(100)
        ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", $sFC_EDS)
        Sleep(200)
        Local $sReadBack = ControlGetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]")
        _FC_AuditLog("  Champ FileOpen apres ecriture = '" & $sReadBack & "'")
        If Not StringInStr($sReadBack, "EXPORT_HPE_FILECLOSING") Then
            _FC_AuditLog("  *** ECHEC : le texte n'a pas ete ecrit correctement ***")
            ContinueLoop
        EndIf
        Send("{ENTER}")
        Local $iWait = TimerInit()
        While WinExists($sFC_FILEOPEN)
            Sleep(100)
            If TimerDiff($iWait) > 5000 Then ExitLoop
        WEnd
        _FC_AuditTiming("Fermeture FileOpen apres ENTER", TimerDiff($iWait))
        If Not WinExists($sFC_FILEOPEN) Then
            $bFileOK = True
            _FC_AuditLog("  FileOpen OK, fichier accepte")
        Else
            _FC_AuditLog("  *** FileOpen toujours ouvert apres 5s — fichier refuse ? ***")
            _FC_AuditWinState($sFC_FILEOPEN, "FileOpen bloque")
        EndIf
        _FC_AuditTiming("Step3-Tentative" & $iTentative, TimerDiff($t3))
    WEnd
    If Not $bFileOK Then
        _FC_AuditLog("*** ECHEC FINAL : 3 tentatives FileOpen echouees ***")
        _FC_AuditTiming("TOTAL", TimerDiff($tTotal))
        _FC_AuditShow($Num)
        MsgBox(16+262144, "Erreur FC-UPS", "Impossible d'ouvrir le fichier EDS." & @CRLF & "Dossier : " & $Num)
        $bFC_Stop = True
        Return
    EndIf
    _FC_WaitIfPaused()
    If $bFC_Stop Then
        _FC_AuditLog("*** STOP par utilisateur Step 3 ***")
        _FC_AuditShow($Num)
        Return
    EndIf

    ; ── 4. Menu Selection ────────────────────────────────────────────────────
    $iFC_StepCurrent = 4
    _FC_AuditStep(4, "Menu Selection")
    Local $t4 = TimerInit()
    _WinWaitSpinner($sFC_MENU, "FC-UPS [" & $Num & "] 4/5 - Menu Selection...")
    _FC_AuditTiming("Attente Menu Selection", TimerDiff($t4))
    If $bFC_Stop Then
        _FC_AuditLog("*** STOP Step 4 ***")
        _FC_AuditShow($Num)
        Return
    EndIf
    _FC_AuditWinState($sFC_MENU, "Menu Selection")
    Local $hMenu = WinActivate($sFC_MENU)
    WinWaitActive($hMenu, "", 3)
    _FC_SmartSleep(100)
    ControlSetText($hMenu, "", "[CLASS:TEdit; INSTANCE:1]", "1")
    _FC_SmartSleep(150)
    ControlClick($hMenu, "", "[TEXT:OK]")
    WinWaitClose($hMenu, "", 5)
    _FC_SmartSleep(200)
    _FC_AuditTiming("Step4-MenuSelection", TimerDiff($t4))
    If $bFC_Stop Or $bFC_Skip Then
        _FC_AuditLog("*** STOP/SKIP Step 4 apres ***")
        _FC_AuditShow($Num)
        Return
    EndIf

    ; ── 5. Numéro J ──────────────────────────────────────────────────────────
    $iFC_StepCurrent = 5
    _FC_AuditStep(5, "Numero J = " & $Num)
    Local $t5 = TimerInit()
    _WinWaitSpinner($sFC_INPUT, "FC-UPS [" & $Num & "] 5/5 - Numéro J...")
    _FC_AuditTiming("Attente Input Num J", TimerDiff($t5))
    If $bFC_Stop Or $bFC_Skip Then
        _FC_AuditLog("*** STOP/SKIP Step 5 ***")
        _FC_AuditShow($Num)
        Return
    EndIf
    Local $hInput1 = WinActivate($sFC_INPUT)
    WinWaitActive($hInput1, "", 3)
    _FC_SmartSleep(100)
    ControlSetText($hInput1, "", "[CLASS:TEdit; INSTANCE:1]", $Num)
    _FC_SmartSleep(150)
    ControlClick($hInput1, "", "[TEXT:OK]")
    WinWaitClose($hInput1, "", 5)
    _FC_SmartSleep(1500) ; E.TMS charge
    _FC_AuditTiming("Step5-NumeroJ", TimerDiff($t5))
    If $bFC_Stop Or $bFC_Skip Then
        _FC_AuditLog("*** STOP/SKIP Step 5 apres ***")
        _FC_AuditShow($Num)
        Return
    EndIf

    ; ── 6. UPS = pas de Carrier ID → première popup = DEF → terminé ──────────
    $iFC_StepCurrent = 6
    _FC_AuditStep(6, "DEF")
    Local $t6 = TimerInit()
    _WinWaitSpinner($sFC_INPUT, "FC-UPS [" & $Num & "] DEF...")
    If $bFC_Stop Or $bFC_Skip Then
        _FC_AuditLog("*** STOP/SKIP Step 6 ***")
        _FC_AuditShow($Num)
        Return
    EndIf
    Local $hDef = WinActivate($sFC_INPUT)
    WinWaitActive($hDef, "", 3)
    _FC_SmartSleep(100)
    ControlSetText($hDef, "", "[CLASS:TEdit; INSTANCE:1]", "DEF")
    _FC_SmartSleep(150)
    ControlClick($hDef, "", "[TEXT:OK]")
    WinWaitClose($hDef, "", 5)
    _FC_AuditTiming("Step6-DEF", TimerDiff($t6))

    ; ── Attendre le script auto E.TMS (min 20s) ──────────────────────────────
    _Spinner("FC-UPS [" & $Num & "] Script auto en cours... (20s)")
    _FC_SmartSleep(20000)
    If $bFC_Stop Or $bFC_Skip Then
        _FC_AuditLog("*** STOP/SKIP pendant attente script auto ***")
        _FC_AuditShow($Num)
        Return
    EndIf

    ; ── C'EST TOUT POUR UPS ──────────────────────────────────────────────────
    _FC_AuditTiming("TOTAL FC-UPS", TimerDiff($tTotal))
    _FC_AuditLog("====== FIN FC-UPS OK ======")
    _FC_AuditSave($Num)
    $iFC_StepCurrent = 0
    ToolTip("")
EndFunc

; ==============================================================================
; FILE CLOSING STANDARD (DGS, EFDS, Groussard, etc.)
; ==============================================================================

Func _Run_FileClosing_Single($Num, $CarrierID = "13", $DateGOverride = "", $Horaire = "09h et 12h", $Notes = "", $DLY = "", $DLYNotes = "", $iStartStep = 1)
    Local Const $sFC_LOG      = "[CLASS:TEIEdit; INSTANCE:91]"
    Local Const $sFC_TOOLBAR  = "[CLASS:TRzToolbar; INSTANCE:1]"
    Local Const $sFC_FILEOPEN = "[CLASS:TRzShellOpenSaveForm]"
    Local Const $sFC_MENU     = "[CLASS:TEIInputQueryForm; REGEXPTITLE:(?i).*MENU SELECTION.*]"
    Local Const $sFC_INPUT    = "[CLASS:TInputQueryForm]"
    Local Const $sFC_CARRIER  = "[CLASS:TEIInputQueryForm; REGEXPTITLE:(?i).*Carrier ID.*]"
    Local Const $sFC_EDS      = "F:\Scripting\Export\EXPORT_HPE_FILECLOSING_001\EXPORT_HPE_FILECLOSING_031.eds"

    _FC_AuditInit("FC-Single | Num=" & $Num & " | Carrier=" & $CarrierID & " | StartStep=" & $iStartStep)
    _FC_AuditFileCheck($sFC_EDS)
    _FC_AuditLog("Params: DateG=" & $DateGOverride & " Horaire=" & $Horaire & " DLY=" & $DLY)
    Local $tTotal = TimerInit()

    Local $hWnd = _GetWindowETMS()
    If $hWnd = 0 Then
        _FC_AuditLog("*** ERREUR : E.TMS introuvable (hWnd=0) ***")
        _FC_AuditShow($Num)
        Return
    EndIf
    _FC_AuditLog("E.TMS hwnd=" & $hWnd)
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    $bFC_Stop  = False
    $bFC_Pause = False

    If $iStartStep <= 1 Then
        $iFC_StepCurrent = 1
        _FC_AuditStep(1, "LOG J")
        Local $t1 = TimerInit()
        _Spinner("FC [" & $Num & "] 1/7 - LOG J...")
        ControlSetText($hWnd, "", $sFC_LOG, "LOG " & $Num)
        _FC_SmartSleep(300)
        ControlSend($hWnd, "", $sFC_LOG, "{F8}")
        _FC_SmartSleep(3000)
        _FC_AuditTiming("Step1-LOGJ", TimerDiff($t1))
        If $bFC_Stop Or $bFC_Skip Then
            _FC_AuditLog("*** STOP Step 1 ***")
            _FC_AuditShow($Num)
            Return
        EndIf
    EndIf

    If $iStartStep <= 2 Then
        $iFC_StepCurrent = 2
        _FC_AuditStep(2, "Toolbar EDS")
        Local $t2 = TimerInit()
        _Spinner("FC [" & $Num & "] 2/7 - Lancement EDS...")
        WinActivate($hWnd)
        WinWaitActive($hWnd, "", 3)
        _FC_SmartSleep(500)
        ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
        _FC_AuditTiming("Step2-Toolbar", TimerDiff($t2))
        If $bFC_Stop Or $bFC_Skip Then
            _FC_AuditLog("*** STOP Step 2 ***")
            _FC_AuditShow($Num)
            Return
        EndIf
    EndIf

    If $iStartStep <= 3 Then
        $iFC_StepCurrent = 3
        _FC_AuditStep(3, "FileOpen")
        Local $bFileOK    = False
        Local $iTentative = 0
        While Not $bFileOK And $iTentative < 3
            $iTentative += 1
            Local $t3 = TimerInit()
            _FC_AuditLog("  Tentative " & $iTentative & "/3")
            _Spinner("FC [" & $Num & "] 3/7 - FileOpen (essai " & $iTentative & "/3)...")
            If $iTentative > 1 Then
                If WinExists($sFC_FILEOPEN) Then
                    WinClose($sFC_FILEOPEN)
                    WinWaitClose($sFC_FILEOPEN, "", 3)
                EndIf
                WinActivate($hWnd)
                WinWaitActive($hWnd, "", 3)
                _FC_SmartSleep(500)
                If $bFC_Stop Or $bFC_Skip Then ExitLoop
                ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
            EndIf
            Local $iTimer = TimerInit()
            While Not WinExists($sFC_FILEOPEN)
                _Tracker_PollButtons()
                If $bFC_Stop Or $bFC_Skip Then ExitLoop
                Sleep(100)
                If _IsPressed("1B") Then
                    _FC_AuditLog("*** ECHAP pendant attente FileOpen ***")
                    _FC_AuditShow($Num)
                    Return
                EndIf
                If TimerDiff($iTimer) > 10000 Then
                    _FC_AuditLog("  TIMEOUT 10s : FileOpen ne s'ouvre pas")
                    ExitLoop
                EndIf
            WEnd
            If $bFC_Stop Or $bFC_Skip Then ExitLoop
            _FC_AuditTiming("Attente FileOpen", TimerDiff($iTimer))
            If Not WinExists($sFC_FILEOPEN) Then
                _FC_AuditLog("  FileOpen absent apres timeout")
                ContinueLoop
            EndIf
            WinActivate($sFC_FILEOPEN)
            WinWaitActive($sFC_FILEOPEN, "", 3)
            _FC_AuditWinState($sFC_FILEOPEN, "FileOpen")
            _FC_SmartSleep(300)
            If $bFC_Stop Or $bFC_Skip Then ExitLoop
            ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", "")
            _FC_SmartSleep(150)
            ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", $sFC_EDS)
            _FC_SmartSleep(500)
            Local $sReadBack = ControlGetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]")
            _FC_AuditLog("  Champ apres ecriture = '" & $sReadBack & "'")
            If Not StringInStr($sReadBack, "EXPORT_HPE_FILECLOSING") Then
                _FC_AuditLog("  *** ECHEC ecriture champ ***")
                ContinueLoop
            EndIf
            Send("{ENTER}")
            Local $iWait = TimerInit()
            While WinExists($sFC_FILEOPEN)
                Sleep(100)
                If TimerDiff($iWait) > 5000 Then ExitLoop
            WEnd
            _FC_AuditTiming("Fermeture FileOpen", TimerDiff($iWait))
            If Not WinExists($sFC_FILEOPEN) Then
                $bFileOK = True
                _FC_AuditLog("  FileOpen OK")
            Else
                _FC_AuditLog("  *** FileOpen bloque ***")
            EndIf
            _FC_AuditTiming("Step3-Tentative" & $iTentative, TimerDiff($t3))
        WEnd
        If Not $bFileOK Then
            _FC_AuditLog("*** ECHEC FINAL FileOpen 3 tentatives ***")
            _FC_AuditTiming("TOTAL", TimerDiff($tTotal))
            _FC_AuditShow($Num)
            MsgBox(16+262144, "Erreur FC", "Impossible d'ouvrir le fichier EDS." & @CRLF & "Dossier : " & $Num)
            $bFC_Stop = True
            Return
        EndIf
        _FC_WaitIfPaused()
        If $bFC_Stop Then
            _FC_AuditLog("*** STOP Step 3 ***")
            _FC_AuditShow($Num)
            Return
        EndIf
    EndIf

    If $iStartStep <= 4 Then
        $iFC_StepCurrent = 4
        _FC_AuditStep(4, "Menu Selection")
        Local $t4 = TimerInit()
        _WinWaitSpinner($sFC_MENU, "FC [" & $Num & "] 4/7 - Menu Selection...")
        _FC_AuditTiming("Attente Menu", TimerDiff($t4))
        If $bFC_Stop Then
            _FC_AuditLog("*** STOP Step 4 ***")
            _FC_AuditShow($Num)
            Return
        EndIf
        _FC_AuditWinState($sFC_MENU, "Menu Selection")
        Local $hMenu = WinActivate($sFC_MENU)
        WinWaitActive($hMenu, "", 3)
        _FC_SmartSleep(300)
        ControlSetText($hMenu, "", "[CLASS:TEdit; INSTANCE:1]", "1")
        _FC_SmartSleep(300)
        ControlClick($hMenu, "", "[TEXT:OK]")
        WinWaitClose($hMenu, "", 5)
        _FC_SmartSleep(500)
        _FC_AuditTiming("Step4-Menu", TimerDiff($t4))
        If $bFC_Stop Or $bFC_Skip Then
            _FC_AuditLog("*** STOP/SKIP Step 4 apres ***")
            _FC_AuditShow($Num)
            Return
        EndIf
    EndIf

    If $iStartStep <= 5 Then
        $iFC_StepCurrent = 5
        _FC_AuditStep(5, "Numero J = " & $Num)
        Local $t5 = TimerInit()
        _WinWaitSpinner($sFC_INPUT, "FC [" & $Num & "] 5/7 - Numero J...")
        _FC_AuditTiming("Attente Input NumJ", TimerDiff($t5))
        If $bFC_Stop Then
            _FC_AuditLog("*** STOP Step 5 ***")
            _FC_AuditShow($Num)
            Return
        EndIf
        Local $hInput1 = WinActivate($sFC_INPUT)
        WinWaitActive($hInput1, "", 3)
        _FC_SmartSleep(300)
        ControlSetText($hInput1, "", "[CLASS:TEdit; INSTANCE:1]", $Num)
        _FC_SmartSleep(300)
        ControlClick($hInput1, "", "[TEXT:OK]")
        WinWaitClose($hInput1, "", 5)
        _FC_SmartSleep(300)
        _FC_AuditTiming("Step5-NumJ", TimerDiff($t5))
        If $bFC_Stop Or $bFC_Skip Then
            _FC_AuditLog("*** STOP/SKIP Step 5 apres ***")
            _FC_AuditShow($Num)
            Return
        EndIf
    EndIf

    If $iStartStep <= 6 Then
        $iFC_StepCurrent = 6
        _FC_AuditStep(6, "Carrier ID = " & $CarrierID)
        Local $t6 = TimerInit()
        _WinWaitSpinner($sFC_CARRIER, "FC [" & $Num & "] 6/7 - Carrier ID [" & $CarrierID & "]...")
        _FC_AuditTiming("Attente Carrier", TimerDiff($t6))
        If $bFC_Stop Then
            _FC_AuditLog("*** STOP Step 6 ***")
            _FC_AuditShow($Num)
            Return
        EndIf
        Local $hCarrier = WinActivate($sFC_CARRIER)
        WinWaitActive($hCarrier, "", 3)
        _FC_SmartSleep(300)
        ControlSetText($hCarrier, "", "[CLASS:TEdit; INSTANCE:1]", $CarrierID)
        _FC_SmartSleep(300)
        ControlClick($hCarrier, "", "[TEXT:OK]")
        WinWaitClose($hCarrier, "", 5)
        _FC_SmartSleep(3000)
        _FC_AuditTiming("Step6-Carrier", TimerDiff($t6))
        If $bFC_Stop Or $bFC_Skip Then
            _FC_AuditLog("*** STOP/SKIP Step 6 apres ***")
            _FC_AuditShow($Num)
            Return
        EndIf
    EndIf

    ; ── Calcul dates (si pas d'override depuis la modale) ────────────────────
    Local $bApresMidi = ((Number(@HOUR) * 60 + Number(@MIN)) >= (14 * 60 + 30))
    Local $sDateG
    If $DateGOverride <> "" Then
        $sDateG = $DateGOverride
    Else
        Local $iJoursG
        If $bApresMidi Then
            If $CarrierID = "7" Then
                $iJoursG = 2
            Else
                $iJoursG = 3
            EndIf
        Else
            If $CarrierID = "7" Then
                $iJoursG = 1
            Else
                $iJoursG = 2
            EndIf
        EndIf
        Local $dateG = _FC_WorkDay(@MDAY, @MON, @YEAR, $iJoursG)
        $sDateG = StringFormat("%02d.%02d.%02d", $dateG[0], $dateG[1], Mod($dateG[2], 100))
    EndIf
    Local $sDateL = "X"
    If $bApresMidi Then
        Local $dateL = _FC_WorkDay(@MDAY, @MON, @YEAR, 1)
        $sDateL = StringFormat("%02d.%02d.%02d", $dateL[0], $dateL[1], Mod($dateL[2], 100))
    EndIf
    Local $sTextH = $sDateG & " entre " & $Horaire

    Local $aVal[16]
    $aVal[0]  = "DEF"
    $aVal[1]  = "1800"
    $aVal[2]  = "X"
    $aVal[3]  = "1600"
    $aVal[4]  = $sDateG
    $aVal[5]  = $sTextH
    $aVal[6]  = ""
    $aVal[7]  = $Notes
    $aVal[8]  = "1800"
    $aVal[9]  = $sDateL
    $aVal[10] = "1600"
    $aVal[11] = $sDateG
    $aVal[12] = $DLY
    $aVal[13] = $DLYNotes
    $aVal[14] = "1600"
    $aVal[15] = $sDateG

    Local $aSkip[16]
    $aSkip[0]  = False
    $aSkip[1]  = False
    $aSkip[2]  = False
    $aSkip[3]  = False
    $aSkip[4]  = False
    $aSkip[5]  = False
    $aSkip[6]  = False
    $aSkip[7]  = False
    $aSkip[8]  = False
    $aSkip[9]  = False
    $aSkip[10] = False
    $aSkip[11] = False
    $aSkip[12] = False
    If $DLY <> "Y" Then
        $aSkip[13] = True
    Else
        $aSkip[13] = False
    EndIf
    $aSkip[14] = False
    $aSkip[15] = False

    If $iStartStep <= 7 Then
        $iFC_StepCurrent = 7
        _FC_AuditStep(7, "Colonnes C..R (16 popups)")
        Local $p
        For $p = 0 To 15
            If $aSkip[$p] Then
                _FC_AuditLog("  Col " & Chr(67 + $p) & " : SKIP")
                ContinueLoop
            EndIf
            If $bFC_Stop Then
                _FC_AuditLog("*** STOP pendant colonnes (p=" & $p & ") ***")
                _FC_AuditShow($Num)
                Return
            EndIf
            Local $sVal       = $aVal[$p]
            Local $colNom     = Chr(67 + $p)
            Local $bColValidee = False
            Local $iTimeout   = 0
            If $p >= 14 Then $iTimeout = 3
            Local $tCol = TimerInit()
            While Not $bColValidee
                _FC_WaitIfPaused()
                If $bFC_Stop Or $bFC_Skip Then
                    _FC_AuditLog("*** STOP/SKIP Col " & $colNom & " ***")
                    _FC_AuditShow($Num)
                    Return
                EndIf
                _Spinner("FC [" & $Num & "] Col " & $colNom & "...")
                Local $hWinWait = WinWait($sFC_INPUT, "", $iTimeout)
                If $hWinWait = 0 And $p >= 14 Then
                    _FC_AuditLog("  Col " & $colNom & " : timeout (optionnel, fin)")
                    ExitLoop 2
                EndIf
                Local $hWin   = WinActivate($sFC_INPUT)
                Local $sTitre = WinGetTitle($hWin)
                _FC_SmartSleep(150)
                If $bFC_Stop Or $bFC_Skip Then Return
                If StringInStr($sTitre, "REASON") Then
                    _FC_AuditLog("  Col " & $colNom & " : REASON popup -> DE")
                    ControlSetText($hWin, "", "[CLASS:TEdit; INSTANCE:1]", "DE")
                    _FC_SmartSleep(150)
                    ControlClick($hWin, "", "[TEXT:OK]")
                    WinWaitClose($hWin)
                    _FC_SmartSleep(300)
                Else
                    ControlSetText($hWin, "", "[CLASS:TEdit; INSTANCE:1]", $sVal)
                    _FC_SmartSleep(150)
                    ControlClick($hWin, "", "[TEXT:OK]")
                    WinWaitClose($hWin)
                    $bColValidee = True
                    _FC_SmartSleep(300)
                EndIf
            WEnd
            _FC_AuditTiming("Col " & $colNom & " (val='" & StringLeft($sVal, 30) & "')", TimerDiff($tCol))
        Next
        _FC_SmartSleep(500)
    EndIf

    _FC_AuditTiming("TOTAL FC-Single", TimerDiff($tTotal))
    _FC_AuditLog("====== FIN FC-Single OK ======")
    _FC_AuditSave($Num)
    $iFC_StepCurrent = 0
    ToolTip("")
EndFunc


; ==============================================================================
; TRACKER VISUEL — avec boutons Pause / Play / Passer / Stop
; ==============================================================================

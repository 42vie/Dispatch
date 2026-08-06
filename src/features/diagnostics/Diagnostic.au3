; ============================================================================
; Diagnostic.au3
; Diagnostics.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _RunDiagnostic()
    Local $sLog = ""
    Local $tGlobal = TimerInit()

    ; ── Helper interne pour écrire dans le log ──
    $sLog &= "╔══════════════════════════════════════════════════════════════╗" & @CRLF
    $sLog &= "║  DIAGNOSTIC DISPATCH — " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "  ║" & @CRLF
    $sLog &= "╚══════════════════════════════════════════════════════════════╝" & @CRLF & @CRLF

    ; ══════════════════════════════════════════════════════════════════════════
    ; 1. INFOS SYSTÈME
    ; ══════════════════════════════════════════════════════════════════════════
    $sLog &= "── 1. SYSTEME ──────────────────────────────────────────────────" & @CRLF
    $sLog &= "  PC         : " & @ComputerName & @CRLF
    $sLog &= "  User       : " & @UserName & @CRLF
    $sLog &= "  OS         : " & @OSVersion & " " & @OSArch & " (build " & @OSBuild & ")" & @CRLF
    $sLog &= "  CPU        : " & @CPUArch & @CRLF
    Local $aMem = MemGetStats()
    Local $iRamTotalMB = Round($aMem[1] / 1024, 0)
    Local $iRamFreeMB  = Round($aMem[2] / 1024, 0)
    Local $iRamUsedPct = $aMem[0]
    $sLog &= "  RAM        : " & $iRamFreeMB & " MB libre / " & $iRamTotalMB & " MB total (" & $iRamUsedPct & "% utilisé)"
    If $iRamUsedPct > 90 Then
        $sLog &= "  *** RAM CRITIQUE ***"
    ElseIf $iRamUsedPct > 80 Then
        $sLog &= "  ** RAM ELEVEE **"
    EndIf
    $sLog &= @CRLF
    $sLog &= "  Script Dir : " & @ScriptDir & @CRLF
    $sLog &= @CRLF

    ; ══════════════════════════════════════════════════════════════════════════
    ; 2. PROCESSUS
    ; ══════════════════════════════════════════════════════════════════════════
    $sLog &= "── 2. PROCESSUS ────────────────────────────────────────────────" & @CRLF
    Local $aProcs[5] = ["ETMS.exe", "Outlook.exe", "chrome.exe", "msedge.exe", "explorer.exe"]
    For $i = 0 To 4
        Local $aP = ProcessList($aProcs[$i])
        If IsArray($aP) And $aP[0][0] > 0 Then
            $sLog &= "  " & StringFormat("%-15s", $aProcs[$i]) & " : " & $aP[0][0] & " instance(s), PID=" & $aP[1][1] & @CRLF
        Else
            $sLog &= "  " & StringFormat("%-15s", $aProcs[$i]) & " : NON TROUVE" & @CRLF
        EndIf
    Next
    $sLog &= @CRLF

    ; ══════════════════════════════════════════════════════════════════════════
    ; 3. BENCHMARK DISQUE / FICHIERS
    ; ══════════════════════════════════════════════════════════════════════════
    $sLog &= "── 3. BENCHMARK DISQUE ─────────────────────────────────────────" & @CRLF

    ; 3a. Écriture temp
    Local $sTmp = @TempDir & "\dispatch_diag_test.tmp"
    Local $tDisk = TimerInit()
    Local $hTmp = FileOpen($sTmp, 2)
    If $hTmp <> -1 Then
        For $x = 1 To 100
            FileWrite($hTmp, "BENCHMARK LIGNE " & $x & @CRLF)
        Next
        FileClose($hTmp)
    EndIf
    Local $nDiskWrite = TimerDiff($tDisk)
    $sLog &= "  Ecriture 100 lignes (TEMP) : " & Round($nDiskWrite, 0) & " ms"
    If $nDiskWrite > 500 Then $sLog &= "  *** DISQUE LENT ***"
    $sLog &= @CRLF

    ; 3b. Lecture temp
    $tDisk = TimerInit()
    If FileExists($sTmp) Then
        Local $sContent = FileRead($sTmp)
    EndIf
    Local $nDiskRead = TimerDiff($tDisk)
    $sLog &= "  Lecture fichier TEMP       : " & Round($nDiskRead, 0) & " ms"
    If $nDiskRead > 200 Then $sLog &= "  *** LECTURE LENTE ***"
    $sLog &= @CRLF
    FileDelete($sTmp)

    ; 3c. Accès EDS
    Local $sEDS = "F:\Scripting\Export\EXPORT_HPE_FILECLOSING_001\EXPORT_HPE_FILECLOSING_031.eds"
    $tDisk = TimerInit()
    Local $bEdsExists = FileExists($sEDS)
    Local $nEdsCheck = TimerDiff($tDisk)
    $sLog &= "  Accès fichier EDS          : " & Round($nEdsCheck, 0) & " ms"
    If $nEdsCheck > 1000 Then $sLog &= "  *** ACCES RESEAU LENT ***"
    $sLog &= @CRLF
    If $bEdsExists Then
        Local $iEdsSize = FileGetSize($sEDS)
        $sLog &= "    -> existe : OUI | taille=" & $iEdsSize & " octets (" & Round($iEdsSize/1024, 1) & " KB)" & @CRLF
    Else
        $sLog &= "    -> *** FICHIER EDS INTROUVABLE ***" & @CRLF
        ; Lister le dossier
        Local $sEdsDir = "F:\Scripting\Export\EXPORT_HPE_FILECLOSING_001"
        If FileExists($sEdsDir) Then
            $sLog &= "    -> Dossier existe. Contenu :" & @CRLF
            Local $hSearch = FileFindFirstFile($sEdsDir & "\*.*")
            If $hSearch <> -1 Then
                Local $iCnt = 0
                While 1
                    Local $sF = FileFindNextFile($hSearch)
                    If @error Then ExitLoop
                    $sLog &= "       " & $sF & @CRLF
                    $iCnt += 1
                    If $iCnt > 20 Then
                        $sLog &= "       ... (limité à 20)" & @CRLF
                        ExitLoop
                    EndIf
                WEnd
                FileClose($hSearch)
            EndIf
        Else
            $sLog &= "    -> *** DOSSIER INTROUVABLE : " & $sEdsDir & " ***" & @CRLF
        EndIf
    EndIf

    ; 3d. Accès dossier PJ
    Local $sIniDiag = @ScriptDir & "\dispatch_config.ini"
    Local $sPJPath = IniRead($sIniDiag, "PJ", "Path", "")
    If $sPJPath <> "" Then
        $tDisk = TimerInit()
        Local $bPJExists = FileExists($sPJPath)
        Local $nPJ = TimerDiff($tDisk)
        $sLog &= "  Accès dossier PJ           : " & Round($nPJ, 0) & " ms (" & $sPJPath & ")"
        If $nPJ > 1000 Then $sLog &= "  *** RESEAU LENT ***"
        $sLog &= @CRLF
    EndIf

    ; 3e. Accès réseau partagé
    Local $sNetPath = IniRead($sIniDiag, "Network", "StatePath", "")
    If $sNetPath <> "" Then
        $tDisk = TimerInit()
        Local $bNetExists = FileExists($sNetPath)
        Local $nNet = TimerDiff($tDisk)
        $sLog &= "  Accès state réseau         : " & Round($nNet, 0) & " ms (" & $sNetPath & ")"
        If $nNet > 2000 Then $sLog &= "  *** RESEAU TRES LENT ***"
        $sLog &= @CRLF
    EndIf
    $sLog &= @CRLF

    ; ══════════════════════════════════════════════════════════════════════════
    ; 4. BENCHMARK E.TMS
    ; ══════════════════════════════════════════════════════════════════════════
    $sLog &= "── 4. E.TMS ───────────────────────────────────────────────────" & @CRLF

    Local $tETMS = TimerInit()
    Local $hETMS = WinGetHandle("[CLASS:TfmBrowser]")
    Local $nFindETMS = TimerDiff($tETMS)
    $sLog &= "  Trouver fenetre E.TMS      : " & Round($nFindETMS, 0) & " ms"
    If $nFindETMS > 500 Then $sLog &= "  *** LENT ***"
    $sLog &= @CRLF

    If $hETMS = 0 Or Not WinExists($hETMS) Then
        $sLog &= "  *** E.TMS N'EST PAS OUVERT — tests E.TMS ignorés ***" & @CRLF
    Else
        Local $sETMSTitle = WinGetTitle($hETMS)
        Local $aETMSPos   = WinGetPos($hETMS)
        Local $iETMSState = WinGetState($hETMS)
        $sLog &= "  Titre      : " & $sETMSTitle & @CRLF
        If IsArray($aETMSPos) Then
            $sLog &= "  Position   : " & $aETMSPos[0] & "," & $aETMSPos[1] & " taille=" & $aETMSPos[2] & "x" & $aETMSPos[3] & @CRLF
        EndIf
        $sLog &= "  State      : " & $iETMSState & @CRLF

        ; 4a. WinActivate speed
        $tETMS = TimerInit()
        WinActivate($hETMS)
        WinWaitActive($hETMS, "", 3)
        Local $nActivate = TimerDiff($tETMS)
        $sLog &= "  WinActivate                : " & Round($nActivate, 0) & " ms"
        If $nActivate > 1000 Then
            $sLog &= "  *** TRES LENT ***"
        ElseIf $nActivate > 500 Then
            $sLog &= "  ** LENT **"
        EndIf
        $sLog &= @CRLF

        ; 4b. Instance detection
        Local $sInst = _GetETMSInstance($hETMS)
        $sLog &= "  Instance détectée          : " & $sInst & " (titre=" & StringLeft($sETMSTitle, 40) & ")" & @CRLF

        ; 4c. ControlGetText speed (LOG)
        Local $sLogCtrl = "[CLASS:TEIEdit; INSTANCE:91]"
        $tETMS = TimerInit()
        Local $sLogText = ControlGetText($hETMS, "", $sLogCtrl)
        Local $nGetText = TimerDiff($tETMS)
        $sLog &= "  ControlGetText (LOG)       : " & Round($nGetText, 0) & " ms | texte='" & StringLeft($sLogText, 50) & "'"
        If $nGetText > 300 Then $sLog &= "  *** LENT ***"
        $sLog &= @CRLF

        ; 4d. ControlSetText speed (LOG — on va écrire puis remettre)
        Local $sSaveText = $sLogText
        $tETMS = TimerInit()
        ControlSetText($hETMS, "", $sLogCtrl, "DIAG_TEST")
        Local $nSetText = TimerDiff($tETMS)
        $sLog &= "  ControlSetText (LOG)       : " & Round($nSetText, 0) & " ms"
        If $nSetText > 300 Then $sLog &= "  *** LENT ***"
        $sLog &= @CRLF

        ; 4e. Relecture pour vérifier
        $tETMS = TimerInit()
        Local $sVerify = ControlGetText($hETMS, "", $sLogCtrl)
        Local $nVerify = TimerDiff($tETMS)
        Local $bMatch = (StringStripWS($sVerify, 3) = "DIAG_TEST")
        $sLog &= "  Vérification écriture      : " & Round($nVerify, 0) & " ms | match=" & $bMatch
        If Not $bMatch Then $sLog &= "  *** ECHEC ECRITURE : lu='" & StringLeft($sVerify, 30) & "' ***"
        $sLog &= @CRLF

        ; Remettre le texte original
        ControlSetText($hETMS, "", $sLogCtrl, $sSaveText)

        ; 4f. Toolbar detection
        Local $sToolbar = "[CLASS:TRzToolbar; INSTANCE:1]"
        $tETMS = TimerInit()
        Local $hTB = ControlGetHandle($hETMS, "", $sToolbar)
        Local $nToolbar = TimerDiff($tETMS)
        $sLog &= "  Toolbar handle             : " & Round($nToolbar, 0) & " ms | handle=" & $hTB
        If $hTB = 0 Or $hTB = "" Then $sLog &= "  *** TOOLBAR INTROUVABLE ***"
        $sLog &= @CRLF

        ; 4g. Tester toutes les instances TEIEdit connues
        $sLog &= "  Contrôles TEIEdit :" & @CRLF
        Local $aInst[5] = ["83", "91", "109", "207", "300"]
        Local $aInstName[5] = ["NOTES", "LOG", "REFS", "HIST", "DIMST"]
        For $k = 0 To 4
            Local $sC = "[CLASS:TEIEdit; INSTANCE:" & $aInst[$k] & "]"
            $tETMS = TimerInit()
            Local $hC = ControlGetHandle($hETMS, "", $sC)
            Local $nC = TimerDiff($tETMS)
            Local $sT = ""
            If $hC <> 0 And $hC <> "" Then $sT = StringLeft(ControlGetText($hETMS, "", $sC), 30)
            $sLog &= "    [" & $aInstName[$k] & " inst:" & $aInst[$k] & "] handle=" & $hC & " | " & Round($nC, 0) & " ms"
            If $sT <> "" Then $sLog &= " | texte='" & $sT & "'"
            If $nC > 200 Then $sLog &= "  ** LENT **"
            $sLog &= @CRLF
        Next

        ; 4h. Test Send (PgUp) speed
        $tETMS = TimerInit()
        WinActivate($hETMS)
        Send("{PGUP}")
        Local $nSend = TimerDiff($tETMS)
        $sLog &= "  Send(PgUp)                 : " & Round($nSend, 0) & " ms"
        If $nSend > 500 Then $sLog &= "  *** LENT ***"
        $sLog &= @CRLF
    EndIf

    ; 4i. EDOC
    Local $tEdoc = TimerInit()
    Local $hEdoc = WinGetHandle("[CLASS:TfmEdocViewerMainDlg]")
    Local $nEdoc = TimerDiff($tEdoc)
    If $hEdoc <> 0 And WinExists($hEdoc) Then
        $sLog &= "  EDOC                       : OUVERT (hwnd=" & $hEdoc & " | " & Round($nEdoc, 0) & " ms)" & @CRLF
    Else
        $sLog &= "  EDOC                       : non ouvert" & @CRLF
    EndIf
    $sLog &= @CRLF

    ; ══════════════════════════════════════════════════════════════════════════
    ; 5. BENCHMARK RESEAU (serveur local)
    ; ══════════════════════════════════════════════════════════════════════════
    $sLog &= "── 5. SERVEUR LOCAL ────────────────────────────────────────────" & @CRLF
    Local $tNet = TimerInit()
    Local $iSock = TCPConnect("127.0.0.1", $g_iPort)
    Local $nConnect = TimerDiff($tNet)
    If $iSock >= 0 Then
        $sLog &= "  Connexion 127.0.0.1:" & $g_iPort & "   : " & Round($nConnect, 0) & " ms | OK" & @CRLF
        TCPCloseSocket($iSock)
    Else
        $sLog &= "  Connexion 127.0.0.1:" & $g_iPort & "   : ECHEC (err=" & @error & ")" & @CRLF
    EndIf
    $sLog &= @CRLF

    ; ══════════════════════════════════════════════════════════════════════════
    ; 6. RÉSUMÉ / VERDICT
    ; ══════════════════════════════════════════════════════════════════════════
    Local $nTotal = TimerDiff($tGlobal)
    $sLog &= "── 6. RESUME ──────────────────────────────────────────────────" & @CRLF
    $sLog &= "  Durée totale diagnostic    : " & Round($nTotal, 0) & " ms" & @CRLF
    $sLog &= @CRLF

    ; Verdicts
    Local $iProblems = 0
    If $iRamUsedPct > 85 Then
        $sLog &= "  [!] RAM saturée à " & $iRamUsedPct & "% — fermer des applis" & @CRLF
        $iProblems += 1
    EndIf
    If $nDiskWrite > 500 Then
        $sLog &= "  [!] Disque lent en écriture — antivirus ? disque plein ?" & @CRLF
        $iProblems += 1
    EndIf
    If $bEdsExists = False Then
        $sLog &= "  [!] Fichier EDS introuvable — chemin incorrect ou lecteur déconnecté" & @CRLF
        $iProblems += 1
    ElseIf $nEdsCheck > 1000 Then
        $sLog &= "  [!] Accès EDS lent (" & Round($nEdsCheck, 0) & " ms) — réseau lent ou F: surchargé" & @CRLF
        $iProblems += 1
    EndIf
    If $hETMS = 0 Or Not WinExists($hETMS) Then
        $sLog &= "  [!] E.TMS pas ouvert — impossible de tester la réactivité" & @CRLF
        $iProblems += 1
    Else
        If $nActivate > 1000 Then
            $sLog &= "  [!] E.TMS met " & Round($nActivate, 0) & " ms pour s'activer — PC surchargé" & @CRLF
            $iProblems += 1
        EndIf
        If $nGetText > 300 Then
            $sLog &= "  [!] Lecture contrôles E.TMS lente — E.TMS surchargé" & @CRLF
            $iProblems += 1
        EndIf
        If $nSetText > 300 Then
            $sLog &= "  [!] Écriture contrôles E.TMS lente — E.TMS surchargé" & @CRLF
            $iProblems += 1
        EndIf
        If Not $bMatch Then
            $sLog &= "  [!] ControlSetText ne fonctionne pas — E.TMS bloqué ou protégé" & @CRLF
            $iProblems += 1
        EndIf
    EndIf
    If $sPJPath <> "" And $nPJ > 1000 Then
        $sLog &= "  [!] Dossier PJ lent (" & Round($nPJ, 0) & " ms) — réseau" & @CRLF
        $iProblems += 1
    EndIf
    If $sNetPath <> "" And $nNet > 2000 Then
        $sLog &= "  [!] State réseau lent (" & Round($nNet, 0) & " ms) — partage réseau lent" & @CRLF
        $iProblems += 1
    EndIf

    $sLog &= @CRLF
    If $iProblems = 0 Then
        $sLog &= "  ✓ TOUT EST OK — aucun problème détecté" & @CRLF
    Else
        $sLog &= "  ✗ " & $iProblems & " PROBLEME(S) DETECTE(S)" & @CRLF
    EndIf
    $sLog &= @CRLF & "══════════════════════════════════════════════════════════════════" & @CRLF

    ; ── Sauvegarder le log ──
    Local $sLogDir = @ScriptDir & "\logs"
    If Not FileExists($sLogDir) Then DirCreate($sLogDir)
    Local $sLogFile = $sLogDir & "\DIAG_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".log"
    Local $hLogFile = FileOpen($sLogFile, 2)
    If $hLogFile <> -1 Then
        FileWrite($hLogFile, $sLog)
        FileClose($hLogFile)
    EndIf

    ; ── Afficher la GUI ──
    Local $hDiag = GUICreate("DIAGNOSTIC DISPATCH", 800, 600, -1, -1)
    GUISetBkColor(0x1E1E1E, $hDiag)
    GUISetFont(9, 400, 0, "Consolas")
    Local $idDiagEdit = GUICtrlCreateEdit($sLog, 5, 5, 790, 545, BitOR(0x0004, 0x0800, 0x00200000))
    GUICtrlSetBkColor($idDiagEdit, 0x1E1E1E)
    GUICtrlSetColor($idDiagEdit, 0x00FF00)
    Local $idDiagCopy  = GUICtrlCreateButton("Copier", 5, 555, 130, 35)
    Local $idDiagOpen  = GUICtrlCreateButton("Ouvrir le .log", 140, 555, 160, 35)
    Local $idDiagRerun = GUICtrlCreateButton("Relancer", 305, 555, 130, 35)
    Local $idDiagClose = GUICtrlCreateButton("Fermer", 665, 555, 130, 35)
    GUISetState(@SW_SHOW, $hDiag)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idDiagClose
                ExitLoop
            Case $idDiagCopy
                ClipPut($sLog)
                ToolTip("Copié dans le presse-papier !", Default, Default, "Diagnostic", 1)
                Sleep(1000)
                ToolTip("")
            Case $idDiagOpen
                ShellExecute($sLogFile)
            Case $idDiagRerun
                GUIDelete($hDiag)
                _RunDiagnostic()
                Return
        EndSwitch
    WEnd
    GUIDelete($hDiag)
EndFunc

; ==============================================================================
; CLEANER CONTACTS — Nettoie les apostrophes et caractères spéciaux encodés
; ==============================================================================

Func _GetStorageInfo()
    Local $sResult = '{"files":['

    ; Liste des fichiers à vérifier
    Local $aCheck[7][2] = [ _
        [$g_sSaveFile, "historique_dispatch.json"], _
        [$g_sStatusFile, "dispatch_status.json"], _
        [$g_sDataFile, "dispatch_data.json"], _
        [$g_sContactsFile, "dispatch_contacts.json"], _
        [@ScriptDir & "\dispatch_contacts_meta.json", "dispatch_contacts_meta.json"], _
        [@ScriptDir & "\dispatch_config.ini", "dispatch_config.ini"], _
        [@ScriptDir & "\dispatch_contacts_0.json", "dispatch_contacts_0.json"] _
    ]

    Local $iTotalSize = 0
    For $i = 0 To 6
        Local $sFile = $aCheck[$i][0]
        Local $sName = $aCheck[$i][1]
        Local $iSize = 0
        If FileExists($sFile) Then $iSize = FileGetSize($sFile)
        $iTotalSize += $iSize
        If $i > 0 Then $sResult &= ","
        $sResult &= '{"name":"' & $sName & '","size":' & $iSize & '}'
    Next

    ; Chercher les chunks contacts additionnels
    For $c = 1 To 20
        Local $sChk = @ScriptDir & "\dispatch_contacts_" & $c & ".json"
        If Not FileExists($sChk) Then ExitLoop
        Local $iChkSize = FileGetSize($sChk)
        $iTotalSize += $iChkSize
        $sResult &= ',{"name":"dispatch_contacts_' & $c & '.json","size":' & $iChkSize & '}'
    Next

    $sResult &= '],"totalBytes":' & $iTotalSize
    $sResult &= ',"totalKB":' & Round($iTotalSize / 1024, 1)
    $sResult &= ',"totalMB":' & Round($iTotalSize / 1048576, 2)
    $sResult &= '}'
    Return $sResult
EndFunc

; ==============================================================================
; NOTIFICATION ERREUR — TrayTip visible pour tout problème
; ==============================================================================

Func _SilentHealthCheck()
    Local $iErrors = 0

    ; 1. Vérifier que les fichiers JSON sont valides et accessibles
    Local $aCheck[4] = [$g_sStatusFile, $g_sDataFile, $g_sContactsFile, $g_sSaveFile]
    Local $aNames[4] = ["status", "data", "contacts", "historique"]
    For $i = 0 To 3
        If FileExists($aCheck[$i]) Then
            Local $iSize = FileGetSize($aCheck[$i])
            If $iSize = 0 Then
                _AuditLog("WARN", "Fichier vide détecté : " & $aNames[$i] & " (" & $aCheck[$i] & ")")
                $iErrors += 1
            EndIf
            ; Vérifier que le fichier n'est pas corrompu (doit commencer par [ ou {)
            Local $hCheck = FileOpen($aCheck[$i], 256)
            If $hCheck <> -1 Then
                Local $sFirst = StringStripWS(StringLeft(FileRead($hCheck), 5), 3)
                FileClose($hCheck)
                If $sFirst <> "" And StringLeft($sFirst, 1) <> "[" And StringLeft($sFirst, 1) <> "{" Then
                    _AuditLog("ERREUR", "JSON corrompu : " & $aNames[$i] & " commence par '" & StringLeft($sFirst, 3) & "' au lieu de [ ou {")
                    _NotifyError("JSON", "Fichier " & $aNames[$i] & " possiblement corrompu !")
                    $iErrors += 1
                EndIf
            EndIf
        EndIf
    Next

    ; 2. Vérifier que le fichier réseau partagé est accessible
    Local $sIniNet = @ScriptDir & "\dispatch_config.ini"
    Local $sNetPath = IniRead($sIniNet, "Network", "StatePath", "")
    If $sNetPath <> "" Then
        If Not FileExists($sNetPath) Then
            ; Ne pas alerter si le chemin n'a jamais existé (premier lancement)
            Local $sDir = StringLeft($sNetPath, StringInStr($sNetPath, "\", 0, -1))
            If FileExists($sDir) Then
                _AuditLog("WARN", "Fichier réseau introuvable : " & $sNetPath)
            EndIf
        EndIf
    EndIf

    ; 3. Vérifier la taille mémoire des fichiers (alerte si > 5 MB)
    Local $iTotalSize = 0
    For $i = 0 To 3
        If FileExists($aCheck[$i]) Then $iTotalSize += FileGetSize($aCheck[$i])
    Next
    If $iTotalSize > 5242880 Then ; > 5 MB
        _AuditLog("WARN", "Stockage total élevé : " & Round($iTotalSize / 1048576, 2) & " MB — pensez à nettoyer les contacts ou vider les terminés")
        _NotifyError("Stockage", "Les fichiers JSON font " & Round($iTotalSize / 1048576, 1) & " MB — nettoyage recommandé")
    EndIf

    ; 4. Vérifier l'ancienneté du fichier data (si > 24h sans modification = potentiel problème)
    If FileExists($g_sDataFile) Then
        Local $sModTime = FileGetTime($g_sDataFile, 0, 1)
        If $sModTime <> "" Then
            Local $sDiff = _DateDiff("h", StringMid($sModTime, 1, 4) & "/" & StringMid($sModTime, 5, 2) & "/" & StringMid($sModTime, 7, 2) & " " & StringMid($sModTime, 9, 2) & ":" & StringMid($sModTime, 11, 2) & ":" & StringMid($sModTime, 13, 2), _NowCalc())
            If Number($sDiff) > 24 Then
                _AuditLog("INFO", "Fichier data non modifié depuis " & $sDiff & "h")
            EndIf
        EndIf
    EndIf

    If $iErrors > 0 Then
        _AuditLog("WARN", "Health check : " & $iErrors & " problème(s) détecté(s)")
    EndIf
EndFunc


; ══════════════════════════════════════════════════════════════════════════
; FONCTIONS dispatch_patch — Endpoints v2.1
; ══════════════════════════════════════════════════════════════════════════

; ── Validation et sanitisation des chemins réseau ──

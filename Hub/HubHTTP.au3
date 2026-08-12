; ============================================================================
; HubHTTP.au3
; ─────────────────────────────────────────────────────────────────────────────
; Pont HTTP universel et réutilisable — indépendant de Dispatch.
;
; PRINCIPE :
; Ce fichier est UN SEUL .au3 autonome (ou compilé en .exe) qui sert de
; pont HTTP générique vers les fonctionnalités communes :
; • Scan / ouverture mails Outlook
; • Logs & Audit
; • Historique (fichiers .json)
; • EDOC (init / scan / upload)
; • État / ping
;
; N'importe quel outil HTML appelle simplement :
; fetch("http://127.0.0.1:9510/api/hub", { method:"POST",
;   body: JSON.stringify({ action:"MAIL_SCAN", mailbox:"..." }) })
;
; ► Port par défaut : 9510 (≠ 9500 Dispatch, pas de conflit).
; ► Si le port 9510 est pris, un port libre est trouvé automatiquement,
;   puis écrit dans hub_port.txt (dossier du script) pour que votre HTML
;   puisse le relire au démarrage.
; ► CORS * activé → tous vos HTML locaux peuvent l'appeler sans restrictions.
;
; INTÉGRATION DANS UN HTML :
;
; fetch("http://127.0.0.1:9510/api/hub", {
;   method: "POST",
;   headers: { "Content-Type": "application/json" },
;   body: JSON.stringify({ action: "EDOC_INIT" })
; }).then(r => r.json()).then(console.log);
;
#include <Array.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>

; ── Port & socket globaux ──────────────────────────────────────────────────
Global Const $HUB_DEFAULT_PORT = 9510
Global $g_iHubPort = $HUB_DEFAULT_PORT
Global $g_iHubSocket = -1

; ── Outlook (créé à la demande, réutilisé entre les appels) ───────────────
Global $g_oHubOutlook = 0
Global $g_oHubNamespace = 0

; ── Fichiers ──────────────────────────────────────────────────────────────
Global $g_sHubLogFile = @ScriptDir & "\logs\hub_audit.log"
Global $g_sHubPortFile = @ScriptDir & "\hub_port.txt"

; ── EDOC intégré (plus de dépendance externe) ─────────────────────────────
Global $g_bEdocAvailable = True
Global $g_sEdocSessionId = ""
Global $g_aEdocLastScan[1]

; ══════════════════════════════════════════════════════════════════════════
; POINT D'ENTRÉE
; ══════════════════════════════════════════════════════════════════════════
TCPStartup()
_Hub_Start()

Func _Hub_Start()
    ; ── Mini GUI système ────────────────────────────────────────────────
    Local $hGUI = GUICreate("HubHTTP", 360, 90)
    GUISetBkColor(0x1A1A2E, $hGUI)
    Local $lblTitle = GUICtrlCreateLabel("■ HubHTTP — Pont universel", 14, 12, 330, 18)
    GUICtrlSetColor($lblTitle, 0x00E5FF)
    GUICtrlSetBkColor($lblTitle, 0x1A1A2E)
    Local $lblPort = GUICtrlCreateLabel("Démarrage...", 14, 36, 330, 18)
    GUICtrlSetColor($lblPort, 0x80FF80)
    GUICtrlSetBkColor($lblPort, 0x1A1A2E)
    Local $lblEdoc = GUICtrlCreateLabel("", 14, 58, 330, 18)
    GUICtrlSetColor($lblEdoc, 0xAAAAAA)
    GUICtrlSetBkColor($lblEdoc, 0x1A1A2E)
    GUISetState(@SW_SHOW, $hGUI)

    ; ── Trouver un port libre ────────────────────────────────────────────
    $g_iHubSocket = TCPListen("127.0.0.1", $HUB_DEFAULT_PORT)
    If $g_iHubSocket <> -1 Then
        $g_iHubPort = $HUB_DEFAULT_PORT
    Else
        Local $bFound = False
        For $i = 1 To 50
            Local $iTry = Random(10000, 60000, 1)
            $g_iHubSocket = TCPListen("127.0.0.1", $iTry)
            If $g_iHubSocket <> -1 Then
                $g_iHubPort = $iTry
                $bFound = True
                ExitLoop
            EndIf
        Next
        If Not $bFound Then
            MsgBox(16, "HubHTTP", "Impossible de démarrer le serveur TCP." & @CRLF & _
                "Vérifiez votre pare-feu ou redémarrez le PC.")
            Exit
        EndIf
    EndIf

    ; ── Créer les dossiers nécessaires ──────────────────────────────────
    If Not FileExists(@ScriptDir & "\logs") Then DirCreate(@ScriptDir & "\logs")
    If Not FileExists(@ScriptDir & "\Config") Then DirCreate(@ScriptDir & "\Config")
    If Not FileExists(@ScriptDir & "\data") Then DirCreate(@ScriptDir & "\data")

    ; ── Écrire le port dans hub_port.txt (pour les HTML qui le lisent) ──
    Local $hPortFile = FileOpen($g_sHubPortFile, 2)
    FileWrite($hPortFile, $g_iHubPort)
    FileClose($hPortFile)

    _Hub_Log("INFO", "HubHTTP démarré sur le port " & $g_iHubPort)

    GUICtrlSetData($lblPort, "http://127.0.0.1:" & $g_iHubPort & "/api/hub")
    GUICtrlSetData($lblEdoc, $g_bEdocAvailable ? _
        "EDOC : disponible (intégré)" : "EDOC : non disponible")

    ; ── Boucle principale ───────────────────────────────────────────────
    While 1
        Local $iMsg = GUIGetMsg()
        If $iMsg = -3 Then ExitLoop

        Local $iClient = TCPAccept($g_iHubSocket)
        If $iClient <> -1 Then
            _Hub_HandleClient($iClient)
        EndIf
        Sleep(10)
    WEnd

    TCPCloseSocket($g_iHubSocket)
    TCPShutdown()
    FileDelete($g_sHubPortFile)
    Exit
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; SERVEUR HTTP — réception & parsing
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_HandleClient($iSocket)
    Local $sRaw = ""
    Local $hTimer = TimerInit()
    While TimerDiff($hTimer) < 3000
        Local $sChunk = TCPRecv($iSocket, 8192)
        If @error Then ExitLoop
        If $sChunk <> "" Then
            $sRaw &= $sChunk
            If StringInStr($sRaw, @CRLF & @CRLF) > 0 Then ExitLoop
        EndIf
        Sleep(5)
    WEnd

    If $sRaw = "" Then
        TCPCloseSocket($iSocket)
        Return
    EndIf

    ; ── Séparer header / body ──────────────────────────────────────────
    Local $iSep = StringInStr($sRaw, @CRLF & @CRLF)
    Local $sHeader = StringLeft($sRaw, $iSep - 1)
    Local $sBody = StringTrimLeft($sRaw, $iSep + 3)

    ; ── Lire Content-Length et compléter le body si nécessaire ─────────
    Local $iCL = 0
    Local $aCL = StringRegExp($sHeader, "(?i)Content-Length:\s*(\d+)", 1)
    If Not @error And IsArray($aCL) Then $iCL = Int($aCL[0])
    $hTimer = TimerInit()
    While StringLen($sBody) < $iCL And TimerDiff($hTimer) < 3000
        Local $sMore = TCPRecv($iSocket, 8192)
        If @error Then ExitLoop
        If $sMore <> "" Then
            $sBody &= $sMore
            $hTimer = TimerInit()
        EndIf
        Sleep(5)
    WEnd

    ; ── Lire méthode & URL ─────────────────────────────────────────────
    Local $aLines = StringSplit($sHeader, @CRLF, 1)
    If $aLines[0] < 1 Then
        TCPCloseSocket($iSocket)
        Return
    EndIf
    Local $aTop = StringSplit($aLines[1], " ")
    If $aTop[0] < 2 Then
        TCPCloseSocket($iSocket)
        Return
    EndIf
    Local $sMethod = $aTop[1]
    Local $sURL = $aTop[2]

    ; ── CORS preflight ─────────────────────────────────────────────────
    If $sMethod = "OPTIONS" Then
        _Hub_SendRaw($iSocket, _
            "HTTP/1.1 204 No Content" & @CRLF & _
            "Access-Control-Allow-Origin: *" & @CRLF & _
            "Access-Control-Allow-Methods: GET, POST, OPTIONS" & @CRLF & _
            "Access-Control-Allow-Headers: Content-Type" & @CRLF & _
            "Access-Control-Max-Age: 86400" & @CRLF & _
            "Content-Length: 0" & @CRLF & _
            "Connection: close" & @CRLF & @CRLF)
        TCPCloseSocket($iSocket)
        Return
    EndIf

    ; ── Routage ────────────────────────────────────────────────────────
    If $sURL = "/api/hub" Then
        _Hub_Route($iSocket, $sBody)
    ElseIf $sURL = "/api/port" Or $sURL = "/api/ping" Then
        _Hub_Send($iSocket, 200, '{"status":"ok","port":' & $g_iHubPort & '}')
    ElseIf $sURL = "/api/actions" Then
        _Hub_Send($iSocket, 200, _Hub_ActionListJSON())
    Else
        _Hub_Send($iSocket, 404, '{"status":"error","message":"route_not_found","url":"' & _
            _Hub_JsonEsc($sURL) & '"}')
    EndIf

    TCPCloseSocket($iSocket)
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; ROUTEUR D'ACTIONS
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_Route($iSocket, $sBody)
    Local $sAction = _Hub_JsonGet($sBody, "action")
    _Hub_Log("ACTION", $sAction & " | " & StringLeft($sBody, 120))

    Switch $sAction

        ; ── Ping ───────────────────────────────────────────────────────
        Case "PING"
            _Hub_Send($iSocket, 200, '{"status":"ok","pong":true,"port":' & _
                $g_iHubPort & ',"hub":"HubHTTP"}')

        ; ── Mails ──────────────────────────────────────────────────────
        Case "MAIL_SCAN"
            _Hub_Send($iSocket, 200, _Hub_MailScan(_Hub_JsonGet($sBody, "mailbox")))

        Case "MAIL_OPEN"
            _Hub_Send($iSocket, 200, _Hub_MailOpen(_Hub_JsonGet($sBody, "entryId")))

        Case "MAIL_COMPOSE"
            _Hub_Send($iSocket, 200, _Hub_MailCompose( _
                _Hub_JsonGet($sBody, "to"), _
                _Hub_JsonGet($sBody, "subject"), _
                _Hub_JsonGet($sBody, "body")))

        ; ── Logs ───────────────────────────────────────────────────────
        Case "LOG_LIST"
            Local $iLines = Int(_Hub_JsonGet($sBody, "lines"))
            If $iLines <= 0 Then $iLines = 100
            _Hub_Send($iSocket, 200, _Hub_LogList($iLines))

        Case "LOG_CLEAR"
            FileDelete($g_sHubLogFile)
            _Hub_Send($iSocket, 200, '{"status":"ok"}')

        ; ── Historique (fichiers JSON génériques) ──────────────────────
        Case "HIST_LOAD"
            _Hub_Send($iSocket, 200, _Hub_HistLoad(_Hub_JsonGet($sBody, "file")))

        Case "HIST_SAVE"
            Local $sHSFile = _Hub_JsonGet($sBody, "file")
            Local $iHSPos = StringInStr($sBody, '"data"')
            Local $sHSData = ""
            If $iHSPos > 0 Then
                Local $iHSColon = StringInStr($sBody, ":", 0, 1, $iHSPos)
                If $iHSColon > 0 Then
                    $sHSData = StringStripWS(StringMid($sBody, $iHSColon + 1), 3)
                    Local $sHSTrimmed = StringStripWS($sHSData, 2)
                    If StringRight($sHSTrimmed, 1) = "}" Then
                        $sHSData = StringTrimRight($sHSTrimmed, 1)
                        $sHSData = StringStripWS($sHSData, 2)
                    EndIf
                EndIf
            EndIf
            _Hub_Send($iSocket, 200, _Hub_HistSave($sHSFile, $sHSData))

        ; ── EDOC (intégré) ─────────────────────────────────────────────
        Case "EDOC_INIT"
            _Hub_Send($iSocket, 200, _Hub_EdocInit())

        Case "EDOC_SCAN"
            _Hub_Send($iSocket, 200, _Hub_EdocScan($sBody))

        Case "EDOC_UPLOAD"
            _Hub_Send($iSocket, 200, _Hub_EdocUpload($sBody))

        ; ── Infos système ──────────────────────────────────────────────
        Case "STORAGE_INFO"
            _Hub_Send($iSocket, 200, _Hub_StorageInfo())

        Case "DISK_FREE"
            Local $sDrive = _Hub_JsonGet($sBody, "drive")
            If $sDrive = "" Then $sDrive = @ScriptDrive
            _Hub_Send($iSocket, 200, '{"status":"ok","drive":"' & _
                _Hub_JsonEsc($sDrive) & '","freeMB":' & _
                Round(DriveSpaceFree($sDrive)) & '}')

        ; ── Fichiers génériques data\ ──────────────────────────────────
        Case "FILE_READ"
            Local $sFRFile = _Hub_JsonGet($sBody, "file")
            If $sFRFile = "" Or StringInStr($sFRFile, "..") Then
                _Hub_Send($iSocket, 400, '{"status":"error","message":"params_invalides"}')
            Else
                Local $sFRPath = @ScriptDir & "\data\" & $sFRFile
                If FileExists($sFRPath) Then
                    _Hub_Send($iSocket, 200, '{"status":"ok","content":' & FileRead($sFRPath) & "}")
                Else
                    _Hub_Send($iSocket, 404, '{"status":"error","message":"file_not_found"}')
                EndIf
            EndIf

        Case "FILE_WRITE"
            Local $sFWFile = _Hub_JsonGet($sBody, "file")
            Local $iFWPos = StringInStr($sBody, '"content"')
            If $sFWFile = "" Or StringInStr($sFWFile, "..") Or $iFWPos = 0 Then
                _Hub_Send($iSocket, 400, '{"status":"error","message":"params_invalides"}')
            Else
                If Not FileExists(@ScriptDir & "\data") Then DirCreate(@ScriptDir & "\data")
                Local $iFWColon = StringInStr($sBody, ":", 0, 1, $iFWPos)
                Local $sFWRaw = StringStripWS(StringMid($sBody, $iFWColon + 1), 3)
                If StringRight(StringStripWS($sFWRaw, 2), 1) = "}" Then
                    $sFWRaw = StringTrimRight(StringStripWS($sFWRaw, 2), 1)
                EndIf
                Local $hFW = FileOpen(@ScriptDir & "\data\" & $sFWFile, 2 + 128)
                FileWrite($hFW, $sFWRaw)
                FileClose($hFW)
                _Hub_Send($iSocket, 200, '{"status":"ok"}')
            EndIf

        ; ── Action inconnue ────────────────────────────────────────────
        Case Else
            _Hub_Log("WARN", "Action inconnue : " & $sAction)
            _Hub_Send($iSocket, 400, '{"status":"error","message":"action_inconnue","action":"' & _
                _Hub_JsonEsc($sAction) & '"}')

    EndSwitch
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; EDOC — implémentation intégrée (mock)
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_EdocInit()
    $g_sEdocSessionId = "EDOC-" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC
    ReDim $g_aEdocLastScan[1]
    $g_aEdocLastScan[0] = 0

    Local $sRoot = @ScriptDir & "\data\edoc"
    If Not FileExists($sRoot) Then DirCreate($sRoot)

    Return '{"status":"ok","sessionId":"' & _Hub_JsonEsc($g_sEdocSessionId) & _
        '","root":"' & _Hub_JsonEsc($sRoot) & '"}'
EndFunc

Func _Hub_EdocScan($sBody)
    If $g_sEdocSessionId = "" Then
        Return '{"status":"error","message":"no_session"}'
    EndIf

    ; Ex: on peut lire un pattern ou un chemin dans le body si besoin
    Local $sFilter = _Hub_JsonGet($sBody, "filter")
    If $sFilter = "" Then $sFilter = "*.pdf"

    Local $sFolder = @ScriptDir & "\data\edoc_in"
    If Not FileExists($sFolder) Then DirCreate($sFolder)

    ; Mock : lister les fichiers du dossier d'entrée
    Local $aList = _FileListToArray($sFolder, $sFilter, 1)
    Local $sJson = "["
    If Not @error And IsArray($aList) Then
        For $i = 1 To $aList[0]
            If $sJson <> "[" Then $sJson &= ","
            $sJson &= '{"name":"' & _Hub_JsonEsc($aList[$i]) & '",' & _
                '"path":"' & _Hub_JsonEsc($sFolder & "\" & $aList[$i]) & '",' & _
                '"size":' & FileGetSize($sFolder & "\" & $aList[$i]) & "}"
        Next
    EndIf
    $sJson &= "]"

    Return '{"status":"ok","sessionId":"' & _Hub_JsonEsc($g_sEdocSessionId) & _
        '","filter":"' & _Hub_JsonEsc($sFilter) & '","documents":' & $sJson & "}"
EndFunc

Func _Hub_EdocUpload($sBody)
    If $g_sEdocSessionId = "" Then
        Return '{"status":"error","message":"no_session"}'
    EndIf

    Local $sName = _Hub_JsonGet($sBody, "name")
    Local $sContent = _Hub_JsonGet($sBody, "content") ; supposé être du texte ou base64

    If $sName = "" Or $sContent = "" Then
        Return '{"status":"error","message":"params_vides"}'
    EndIf

    Local $sDestDir = @ScriptDir & "\data\edoc_out"
    If Not FileExists($sDestDir) Then DirCreate($sDestDir)

    Local $sPath = $sDestDir & "\" & $sName
    Local $hFile = FileOpen($sPath, 2 + 128)
    If $hFile = -1 Then
        Return '{"status":"error","message":"impossible_ecrire"}'
    EndIf

    ; Ici on écrit le contenu brut tel quel.
    ; Si tu utilises du base64 côté front, tu pourras le décoder ici.
    FileWrite($hFile, $sContent)
    FileClose($hFile)

    Return '{"status":"ok","sessionId":"' & _Hub_JsonEsc($g_sEdocSessionId) & _
        '","saved":"' & _Hub_JsonEsc($sPath) & '"}'
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; IMPLÉMENTATIONS EXISTANTES : OUTLOOK, LOGS, HISTORIQUE, STORAGE, ETC.
; (copiées de ton fichier original, non modifiées, sauf pour edocAvailable)
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_EnsureOutlook()
    If IsObj($g_oHubOutlook) And IsObj($g_oHubNamespace) Then Return True
    $g_oHubOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($g_oHubOutlook) Then Return False
    $g_oHubNamespace = $g_oHubOutlook.GetNamespace("MAPI")
    If Not IsObj($g_oHubNamespace) Then Return False
    Return True
EndFunc

Func _Hub_ResolveMailboxFolders($sMailbox)
    Local $aOut[2] = [0, 0]
    If $sMailbox = "" Then
        $aOut[0] = $g_oHubNamespace.GetDefaultFolder(6)
        $aOut[1] = $g_oHubNamespace.GetDefaultFolder(5)
        Return $aOut
    EndIf
    Local $sWanted = StringUpper(StringStripWS($sMailbox, 3))
    Local $sShort = StringRegExpReplace($sWanted, "@.*$", "")
    For $oStore In $g_oHubNamespace.Stores
        Local $sDisp = StringUpper($oStore.DisplayName)
        If $sDisp = $sWanted Or $sDisp = $sShort Or StringInStr($sDisp, $sShort) Then
            $aOut[0] = $oStore.GetDefaultFolder(6)
            $aOut[1] = $oStore.GetDefaultFolder(5)
            ExitLoop
        EndIf
    Next
    Return $aOut
EndFunc

Func _Hub_MailScan($sMailbox)
    If Not _Hub_EnsureOutlook() Then
        Return '{"status":"error","message":"outlook_indisponible","mails":[]}'
    EndIf
    Local $aFolders = _Hub_ResolveMailboxFolders($sMailbox)
    Local $oInbox = $aFolders[0]
    Local $oSent = $aFolders[1]
    If $sMailbox <> "" And Not IsObj($oInbox) And Not IsObj($oSent) Then
        Return '{"status":"error","message":"mailbox_not_found","mails":[]}'
    EndIf

    Local $sDateLimit = _Hub_FmtDate(_DateAdd('d', -14, _NowCalc()))
    Local $aRows[500][6]
    Local $iRows = 0

    If IsObj($oInbox) Then
        Local $oItems = $oInbox.Items
        $oItems.Sort("[ReceivedTime]", True)
        Local $iScan = 0
        For $oItem In $oItems
            If $iScan >= 300 Or $iRows >= 490 Then ExitLoop
            If $oItem.Class <> 43 Then ContinueLoop
            Local $sDate = _Hub_FmtDate($oItem.ReceivedTime)
            If $sDate < $sDateLimit Then ExitLoop
            $iScan += 1
            $aRows[$iRows][0] = $sDate
            $aRows[$iRows][1] = $oItem.Subject
            $aRows[$iRows][2] = $oItem.SenderName
            $aRows[$iRows][3] = $oItem.EntryID
            $aRows[$iRows][4] = "false"
            $aRows[$iRows][5] = StringRegExp($oItem.Subject, "(?i)^(RE|TR|FW|FWD)\s*:") ? "true" : "false"
            $iRows += 1
        Next
    EndIf

    If IsObj($oSent) Then
        Local $oItemsS = $oSent.Items
        $oItemsS.Sort("[SentOn]", True)
        Local $iScanS = 0
        For $oItemS In $oItemsS
            If $iScanS >= 200 Or $iRows >= 499 Then ExitLoop
            If $oItemS.Class <> 43 Then ContinueLoop
            Local $sDateS = _Hub_FmtDate($oItemS.SentOn)
            If $sDateS < $sDateLimit Then ExitLoop
            $iScanS += 1
            $aRows[$iRows][0] = $sDateS
            $aRows[$iRows][1] = $oItemS.Subject
            $aRows[$iRows][2] = Chr(8594) & " " & $oItemS.To
            $aRows[$iRows][3] = $oItemS.EntryID
            $aRows[$iRows][4] = "true"
            $aRows[$iRows][5] = StringRegExp($oItemS.Subject, "(?i)^(RE|TR|FW|FWD)\s*:") ? "true" : "false"
            $iRows += 1
        Next
    EndIf

    If $iRows > 1 Then _ArraySort($aRows, 1, 0, $iRows - 1, 0, 0)

    Local $sJson = "["
    For $i = 0 To $iRows - 1
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '{"date":"' & _Hub_JsonEsc($aRows[$i][0]) & '"' & _
            ',"subject":"' & _Hub_JsonEsc($aRows[$i][1]) & '"' & _
            ',"sender":"' & _Hub_JsonEsc($aRows[$i][2]) & '"' & _
            ',"entryId":"' & _Hub_JsonEsc($aRows[$i][3]) & '"' & _
            ',"sent":' & $aRows[$i][4] & _
            ',"reply":' & $aRows[$i][5] & "}"
    Next
    $sJson &= "]"
    Return '{"status":"ok","count":' & $iRows & ',"mails":' & $sJson & "}"
EndFunc

Func _Hub_MailOpen($sEntryId)
    If $sEntryId = "" Then Return '{"status":"error","message":"entryId_vide"}'
    If Not _Hub_EnsureOutlook() Then Return '{"status":"error","message":"outlook_indisponible"}'
    Local $oMail = $g_oHubNamespace.GetItemFromID($sEntryId)
    If Not IsObj($oMail) Then Return '{"status":"error","message":"mail_introuvable"}'
    $oMail.Display()
    Return '{"status":"ok"}'
EndFunc

Func _Hub_MailCompose($sTo, $sSubject, $sBodyMail)
    If Not _Hub_EnsureOutlook() Then Return '{"status":"error","message":"outlook_indisponible"}'
    Local $oMail = $g_oHubOutlook.CreateItem(0)
    If $sTo <> "" Then $oMail.To = $sTo
    If $sSubject <> "" Then $oMail.Subject = $sSubject
    If $sBodyMail <> "" Then $oMail.Body = $sBodyMail
    $oMail.Display()
    Return '{"status":"ok"}'
EndFunc

Func _Hub_LogList($iMaxLines)
    If Not FileExists($g_sHubLogFile) Then Return '{"status":"ok","lines":[]}'
    Local $sAll = FileRead($g_sHubLogFile)
    Local $aAll = StringSplit($sAll, @LF, 2)
    Local $iStart = UBound($aAll) - $iMaxLines
    If $iStart < 0 Then $iStart = 0
    Local $sJson = "["
    For $i = $iStart To UBound($aAll) - 1
        If StringStripWS($aAll[$i], 3) = "" Then ContinueLoop
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '"' & _Hub_JsonEsc($aAll[$i]) & '"'
    Next
    $sJson &= "]"
    Return '{"status":"ok","lines":' & $sJson & "}"
EndFunc

Func _Hub_HistLoad($sFileName)
    If $sFileName = "" Then Return '{"status":"error","message":"nom_fichier_vide"}'
    If StringInStr($sFileName, "..") Or StringInStr($sFileName, "/") Or StringInStr($sFileName, "\") Then
        Return '{"status":"error","message":"chemin_invalide"}'
    EndIf
    Local $sPath = @ScriptDir & "\" & $sFileName
    If Not FileExists($sPath) Then
        Return '{"status":"error","message":"fichier_introuvable","file":"' & _
            _Hub_JsonEsc($sFileName) & '"}'
    EndIf
    Return '{"status":"ok","file":"' & _Hub_JsonEsc($sFileName) & '","data":' & _
        FileRead($sPath) & "}"
EndFunc

Func _Hub_HistSave($sFileName, $sData)
    If $sFileName = "" Or $sData = "" Then Return '{"status":"error","message":"params_vides"}'
    If StringInStr($sFileName, "..") Or StringInStr($sFileName, "/") Or StringInStr($sFileName, "\") Then
        Return '{"status":"error","message":"chemin_invalide"}'
    EndIf
    Local $hFile = FileOpen(@ScriptDir & "\" & $sFileName, 2 + 128)
    If $hFile = -1 Then Return '{"status":"error","message":"impossible_ecrire"}'
    FileWrite($hFile, $sData)
    FileClose($hFile)
    Return '{"status":"ok","file":"' & _Hub_JsonEsc($sFileName) & '"}'
EndFunc

Func _Hub_StorageInfo()
    Local $sOut = "["
    Local $aDrives = DriveGetDrive("ALL")
    If Not @error Then
        For $i = 1 To $aDrives[0]
            If $sOut <> "[" Then $sOut &= ","
            $sOut &= '{"drive":"' & _Hub_JsonEsc($aDrives[$i]) & '"' & _
                ',"freeMB":' & Round(DriveSpaceFree($aDrives[$i])) & _
                ',"totalMB":' & Round(DriveSpaceTotal($aDrives[$i])) & "}"
        Next
    EndIf
    $sOut &= "]"
    Return '{"status":"ok","drives":' & $sOut & "}"
EndFunc

Func _Hub_ActionListJSON()
    Return '{"status":"ok","port":' & $g_iHubPort & ',"edocAvailable":' & _
        ($g_bEdocAvailable ? "true" : "false") & _
        ',"actions":["PING","MAIL_SCAN","MAIL_OPEN","MAIL_COMPOSE",' & _
        '"LOG_LIST","LOG_CLEAR","HIST_LOAD","HIST_SAVE",' & _
        '"EDOC_INIT","EDOC_SCAN","EDOC_UPLOAD",' & _
        '"STORAGE_INFO","DISK_FREE","FILE_READ","FILE_WRITE"]}'
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; UTILITAIRES HTTP
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_Send($iSocket, $iCode, $sBody)
    Local $sStatus = "200 OK"
    If $iCode = 400 Then $sStatus = "400 Bad Request"
    If $iCode = 404 Then $sStatus = "404 Not Found"
    If $iCode = 500 Then $sStatus = "500 Internal Server Error"
    Local $bBin = StringToBinary($sBody, 4)
    Local $sResp = "HTTP/1.1 " & $sStatus & @CRLF & _
        "Content-Type: application/json; charset=utf-8" & @CRLF & _
        "Access-Control-Allow-Origin: *" & @CRLF & _
        "Access-Control-Allow-Methods: GET, POST, OPTIONS" & @CRLF & _
        "Access-Control-Allow-Headers: Content-Type" & @CRLF & _
        "Content-Length: " & BinaryLen($bBin) & @CRLF & _
        "Connection: close" & @CRLF & @CRLF
    TCPSend($iSocket, StringToBinary($sResp, 4))
    TCPSend($iSocket, $bBin)
EndFunc

Func _Hub_SendRaw($iSocket, $sRaw)
    TCPSend($iSocket, StringToBinary($sRaw, 4))
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; UTILITAIRES JSON LÉGERS
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_JsonGet($sJson, $sKey)
    Local $aMatch = StringRegExp($sJson, '"' & $sKey & '"\s*:\s*"((?:[^"\\]|\\.)*)"', 1)
    If Not @error Then
        Local $sVal = $aMatch[0]
        $sVal = StringReplace($sVal, '\"', '"')
        $sVal = StringReplace($sVal, "\\", "\")
        $sVal = StringReplace($sVal, "\n", @LF)
        $sVal = StringReplace($sVal, "\r", @CR)
        Return $sVal
    EndIf
    Local $aNum = StringRegExp($sJson, '"' & $sKey & '"\s*:\s*([0-9trueflsnul]+)', 1)
    If Not @error Then Return $aNum[0]
    Return ""
EndFunc

Func _Hub_JsonEsc($s)
    $s = StringReplace($s, "\", "\\")
    $s = StringReplace($s, '"', '\"')
    $s = StringReplace($s, @CR, "\r")
    $s = StringReplace($s, @LF, "\n")
    $s = StringReplace($s, @TAB, "\t")
    Return $s
EndFunc

; ══════════════════════════════════════════════════════════════════════════
; AUDIT LOG
; ══════════════════════════════════════════════════════════════════════════

Func _Hub_Log($sLevel, $sMsg)
    Local $sLine = "[" & @YEAR & "-" & @MON & "-" & @MDAY & " " & _
        @HOUR & ":" & @MIN & ":" & @SEC & "] " & "[" & $sLevel & "] " & $sMsg
    Local $hFile = FileOpen($g_sHubLogFile, 1 + 128)
    If $hFile <> -1 Then
        FileWriteLine($hFile, $sLine)
        FileClose($hFile)
    EndIf
EndFunc

Func _Hub_FmtDate($oDate)
    Local $sRaw = String($oDate)
    If StringLen($sRaw) >= 19 Then Return StringLeft($sRaw, 19)
    Local $sDate = _DateTimeFormat($oDate, 2)
    If @error Then Return ""
    Return $sDate
EndFunc

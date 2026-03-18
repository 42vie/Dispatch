; ╔══════════════════════════════════════════════════════════════════════════╗
; ║  DispatchMaster 2.0 - Serveur Local & Hub AutoIt                         ║
; ╚══════════════════════════════════════════════════════════════════════════╝
#include <File.au3>
#include <String.au3>
#include <Date.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

; ═══════════════════════════════════════════════════════════════════════════
; VARIABLES GLOBALES
; ═══════════════════════════════════════════════════════════════════════════
Global $g_iTrackCount   = 0
Global $g_aTrackIDs[1]
Global $g_hTracker      = 0
Global $g_idTrackProg = 0
Global $g_idTrackLbl  = 0
Global $g_idTrackLV   = 0
Global $bFC_Stop        = False
Global $bFC_Pause       = False
Global $bFC_Skip        = False
Global $iFC_StepCurrent = 0
Global $g_sFC_AuditLog  = ""
Global $g_bFC_Audit     = True
Global $bCOMAT_Stop  = False
Global $bCOMAT_Pause = False
Global $bCOMAT_Skip  = False

; GUI Batch control buttons
Global $g_idBtnPause  = 0
Global $g_idBtnPlay   = 0
Global $g_idBtnSkip   = 0
Global $g_idBtnStop   = 0
Global $g_idBatchInfo = 0

Global $sClassFileOpen  = "[CLASS:#32770; TITLE:Open]"
Global $sClassMenu      = "[CLASS:TfmMenuSelection]"
Global $sClassInput     = "[CLASS:TfmInput]"
Global $idToolbar       = "[CLASS:TToolBar; INSTANCE:1]"
Global $DELAY_MEDIUM    = 500
Global $DELAY_LONG      = 1000
Global $DELAY_ETMS_LOAD = 1500

Global Const $COMAT_LOG_CTRL   = "[CLASS:TEIEdit; INSTANCE:91]"
Global Const $COMAT_DELAY_S    = 150
Global Const $COMAT_DELAY_M    = 300
Global Const $COMAT_DELAY_L    = 500
Global Const $COMAT_DELAY_LOAD = 3000

Opt("TrayIconDebug", 0)
Opt("TrayMenuMode", 3)        ; pas de menu Pause/Exit par défaut
Opt("TrayAutoPause", 0)       ; ne jamais auto-pauser le script
TCPStartup()

Global $g_iPort       = 9500
Global $g_iMainSocket = -1

; Fermer les instances en double
Local $iPID  = @AutoItPID
Local $aList = ProcessList(@ScriptName)
For $i = 1 To $aList[0][0]
    If $aList[$i][1] <> $iPID Then ProcessClose($aList[$i][1])
Next
Sleep(400)

; Essai port 9500, puis ports aléatoires si occupé
$g_iMainSocket = TCPListen("127.0.0.1", 9500)
If $g_iMainSocket <> -1 Then
    $g_iPort = 9500
Else
    Local $iFound = 0
    Local $iP = 0
    For $iP = 1 To 50
        Local $iTryPort = Random(10000, 60000, 1)
        $g_iMainSocket = TCPListen("127.0.0.1", $iTryPort)
        If $g_iMainSocket <> -1 Then
            $g_iPort = $iTryPort
            $iFound  = 1
            ExitLoop
        EndIf
    Next
    If $iFound = 0 Then
        MsgBox(16, "Erreur Serveur", "Impossible de démarrer le serveur TCP." & @CRLF & "Vérifiez votre pare-feu ou redémarrez le PC.")
        Exit
    EndIf
EndIf

Global $g_sSaveFile = @ScriptDir & "\historique_dispatch.json"
Global $g_sAuditLog = @ScriptDir & "\logs\dispatch_audit.log"
Global $g_iAuditCheckTimer = 0
Global Const $AUDIT_CHECK_INTERVAL = 60000 ; 60 secondes entre chaque vérification silencieuse
Global $g_sStatusFile   = @ScriptDir & "\dispatch_status.json"
Global $g_sDataFile     = @ScriptDir & "\dispatch_data.json"
Global $g_sContactsFile = @ScriptDir & "\dispatch_contacts.json"
Global $g_sHtmlFile = @ScriptDir & "\interface.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\Interface.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\Interface_v2.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\DispatchInterface.html"
If Not FileExists($g_sHtmlFile) Then
    MsgBox(16, "Erreur", "Fichier HTML introuvable dans : " & @ScriptDir & @CRLF & "Attendu : interface.html")
    Exit
EndIf
Global $g_sHTML = FileRead($g_sHtmlFile)
; Pré-calculer le HTML avec le bon port (une seule fois)
$g_sHTML = StringReplace($g_sHTML, "127.0.0.1:9500", "127.0.0.1:" & $g_iPort)
$g_sHTML = StringReplace($g_sHTML, "127.0.0.1:8080", "127.0.0.1:" & $g_iPort)

ShellExecute("http://127.0.0.1:" & $g_iPort)

; Créer les dossiers nécessaires
If Not FileExists(@ScriptDir & "\logs") Then DirCreate(@ScriptDir & "\logs")
If Not FileExists(@ScriptDir & "\backups") Then DirCreate(@ScriptDir & "\backups")
_AuditLog("INFO", "Serveur démarré sur le port " & $g_iPort)
$g_iAuditCheckTimer = TimerInit()

; Boucle principale
While 1
    Local $iClientSocket = TCPAccept($g_iMainSocket)
    If $iClientSocket <> -1 Then _HandleClient($iClientSocket)

    ; Vérification silencieuse en arrière-plan (chaque minute)
    If TimerDiff($g_iAuditCheckTimer) > $AUDIT_CHECK_INTERVAL Then
        _SilentHealthCheck()
        $g_iAuditCheckTimer = TimerInit()
    EndIf

    Sleep(10)
WEnd

; ==============================================================================
; SERVEUR HTTP
; ==============================================================================
Func _HandleClient($iSocket)
    Local $sHeader = ""
    Local $iContentLength = 0
    Local $sBody = ""
    Local $hTimeout = TimerInit()
    Local Const $RECV_TIMEOUT = 3000 ; 3s max pour recevoir les headers

    ; ── Recevoir les headers (avec timeout) ──
    While TimerDiff($hTimeout) < $RECV_TIMEOUT
        Local $sRecv = TCPRecv($iSocket, 8192)
        If @error Then ExitLoop
        If $sRecv <> "" Then
            $sHeader &= $sRecv
            Local $iHeaderEnd = StringInStr($sHeader, @CRLF & @CRLF)
            If $iHeaderEnd > 0 Then
                Local $aMatch = StringRegExp($sHeader, "(?i)Content-Length:\s*(\d+)", 3)
                If IsArray($aMatch) Then $iContentLength = Int($aMatch[0])
                $sBody   = StringTrimLeft($sHeader, $iHeaderEnd + 3)
                $sHeader = StringLeft($sHeader, $iHeaderEnd - 1)
                ExitLoop
            EndIf
        EndIf
        Sleep(5)
    WEnd

    If StringInStr($sHeader, @CRLF) = 0 Then
        TCPCloseSocket($iSocket)
        Return
    EndIf

    ; ── Recevoir le body (avec timeout) ──
    $hTimeout = TimerInit()
    While StringLen($sBody) < $iContentLength And TimerDiff($hTimeout) < $RECV_TIMEOUT
        Local $sRecv2 = TCPRecv($iSocket, 8192)
        If @error Then ExitLoop
        If $sRecv2 <> "" Then
            $sBody &= $sRecv2
            $hTimeout = TimerInit() ; reset timer si on reçoit des données
        EndIf
        Sleep(5)
    WEnd

    Local $aLines = StringSplit($sHeader, @CRLF, 1)
    If $aLines[0] < 1 Then Return TCPCloseSocket($iSocket)
    Local $aTop = StringSplit($aLines[1], " ")
    If $aTop[0] < 2 Then Return TCPCloseSocket($iSocket)
    Local $sMethod = $aTop[1]
    Local $sURL = $aTop[2]

    ; ── Gérer preflight CORS (OPTIONS) ──
    If $sMethod = "OPTIONS" Then
        Local $sCors = "HTTP/1.1 204 No Content" & @CRLF & _
            "Access-Control-Allow-Origin: *" & @CRLF & _
            "Access-Control-Allow-Methods: GET, POST, OPTIONS" & @CRLF & _
            "Access-Control-Allow-Headers: Content-Type" & @CRLF & _
            "Access-Control-Max-Age: 86400" & @CRLF & _
            "Content-Length: 0" & @CRLF & _
            "Connection: close" & @CRLF & @CRLF
        TCPSend($iSocket, StringToBinary($sCors, 4))
        TCPCloseSocket($iSocket)
        Return
    EndIf

    If $sURL = "/" Then
        _SendHttpResponse($iSocket, 200, "text/html", $g_sHTML)

    ElseIf $sURL = "/api/ping" Then
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')

    ElseIf $sURL = "/api/load" Then
        Local $sJson = "{}"
        If FileExists($g_sSaveFile) Then
            Local $hJsonRead = FileOpen($g_sSaveFile, 256) ; 256 = UTF-8 sans BOM
            If $hJsonRead <> -1 Then
                $sJson = FileRead($hJsonRead)
                FileClose($hJsonRead)
            EndIf
        EndIf
        _SendHttpResponse($iSocket, 200, "application/json", $sJson)

    ElseIf $sURL = "/api/save" Then
        Local $hFile = FileOpen($g_sSaveFile, 2 + 256) ; 256 = UTF-8 sans BOM
        FileWrite($hFile, $sBody)
        FileClose($hFile)
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')

    ; ── Fichiers séparés : STATUS ──
    ElseIf $sURL = "/api/save-status" Then
        _AuditLog("SAVE", "status — " & StringLen($sBody) & " bytes")
        Local $hFileS = FileOpen($g_sStatusFile, 2 + 256)
        FileWrite($hFileS, $sBody)
        FileClose($hFileS)
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')

    ElseIf $sURL = "/api/load-status" Then
        Local $sJsonS = "[]"
        If FileExists($g_sStatusFile) Then
            Local $hReadS = FileOpen($g_sStatusFile, 256)
            If $hReadS <> -1 Then
                $sJsonS = FileRead($hReadS)
                FileClose($hReadS)
            EndIf
        EndIf
        _SendHttpResponse($iSocket, 200, "application/json", $sJsonS)

    ; ── Fichiers séparés : DATA ──
    ElseIf $sURL = "/api/save-data" Then
        _AuditLog("SAVE", "data — " & StringLen($sBody) & " bytes")
        _BackupRotate($g_sDataFile, 5)
        Local $hFileD = FileOpen($g_sDataFile, 2 + 256)
        FileWrite($hFileD, $sBody)
        FileClose($hFileD)
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')

    ElseIf $sURL = "/api/load-data" Then
        Local $sJsonD = "{}"
        If FileExists($g_sDataFile) Then
            Local $hReadD = FileOpen($g_sDataFile, 256)
            If $hReadD <> -1 Then
                $sJsonD = FileRead($hReadD)
                FileClose($hReadD)
            EndIf
        EndIf
        _SendHttpResponse($iSocket, 200, "application/json", $sJsonD)

    ; ── Fichiers séparés : CONTACTS (multi-chunks) ──
    ElseIf $sURL = "/api/save-contacts" Then
        _AuditLog("SAVE", "contacts — " & StringLen($sBody) & " bytes")
        ; Body = {chunk:0, total:1, data:[...]}
        ; Si chunk=0 et total=1 → un seul fichier dispatch_contacts.json
        ; Si multi-chunks → dispatch_contacts_0.json, dispatch_contacts_1.json, etc.
        Local $sChunk = _GetJsonValue($sBody, "chunk")
        Local $sTotal = _GetJsonValue($sBody, "total")
        Local $sData  = _GetJsonArrayValue($sBody, "data")
        If $sTotal = "1" Or $sTotal = "" Then
            Local $hFileC = FileOpen($g_sContactsFile, 2 + 256)
            FileWrite($hFileC, $sData)
            FileClose($hFileC)
        Else
            Local $sChunkFile = @ScriptDir & "\dispatch_contacts_" & $sChunk & ".json"
            Local $hFileC2 = FileOpen($sChunkFile, 2 + 256)
            FileWrite($hFileC2, $sData)
            FileClose($hFileC2)
            ; Sauvegarder le nombre total de chunks
            Local $hMeta = FileOpen(@ScriptDir & "\dispatch_contacts_meta.json", 2 + 256)
            FileWrite($hMeta, '{"total":' & $sTotal & '}')
            FileClose($hMeta)
        EndIf
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')

    ElseIf $sURL = "/api/load-contacts" Then
        Local $sJsonC = "[]"
        ; Vérifier si multi-chunks
        Local $sMetaFile = @ScriptDir & "\dispatch_contacts_meta.json"
        If FileExists($sMetaFile) Then
            Local $hMeta2 = FileOpen($sMetaFile, 256)
            Local $sMeta = FileRead($hMeta2)
            FileClose($hMeta2)
            Local $iTotalC = Number(_GetJsonValue($sMeta, "total"))
            If $iTotalC > 1 Then
                ; Charger et fusionner tous les chunks
                Local $sAll = ""
                For $i = 0 To $iTotalC - 1
                    Local $sChkFile = @ScriptDir & "\dispatch_contacts_" & $i & ".json"
                    If FileExists($sChkFile) Then
                        Local $hChk = FileOpen($sChkFile, 256)
                        Local $sPart = FileRead($hChk)
                        FileClose($hChk)
                        ; Enlever les crochets pour fusionner
                        $sPart = StringRegExpReplace($sPart, "^\s*\[", "")
                        $sPart = StringRegExpReplace($sPart, "\]\s*$", "")
                        If $sAll <> "" And $sPart <> "" Then $sAll &= ","
                        $sAll &= $sPart
                    EndIf
                Next
                $sJsonC = "[" & $sAll & "]"
            EndIf
        ElseIf FileExists($g_sContactsFile) Then
            Local $hReadC = FileOpen($g_sContactsFile, 256)
            If $hReadC <> -1 Then
                $sJsonC = FileRead($hReadC)
                FileClose($hReadC)
            EndIf
        EndIf
        _SendHttpResponse($iSocket, 200, "application/json", $sJsonC)

    ElseIf StringLeft($sURL, 13) = "/api/net-save" Then
        ; /api/net-save?path=F:\...\state.json — le body EST le JSON à écrire
        Local $sNetPath = StringMid($sURL, 20) ; après "/api/net-save?path="
        $sNetPath = _URIDecode($sNetPath)
        If $sNetPath <> "" Then
            _Net_SaveState($sNetPath, $sBody)
            _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
        Else
            _SendHttpResponse($iSocket, 400, "application/json", '{"error":"missing path"}')
        EndIf

    ElseIf StringLeft($sURL, 13) = "/api/net-load" Then
        ; /api/net-load?path=F:\...\state.json — retourne le contenu du fichier
        Local $sNetPath2 = StringMid($sURL, 20) ; après "/api/net-load?path="
        $sNetPath2 = _URIDecode($sNetPath2)
        Local $sNetJSON = _Net_LoadState($sNetPath2)
        _SendHttpResponse($iSocket, 200, "application/json", $sNetJSON)

    ElseIf $sURL = "/api/action" Then
        Local $sAction = _GetJsonValue($sBody, "action")

        ; Variables déclarées AVANT Switch (Local interdit dans Case)
        Local $sCmd_a    = ""
        Local $sFile_a   = ""
        Local $sClient_a = ""
        Local $sEmail_a  = ""
        Local $sTrack_a  = ""
        Local $sLogErr_a = ""
        Local $sData_a   = ""
        Local $sPath_a   = ""
        Local $sState_a  = ""
        Local $sJSON_a   = ""
        Local $sIni_a    = ""
        Local $sCpRaw_a  = ""

        _AuditLog("ACTION", $sAction & " — " & StringLeft($sBody, 200))

        ; ══ RÉPONSE IMMÉDIATE — l'HTML est débloqué avant l'exécution ══
        ; Pour les actions longues (ETMS, COMAT, FC), on répond OK tout de suite
        ; puis on exécute l'action. Le HTML n'attend plus.
        If $sAction = "ETMS_CMD" Or $sAction = "MAIL_RDV" Or $sAction = "KANBAN_2" Or _
           $sAction = "KANBAN_4" Or $sAction = "KANBAN_5" Or $sAction = "KANBAN_6" Or _
           $sAction = "COMAT_MULTI" Or $sAction = "COMAT_SOLO" Or $sAction = "BATCH_CP" Then
            _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","async":true}')
            TCPCloseSocket($iSocket)
            $iSocket = -1  ; Marquer comme déjà fermé
        EndIf

        Switch $sAction

            Case "ETMS_CMD"
                $sCmd_a   = _GetJsonValue($sBody, "cmd")
                $sFile_a  = _GetJsonValue($sBody, "file")
                _ActionETMS($sCmd_a, $sFile_a)

            Case "MAIL_RDV"
                $sClient_a = _GetJsonValue($sBody, "client")
                $sEmail_a  = _GetJsonValue($sBody, "email")
                $sTrack_a  = _GetJsonValue($sBody, "file")
                $sLogErr_a = ""
                _Mail_DemandeRDV($sTrack_a, $sClient_a, $sEmail_a, $sLogErr_a)

            Case "KANBAN_2"
                $sData_a = _GetJsonValue($sBody, "data")
                _Batch_Mails_RDV($sData_a)

            Case "KANBAN_4"
                $sData_a = _GetJsonValue($sBody, "data")
                _Batch_Mails_Alerte($sData_a)

            Case "KANBAN_5"
                $sData_a = _GetJsonValue($sBody, "data")
                _Batch_FC($sData_a)

            Case "KANBAN_6", "COMAT_MULTI"
                $sData_a = _GetJsonValue($sBody, "data")
                _Batch_COMAT($sData_a)

            Case "COMAT_SOLO"
                $sFile_a = _GetJsonValue($sBody, "file")
                _Action_COMAT_Solo($sFile_a)

            Case "FC_PAUSE"
                $bFC_Pause = True

            Case "FC_PLAY"
                $bFC_Pause = False

            Case "FC_SKIP"
                $bFC_Skip = True
                $bFC_Pause = False

            Case "FC_STOP"
                $bFC_Stop = True
                $bFC_Pause = False

            Case "COMAT_PAUSE"
                $bCOMAT_Pause = True

            Case "COMAT_PLAY"
                $bCOMAT_Pause = False

            Case "COMAT_SKIP"
                $bCOMAT_Skip = True
                $bCOMAT_Pause = False

            Case "COMAT_STOP"
                $bCOMAT_Stop = True
                $bCOMAT_Pause = False

            Case "BATCH_CP"
                $sData_a = _GetJsonValue($sBody, "data")
                _Batch_Mails_CP($sData_a)

            Case "save-network-state"
                $sPath_a = _GetJsonValue($sBody, "path")
                ; Extraire tout ce qui est après "state": directement (évite le parsing lent)
                Local $iStatePos = StringInStr($sBody, '"state"')
                If $iStatePos > 0 Then
                    Local $iColonPos = StringInStr($sBody, ":", 0, 1, $iStatePos)
                    If $iColonPos > 0 Then
                        ; Prendre tout après "state": et retirer la dernière } du body
                        $sState_a = StringStripWS(StringMid($sBody, $iColonPos + 1), 3)
                        ; Retirer la } fermante du body JSON parent
                        If StringRight($sState_a, 1) = "}" Then $sState_a = StringTrimRight($sState_a, 1)
                        $sState_a = StringStripWS($sState_a, 2)
                    EndIf
                EndIf
                If $sPath_a <> "" And $sState_a <> "" Then _Net_SaveState($sPath_a, $sState_a)

            Case "load-network-state"
                $sPath_a = _GetJsonValue($sBody, "path")
                $sJSON_a = _Net_LoadState($sPath_a)
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","state":' & $sJSON_a & '}')
                TCPCloseSocket($iSocket)
                Return

            Case "CHECK_PDF"
                $sData_a = _GetJsonValue($sBody, "data")
                Local $sCheminCheck = "F:\CDG\PRODUCT\TRANSCON\Shared\Clients\HPE\Pre-alertes\"
                ; Lire chemin personnalisé depuis config si disponible
                Local $sIniCheck = @ScriptDir & "\dispatch_config.ini"
                Local $sCfgPath = IniRead($sIniCheck, "PJ", "Path", "")
                If $sCfgPath <> "" Then $sCheminCheck = $sCfgPath & "\"
                Local $aCheckFiles = StringSplit($sData_a, "|")
                Local $sMissing = ""
                For $j = 1 To $aCheckFiles[0]
                    Local $sF = StringStripWS($aCheckFiles[$j], 3)
                    If $sF <> "" And Not FileExists($sCheminCheck & $sF & ".pdf") Then
                        If $sMissing <> "" Then $sMissing &= "|"
                        $sMissing &= $sF
                    EndIf
                Next
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","missing":"' & $sMissing & '"}')
                TCPCloseSocket($iSocket)
                Return

            Case "save-pj-config"
                $sIni_a = @ScriptDir & "\dispatch_config.ini"
                IniWrite($sIni_a, "PJ", "Path",         _GetJsonValue($sBody, "path"))
                IniWrite($sIni_a, "PJ", "RDV_Ext",      _GetJsonValue($sBody, "rdvExt"))
                IniWrite($sIni_a, "PJ", "Prealert_Ext", _GetJsonValue($sBody, "prealertExt"))
                IniWrite($sIni_a, "PJ", "UPS_Folder",   _GetJsonValue($sBody, "upsFolder"))
                IniWrite($sIni_a, "PJ", "DGS_Folder",   _GetJsonValue($sBody, "dgsFolder"))

            Case "save-config"
                $sIni_a = @ScriptDir & "\dispatch_config.ini"
                IniWrite($sIni_a, "Network", "StatePath",    _GetJsonValue($sBody, "statePath"))
                IniWrite($sIni_a, "Network", "OperatorName", _GetJsonValue($sBody, "operatorName"))

            Case "save-cp-config"
                $sIni_a  = @ScriptDir & "\dispatch_config.ini"
                $sCpRaw_a = _GetJsonValue($sBody, "cpConfig")
                IniWrite($sIni_a, "CP", "Config", $sCpRaw_a)

            Case "DIAG"
                _RunDiagnostic()

            Case "CLEAN_CONTACTS"
                _CleanContactsFiles()

            Case "STORAGE_INFO"
                Local $sInfo = _GetStorageInfo()
                _SendHttpResponse($iSocket, 200, "application/json", $sInfo)
                TCPCloseSocket($iSocket)
                Return

        EndSwitch

        ; Réponse seulement si pas déjà envoyée (actions async)
        If $iSocket <> -1 Then
            _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
        EndIf

    Else
        _SendHttpResponse($iSocket, 404, "text/plain", "Not Found")
    EndIf

    If $iSocket <> -1 Then TCPCloseSocket($iSocket)
EndFunc

Func _SendHttpResponse($iSocket, $iCode, $sContentType, $sData)
    Local $sStatus = "200 OK"
    If $iCode = 400 Then $sStatus = "400 Bad Request"
    If $iCode = 404 Then $sStatus = "404 Not Found"
    Local $bData    = StringToBinary($sData, 4)
    Local $iLen     = BinaryLen($bData)
    Local $sHeaders = "HTTP/1.1 " & $sStatus & @CRLF & _
                      "Content-Type: " & $sContentType & "; charset=UTF-8" & @CRLF & _
                      "Content-Length: " & $iLen & @CRLF & _
                      "Access-Control-Allow-Origin: *" & @CRLF & _
                      "Access-Control-Allow-Headers: Content-Type" & @CRLF & _
                      "Cache-Control: no-cache" & @CRLF & _
                      "Connection: close" & @CRLF & @CRLF
    ; Envoyer par blocs pour éviter les envois partiels sur gros payloads
    Local $bAll = StringToBinary($sHeaders, 4) & $bData
    Local $iTotal = BinaryLen($bAll)
    Local $iSent = 0
    While $iSent < $iTotal
        Local $bChunk = BinaryMid($bAll, $iSent + 1, 8192)
        Local $iRes = TCPSend($iSocket, $bChunk)
        If @error Then ExitLoop
        If $iRes > 0 Then
            $iSent += $iRes
        Else
            Sleep(5)
        EndIf
    WEnd
EndFunc

; ==============================================================================
; UTILITAIRES
; ==============================================================================
Func _GetJsonValue($sJson, $sKey)
    ; 1. Essayer valeur string : "key":"value"
    Local $aMatch = StringRegExp($sJson, '(?i)"' & $sKey & '"\s*:\s*"([^"]*)"', 3)
    If IsArray($aMatch) Then Return $aMatch[0]
    ; 2. Essayer objet/array JSON : "key":{...} ou "key":[...]
    Local $iPos = StringInStr($sJson, '"' & $sKey & '"')
    If $iPos > 0 Then
        Local $iColon = StringInStr($sJson, ":", 0, 1, $iPos)
        If $iColon > 0 Then
            Local $sAfter = StringStripWS(StringMid($sJson, $iColon + 1), 1)
            Local $sFirst = StringLeft($sAfter, 1)
            If $sFirst = "{" Or $sFirst = "[" Then
                ; Trouver la fermeture correspondante en comptant les niveaux
                Local $sOpen = $sFirst, $sClose = ($sFirst = "{") ? "}" : "]"
                Local $iDepth = 0, $bInStr = False
                For $i = 1 To StringLen($sAfter)
                    Local $c = StringMid($sAfter, $i, 1)
                    If $c = '"' And ($i = 1 Or StringMid($sAfter, $i - 1, 1) <> "\") Then $bInStr = Not $bInStr
                    If Not $bInStr Then
                        If $c = $sOpen Then $iDepth += 1
                        If $c = $sClose Then
                            $iDepth -= 1
                            If $iDepth = 0 Then Return StringLeft($sAfter, $i)
                        EndIf
                    EndIf
                Next
            EndIf
            ; 3. Essayer valeur numérique/booléenne
            Local $aNum = StringRegExp($sAfter, '^([0-9.eE+\-]+|true|false|null)', 3)
            If IsArray($aNum) Then Return $aNum[0]
        EndIf
    EndIf
    Return ""
EndFunc

; Alias pour récupérer un array JSON (utilise _GetJsonValue qui gère déjà les [...])
Func _GetJsonArrayValue($sJson, $sKey)
    Local $sVal = _GetJsonValue($sJson, $sKey)
    If $sVal = "" Then Return "[]"
    Return $sVal
EndFunc

Func _GetWindowETMS()
    Return WinGetHandle("[CLASS:TfmBrowser]")
EndFunc

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
Func _BackupRotate($sFile, $iMax)
    If Not FileExists($sFile) Then Return
    Local $sBackDir = @ScriptDir & "\backups"
    Local $sBase = StringRegExpReplace($sFile, ".*\\", "")
    ; Rotation : supprimer le plus ancien, décaler les autres
    Local $sOldest = $sBackDir & "\" & $sBase & "." & $iMax & ".bak"
    If FileExists($sOldest) Then FileDelete($sOldest)
    For $b = $iMax - 1 To 1 Step -1
        Local $sSrc = $sBackDir & "\" & $sBase & "." & $b & ".bak"
        Local $sDst = $sBackDir & "\" & $sBase & "." & ($b + 1) & ".bak"
        If FileExists($sSrc) Then FileMove($sSrc, $sDst, 1)
    Next
    ; Copier le fichier actuel en .1.bak
    FileCopy($sFile, $sBackDir & "\" & $sBase & ".1.bak", 1)
EndFunc

; ==============================================================================
; AUDIT FC — Diagnostic détaillé pour FileClosing
; ==============================================================================
Func _FC_AuditInit($sLabel)
    $g_sFC_AuditLog = ""
    _FC_AuditLog("====== AUDIT FC : " & $sLabel & " ======")
    _FC_AuditLog("PC       : " & @ComputerName)
    _FC_AuditLog("User     : " & @UserName)
    _FC_AuditLog("OS       : " & @OSVersion & " " & @OSArch)
    _FC_AuditLog("RAM Free : " & Round(MemGetStats()[2] / 1024, 0) & " MB / " & Round(MemGetStats()[1] / 1024, 0) & " MB")
    _FC_AuditLog("CPU      : " & @CPUArch)
    Local $aProc = ProcessList("ETMS.exe")
    If IsArray($aProc) Then
        _FC_AuditLog("ETMS PID : " & ($aProc[0][0] > 0 ? $aProc[1][1] : "NON TROUVE"))
    Else
        _FC_AuditLog("ETMS PID : NON TROUVE")
    EndIf
EndFunc

Func _FC_AuditLog($sMsg)
    If Not $g_bFC_Audit Then Return
    Local $sLine = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & "  " & $sMsg
    $g_sFC_AuditLog &= $sLine & @CRLF
EndFunc

Func _FC_AuditStep($iStep, $sDesc)
    _FC_AuditLog("── STEP " & $iStep & " : " & $sDesc & " ──")
EndFunc

Func _FC_AuditWinState($sClass, $sLabel)
    Local $bExists = WinExists($sClass)
    Local $sState  = "exists=" & $bExists
    If $bExists Then
        Local $hW  = WinGetHandle($sClass)
        Local $aPos = WinGetPos($hW)
        Local $sTitle = WinGetTitle($hW)
        $sState &= " | hwnd=" & $hW & " | title=" & $sTitle
        If IsArray($aPos) Then $sState &= " | pos=" & $aPos[0] & "," & $aPos[1] & " size=" & $aPos[2] & "x" & $aPos[3]
        $sState &= " | state=" & WinGetState($hW)
    EndIf
    _FC_AuditLog("  WIN[" & $sLabel & "] " & $sState)
EndFunc

Func _FC_AuditCtrl($hWnd, $sCtrl, $sLabel)
    Local $sTxt = ControlGetText($hWnd, "", $sCtrl)
    Local $hCtrl = ControlGetHandle($hWnd, "", $sCtrl)
    _FC_AuditLog("  CTRL[" & $sLabel & "] handle=" & $hCtrl & " | text='" & StringLeft($sTxt, 120) & "'")
EndFunc

Func _FC_AuditTiming($sLabel, $nMs)
    Local $sSuffix = ""
    If $nMs > 5000 Then $sSuffix = " *** LENT ***"
    If $nMs > 10000 Then $sSuffix = " *** TRES LENT ***"
    If $nMs > 20000 Then $sSuffix = " *** CRITIQUE ***"
    _FC_AuditLog("  TIMING[" & $sLabel & "] " & Round($nMs, 0) & " ms" & $sSuffix)
EndFunc

Func _FC_AuditFileCheck($sPath)
    _FC_AuditLog("  FILE[" & $sPath & "]")
    If FileExists($sPath) Then
        Local $iSize = FileGetSize($sPath)
        Local $sTime = FileGetTime($sPath, 0, 1)
        _FC_AuditLog("    exists=TRUE | size=" & $iSize & " octets (" & Round($iSize/1024, 1) & " KB) | modified=" & $sTime)
    Else
        _FC_AuditLog("    exists=FALSE *** FICHIER INTROUVABLE ***")
        ; Verifier le dossier parent
        Local $sDir = StringRegExpReplace($sPath, "\\[^\\]+$", "")
        If FileExists($sDir) Then
            _FC_AuditLog("    dossier parent OK : " & $sDir)
            ; Lister les .eds dans le dossier
            Local $hSearch = FileFindFirstFile($sDir & "\*.eds")
            If $hSearch <> -1 Then
                Local $sFiles = ""
                Local $iCount = 0
                While 1
                    Local $sFile = FileFindNextFile($hSearch)
                    If @error Then ExitLoop
                    $sFiles &= $sFile & ", "
                    $iCount += 1
                WEnd
                FileClose($hSearch)
                _FC_AuditLog("    .eds trouves (" & $iCount & ") : " & StringTrimRight($sFiles, 2))
            Else
                _FC_AuditLog("    AUCUN .eds dans le dossier !")
            EndIf
        Else
            _FC_AuditLog("    *** DOSSIER PARENT INTROUVABLE : " & $sDir & " ***")
        EndIf
    EndIf
EndFunc

Func _FC_AuditSave($sNum)
    If $g_sFC_AuditLog = "" Then Return ""
    Local $sDir = @ScriptDir & "\logs"
    If Not FileExists($sDir) Then DirCreate($sDir)
    Local $sFile = $sDir & "\FC_AUDIT_" & $sNum & "_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".log"
    Local $hFile = FileOpen($sFile, 2)
    If $hFile = -1 Then Return ""
    FileWrite($hFile, $g_sFC_AuditLog)
    FileClose($hFile)
    Return $sFile
EndFunc

Func _FC_AuditShow($sNum)
    If Not $g_bFC_Audit Then Return
    Local $sFile = _FC_AuditSave($sNum)
    If $sFile = "" Then Return

    ; GUI avec le rapport complet
    Local $hGUI = GUICreate("AUDIT FC — " & $sNum, 750, 550, -1, -1)
    GUISetBkColor(0x1E1E1E, $hGUI)
    GUISetFont(9, 400, 0, "Consolas")
    Local $idEdit = GUICtrlCreateEdit($g_sFC_AuditLog, 5, 5, 740, 495, BitOR(0x0004, 0x0800, 0x00200000))
    GUICtrlSetBkColor($idEdit, 0x1E1E1E)
    GUICtrlSetColor($idEdit, 0x00FF00)
    Local $idBtnCopy = GUICtrlCreateButton("Copier", 5, 505, 120, 35)
    Local $idBtnOpen = GUICtrlCreateButton("Ouvrir le .log", 130, 505, 150, 35)
    Local $idBtnClose = GUICtrlCreateButton("Fermer", 625, 505, 120, 35)
    GUISetState(@SW_SHOW, $hGUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idBtnClose
                ExitLoop
            Case $idBtnCopy
                ClipPut($g_sFC_AuditLog)
                ToolTip("Copié !", Default, Default, "Audit FC", 1)
                Sleep(1000)
                ToolTip("")
            Case $idBtnOpen
                ShellExecute($sFile)
        EndSwitch
    WEnd
    GUIDelete($hGUI)
EndFunc

Func _FC_WaitIfPaused()
    While $bFC_Pause And Not $bFC_Stop And Not $bFC_Skip
        _Spinner("EN PAUSE — cliquez Play dans la fenêtre de contrôle")
        _Tracker_PollButtons()
        Sleep(80)
    WEnd
EndFunc

; Sleep réactif FC : découpe en tranches de 100ms + poll GUI
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

Func _COMAT_WaitIfPaused2()
    While $bCOMAT_Pause And Not $bCOMAT_Stop And Not $bCOMAT_Skip
        _Tracker_PollButtons()
        Sleep(100)
    WEnd
EndFunc

; Les HotKeys restent comme fallback (si la GUI n'est pas visible)
Func _HK_FC_PauseToggle()
    $bFC_Pause = Not $bFC_Pause
EndFunc

Func _HK_FC_Stop()
    $bFC_Stop = True
    $bFC_Pause = False
EndFunc

Func _HK_COMAT_PauseToggle()
    $bCOMAT_Pause = Not $bCOMAT_Pause
EndFunc

Func _HK_COMAT_Stop()
    $bCOMAT_Stop = True
    $bCOMAT_Pause = False
EndFunc

; ==============================================================================
; E.TMS ET EDOC
; ==============================================================================
Func _GetETMSInstance($hWnd)
    Local $sTitle = WinGetTitle($hWnd)
    If StringInStr($sTitle, "(LOG)")   Then Return "91"
    If StringInStr($sTitle, "(NOTES)") Then Return "83"
    If StringInStr($sTitle, "(REFS)")  Then Return "109"
    If StringInStr($sTitle, "(DIMST)") Then Return "300"
    If StringInStr($sTitle, "(HIST)")  Then Return "207"
    Return "91"
EndFunc

Func _ActionEDOC($sNumDossier)
    Local $hWndEdoc = WinGetHandle("[CLASS:TfmEdocViewerMainDlg]")
    If Not WinExists($hWndEdoc) Then Return False
    $sNumDossier = StringStripWS($sNumDossier, 8)
    ; Mode arrière-plan — pas de WinActivate
    ControlSetText($hWndEdoc, "", "[CLASS:Edit; INSTANCE:1]", $sNumDossier)
    ControlSend($hWndEdoc, "", "[CLASS:Edit; INSTANCE:1]", "{ENTER}")
    Return True
EndFunc

Func _ActionETMS($sBouton, $sNumDossier)
    If $sBouton = "EDOC" Then
        If _ActionEDOC($sNumDossier) Then Return
    EndIf
    Local $hWnd = WinGetHandle("[CLASS:TfmBrowser]")
    If Not WinExists($hWnd) Then
        _NotifyError("E.TMS", "E.TMS est fermé ou introuvable. Ouvrez E.TMS avant de lancer une action.")
        Return
    EndIf

    ; Préparer la commande
    Local $sCommande = $sBouton & " " & $sNumDossier
    If $sBouton = "LOG X" Then $sCommande = "LOG X"
    Local $sInst = _GetETMSInstance($hWnd)
    Local $sCtrl = "[CLASS:TEIEdit; INSTANCE:" & $sInst & "]"

    ; ══ MODE ARRIÈRE-PLAN : ControlSend/ControlSetText SANS WinActivate ══
    ; L'utilisateur garde le focus sur sa fenêtre active
    ; PgUp pour remonter en haut du champ de commande
    ControlSend($hWnd, "", $sCtrl, "{PGUP}")
    Sleep(30)
    ; Écrire la commande directement dans le contrôle (pas besoin de focus)
    ControlSetText($hWnd, "", $sCtrl, $sCommande)
    Sleep(30)
    ; Exécuter avec F8 envoyé au contrôle (pas au clavier global)
    ControlSend($hWnd, "", $sCtrl, "{F8}")
    _AuditLog("ETMS", "BG: " & $sCommande)
EndFunc

; ==============================================================================
; CALCUL DATES / JOURS OUVRÉS
; ==============================================================================
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
Func _Batch_Mails_RDV($sData)
    If $sData = "" Then Return
    Local $aJobs = StringSplit($sData, "|")
    Local $iOk = 0
    Local $iErr = 0
    Local $sLogErr = ""
    For $i = 1 To $aJobs[0]
        Local $aInfos = StringSplit($aJobs[$i], ";")
        If $aInfos[0] >= 3 Then
            If _Mail_DemandeRDV($aInfos[1], $aInfos[3], $aInfos[2], $sLogErr) Then
                $iOk += 1
                Sleep(1500)
            Else
                $iErr += 1
            EndIf
        EndIf
    Next
    MsgBox(64+262144, "Bilan RDV", $iOk & " mail(s)." & @CRLF & $iErr & " erreur(s)." & @CRLF & $sLogErr)
EndFunc

Func _Mail_DemandeRDV($sTracking, $sClient, $sEmail, ByRef $sLogErr)
    Local $sCheminBase = "F:\CDG\PRODUCT\TRANSCON\Shared\Clients\HPE\Pre-alertes\"
    Local $aCmds = StringSplit(StringRegExpReplace($sTracking, "[,;\s]*\+[,;\s]*|[,;]+", " "), " ", 2)
    Local $sCmdListe = ""
    Local $iNbCmd = 0
    Local $bFichiersOK = True
    Local $aFichiers[UBound($aCmds) + 1]
    For $c = 0 To UBound($aCmds) - 1
        Local $sCmd = StringStripWS($aCmds[$c], 3)
        If $sCmd <> "" Then
            $iNbCmd += 1
            $aFichiers[$iNbCmd] = $sCheminBase & $sCmd & ".pdf"
            If Not FileExists($aFichiers[$iNbCmd]) Then $bFichiersOK = False
            If $iNbCmd = 1 Then
                $sCmdListe &= $sCmd
            Else
                $sCmdListe &= ", " & $sCmd
            EndIf
        EndIf
    Next
    If $iNbCmd = 0 Or Not $bFichiersOK Then
        $sLogErr &= "Erreur (" & $sCmdListe & ") : PDF introuvable." & @CRLF
        Return False
    EndIf
    Local $oOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($oOutlook) Then Return False
    Local $oTemp = $oOutlook.CreateItem(0)
    $oTemp.GetInspector.Display
    Local $sSig = $oTemp.HTMLBody
    $oTemp.Close(1)
    Local $oMail = $oOutlook.CreateItem(0)
    $oMail.To = $sEmail
    $oMail.Subject = "Demande de rendez-vous pour livraison HPE - " & $sCmdListe
    Local $sPJ = "les Packing Lists"
    If $iNbCmd = 1 Then $sPJ = "la Packing List"
    Local $sBody = "Bonjour,<br><br>Nous vous contactons car nous avons de la marchandise de la part de HPE à vous livrer. Vous trouverez " & $sPJ & " en PJ.<br><br>" & _
                   "Nous souhaiterions savoir quand est-ce qu'une livraison vous arrangerait, avec les horaires d'ouvertures s'il vous plaît ? Si la demande de rendez-vous et/ou la réponse est avant 14h alors le délai est 48h ouvrés pour la livraison, sinon le délai passe à 72h ouvrés.<br><br>" & _
                   "Veuillez aussi nous communiquer un numéro de téléphone pour que nous puissions le transmettre à notre service de livraison, qu'il puisse contacter une personne sur place le jour de la livraison.<br><br>Merci d'avance."
    $oMail.HTMLBody = "<div style='font-family: Aptos, Calibri, sans-serif; font-size: 14pt;'>" & $sBody & "</div>" & $sSig
    For $f = 1 To $iNbCmd
        $oMail.Attachments.Add($aFichiers[$f])
    Next
    $oMail.Display
    Return True
EndFunc

; ==============================================================================
; PRÉ-ALERTES (Colonne 4)
; ==============================================================================
Func _Batch_Mails_Alerte($sData)
    If $sData = "" Then Return
    Local $aJobs = StringSplit($sData, "|")
    Local $iOk = 0
    Local $iErr = 0
    Local $sLogErr = ""
    For $i = 1 To $aJobs[0]
        Local $aInfos = StringSplit($aJobs[$i], ";")
        If $aInfos[0] >= 3 Then
            Local $sCarrier = "Inconnu"
            If $aInfos[0] >= 4 Then $sCarrier = $aInfos[4]
            If _Mail_Alerte($aInfos[1], $aInfos[3], $aInfos[2], $sCarrier, $sLogErr) Then
                $iOk += 1
                Sleep(1500)
            Else
                $iErr += 1
            EndIf
        EndIf
    Next
    MsgBox(64+262144, "Bilan Pré-Alertes", $iOk & " mail(s)." & @CRLF & $iErr & " erreur(s)." & @CRLF & $sLogErr)
EndFunc

Func _Mail_Alerte($sTracking, $sClient, $sEmail, $sCarrier, ByRef $sLogErr)
    Local $sCheminBase = "F:\CDG\PRODUCT\TRANSCON\Shared\Clients\HPE\Pre-alertes\"
    Local $aCmds = StringSplit(StringRegExpReplace($sTracking, "[,;\s]*\+[,;\s]*|[,;]+", " "), " ", 2)
    Local $sCmdListe = ""
    Local $iNbCmd = 0
    Local $bFichiersOK = True
    Local $aFichiers[UBound($aCmds) + 1]
    For $c = 0 To UBound($aCmds) - 1
        Local $sCmd = StringStripWS($aCmds[$c], 3)
        If $sCmd <> "" Then
            $iNbCmd += 1
            $aFichiers[$iNbCmd] = $sCheminBase & $sCmd & ".pdf"
            If Not FileExists($aFichiers[$iNbCmd]) Then $bFichiersOK = False
            If $iNbCmd = 1 Then
                $sCmdListe &= $sCmd
            Else
                $sCmdListe &= ", " & $sCmd
            EndIf
        EndIf
    Next
    If $iNbCmd = 0 Or Not $bFichiersOK Then
        $sLogErr &= "Erreur (" & $sCmdListe & ") : PDF introuvable." & @CRLF
        Return False
    EndIf
    Local $bPastCutOff  = (Number(@HOUR) > 14) Or (Number(@HOUR) = 14 And Number(@MIN) >= 30)
    Local $iDays        = 2
    If StringInStr($sCarrier, "7") Or StringInStr($sCarrier, "Flex") Then $iDays = 1
    If $bPastCutOff Then $iDays += 1
    Local $sDateLivraison = _AddWorkingDays($iDays)
    Local $oOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($oOutlook) Then Return False
    Local $oTemp = $oOutlook.CreateItem(0)
    $oTemp.GetInspector.Display
    Local $sSig = $oTemp.HTMLBody
    $oTemp.Close(1)
    Local $oMail = $oOutlook.CreateItem(0)
    $oMail.To      = $sEmail
    $oMail.Subject = "Livraison HPE - Commande " & $sCmdListe
    Local $sBody   = "Bonjour,<br><br>Merci de noter que vous recevrez une livraison HPE d'ici le <b>" & $sDateLivraison & "</b>.<br><br>Vous trouverez la commande en pièce jointe.<br><br>Bonne journée."
    $oMail.HTMLBody = "<div style='font-family: Aptos, Calibri, sans-serif; font-size: 14pt;'>" & $sBody & "</div>" & $sSig
    For $f = 1 To $iNbCmd
        $oMail.Attachments.Add($aFichiers[$f])
    Next
    $oMail.Display
    Return True
EndFunc

; ==============================================================================
; CHANNEL PARTNERS
; ==============================================================================
Func _Batch_Mails_CP($sData)
    If $sData = "" Then Return
    Local $aJobs = StringSplit($sData, "|")
    Local $iOk = 0
    Local $iErr = 0
    Local $iSkip = 0
    Local $sLogErr = ""
    For $i = 1 To $aJobs[0]
        ; Auto-détection séparateur : tester ;  ~  et chr(167)=§ reçu en ANSI
        Local $aInfos = StringSplit($aJobs[$i], ";")
        If $aInfos[0] < 8 Then $aInfos = StringSplit($aJobs[$i], "~")
        If $aInfos[0] < 8 Then $aInfos = StringSplit($aJobs[$i], Chr(167))
        If $aInfos[0] >= 8 Then
            If _Mail_CP($aInfos[1],$aInfos[2],$aInfos[3],$aInfos[4],$aInfos[5],$aInfos[6],$aInfos[7],$aInfos[8],$sLogErr) Then
                $iOk += 1
                Sleep(1500)
            Else
                $iErr += 1
            EndIf
        Else
            $iSkip += 1
            $sLogErr &= "Ignoré (champs:" & $aInfos[0] & "): " & StringLeft($aJobs[$i], 50) & @CRLF
        EndIf
    Next
    MsgBox(64+262144, "Bilan CP", $iOk & " mail(s)." & @CRLF & $iErr & " erreur(s)." & @CRLF & $iSkip & " ignoré(s)." & @CRLF & $sLogErr)
EndFunc

Func _Mail_CP($sClient,$sCmds,$sPal,$sColis,$sPoids,$sConso,$sEmailTo,$sEmailCC, ByRef $sLogErr)
    Local $sCheminBase = "F:\CDG\PRODUCT\TRANSCON\Shared\Clients\HPE\Pre-alertes\"
    Local $aCmdsArr    = StringSplit(StringRegExpReplace($sCmds, "[,;\s]*\+[,;\s]*|[,;]+", " "), " ", 2)
    Local $sCmdListe = ""
    Local $iNbCmd = 0
    Local $bFichiersOK = True
    Local $aFichiers[UBound($aCmdsArr) + 1]
    For $c = 0 To UBound($aCmdsArr) - 1
        Local $sCmd = StringRegExpReplace(StringStripWS($aCmdsArr[$c], 3), "[^\w]", "")
        If $sCmd <> "" Then
            $iNbCmd += 1
            $aFichiers[$iNbCmd] = $sCheminBase & $sCmd & ".pdf"
            If Not FileExists($aFichiers[$iNbCmd]) Then $bFichiersOK = False
            If $iNbCmd = 1 Then
                $sCmdListe &= $sCmd
            Else
                $sCmdListe &= ", " & $sCmd
            EndIf
        EndIf
    Next
    ; Chercher le fichier consolidé : J_document Check List.pdf (prioritaire) ou J.pdf (fallback)
    Local $sConsoPath = ""
    $sConso = StringRegExpReplace($sConso, "[^\w]", "")
    If $sConso <> "" Then
        If FileExists($sCheminBase & $sConso & "_document Check List.pdf") Then
            $sConsoPath = $sCheminBase & $sConso & "_document Check List.pdf"
        ElseIf FileExists($sCheminBase & $sConso & ".pdf") Then
            $sConsoPath = $sCheminBase & $sConso & ".pdf"
        Else
            $bFichiersOK = False
        EndIf
    EndIf
    If $iNbCmd = 0 Or Not $bFichiersOK Then
        $sLogErr &= "Erreur (" & $sCmdListe & ") : PDF manquant." & @CRLF
        Return False
    EndIf
    Local $sDateLivraison = _AddWorkingDays(2)
    Local $oOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($oOutlook) Then Return False
    Local $oTemp = $oOutlook.CreateItem(0)
    $oTemp.GetInspector.Display
    Local $sSig = $oTemp.HTMLBody
    $oTemp.Close(1)
    Local $oMail = $oOutlook.CreateItem(0)
    If $sEmailTo <> "" Then $oMail.To = $sEmailTo
    If $sEmailCC <> "" Then $oMail.CC = $sEmailCC
    $oMail.Subject = "Livraison HPE - " & $sCmdListe
    Local $sTxtPal = "palettes"
    If Number($sPal) <= 1 Then $sTxtPal = "palette"
    Local $sBody   = "Bonjour,<br><br>Nous avons une nouvelle commande HPE à vous livrer, dont vous trouverez les détails ci-joints.<br><br>" & _
                     "<ul><li>Nombre de " & $sTxtPal & " : " & $sPal & "</li>" & _
                     "<li>Nombre total de colis : " & $sColis & "</li>" & _
                     "<li>Poids total : " & $sPoids & " kg</li></ul><br>" & _
                     "Pourriez-vous svp nous confirmer un RDV pour le <b>" & $sDateLivraison & "</b> ? Si la réponse est après 14h30 alors la livraison sera décalée au lendemain.<br><br>" & _
                     "A la livraison, merci de signer le document FCDN ci-joint. A défaut, de nous le retourner signé dans un délai de 24h.<br><br>Merci d'avance."
    $oMail.HTMLBody = "<div style='font-family: Aptos, Calibri, sans-serif; font-size: 14pt;'>" & $sBody & "</div>" & $sSig
    For $f = 1 To $iNbCmd
        $oMail.Attachments.Add($aFichiers[$f])
    Next
    If $sConsoPath <> "" Then $oMail.Attachments.Add($sConsoPath)
    $oMail.Display
    Return True
EndFunc

; ==============================================================================
; FILE CLOSING — HELPER : résolution Carrier ID selon transporteur
; ══════════════════════════════════════════════════════════════════
;  UPS        → pas de Carrier ID → séquence courte
;  EFDS/GROUSSARD → transp contient directement le numéro (ex: "13")
;  Autres     → extraire le numéro entre parenthèses (ex: "DGS (8)" → "8")
; ==============================================================================
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
EndFunc

; ==============================================================================
; FILE CLOSING UPS — séquence courte (pas de Carrier ID, juste DEF)
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
    _FC_SmartSleep(300)
    _FC_AuditCtrl($hWnd, $sFC_LOG, "LOG apres")
    ControlSend($hWnd, "", $sFC_LOG, "{F8}")
    _FC_SmartSleep(3000)
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
    _FC_SmartSleep(500)
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
        Sleep(300)
        ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", "")
        Sleep(150)
        ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", $sFC_EDS)
        Sleep(500)
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
    _FC_SmartSleep(300)
    ControlSetText($hMenu, "", "[CLASS:TEdit; INSTANCE:1]", "1")
    _FC_SmartSleep(300)
    ControlClick($hMenu, "", "[TEXT:OK]")
    WinWaitClose($hMenu, "", 5)
    _FC_SmartSleep(500)
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
    _FC_SmartSleep(300)
    ControlSetText($hInput1, "", "[CLASS:TEdit; INSTANCE:1]", $Num)
    _FC_SmartSleep(300)
    ControlClick($hInput1, "", "[TEXT:OK]")
    WinWaitClose($hInput1, "", 5)
    _FC_SmartSleep(3000) ; E.TMS charge
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
    _FC_SmartSleep(300)
    ControlSetText($hDef, "", "[CLASS:TEdit; INSTANCE:1]", "DEF")
    _FC_SmartSleep(300)
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
                Sleep(500)
                ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
            EndIf
            Local $iTimer = TimerInit()
            While Not WinExists($sFC_FILEOPEN)
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
            _FC_AuditTiming("Attente FileOpen", TimerDiff($iTimer))
            If Not WinExists($sFC_FILEOPEN) Then
                _FC_AuditLog("  FileOpen absent apres timeout")
                ContinueLoop
            EndIf
            WinActivate($sFC_FILEOPEN)
            WinWaitActive($sFC_FILEOPEN, "", 3)
            _FC_AuditWinState($sFC_FILEOPEN, "FileOpen")
            Sleep(300)
            ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", "")
            Sleep(150)
            ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", $sFC_EDS)
            Sleep(500)
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
                If $bFC_Stop Then
                    _FC_AuditLog("*** STOP Col " & $colNom & " ***")
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
                Sleep(150)
                If StringInStr($sTitre, "REASON") Then
                    _FC_AuditLog("  Col " & $colNom & " : REASON popup -> DE")
                    ControlSetText($hWin, "", "[CLASS:TEdit; INSTANCE:1]", "DE")
                    Sleep(150)
                    ControlClick($hWin, "", "[TEXT:OK]")
                    WinWaitClose($hWin)
                    Sleep(300)
                Else
                    ControlSetText($hWin, "", "[CLASS:TEdit; INSTANCE:1]", $sVal)
                    Sleep(150)
                    ControlClick($hWin, "", "[TEXT:OK]")
                    WinWaitClose($hWin)
                    $bColValidee = True
                    Sleep(300)
                EndIf
            WEnd
            _FC_AuditTiming("Col " & $colNom & " (val='" & StringLeft($sVal, 30) & "')", TimerDiff($tCol))
        Next
        Sleep(500)
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
Func _Tracker_Start($sTitle, $aList)
    $g_iTrackCount = UBound($aList)
    If $g_iTrackCount = 0 Then Return False
    ReDim $g_aTrackIDs[$g_iTrackCount]

    ; GUI toujours au premier plan (0x00000008 = WS_EX_TOPMOST) mais sans voler le focus
    $g_hTracker = GUICreate($sTitle, 320, 460, @DesktopWidth - 340, 50, -1, 0x00000008)
    GUISetBkColor(0x2D2D30, $g_hTracker)
    GUISetFont(9, 400, 0, "Segoe UI")

    ; Barre de progression
    $g_idTrackProg = GUICtrlCreateProgress(10, 10, 300, 18)

    ; Label info
    $g_idTrackLbl = GUICtrlCreateLabel("0 / " & $g_iTrackCount & " dossiers", 10, 33, 300, 18, 1)
    GUICtrlSetColor(-1, 0xFFFFFF)

    ; Info dossier en cours
    $g_idBatchInfo = GUICtrlCreateLabel("", 10, 52, 300, 18, 1)
    GUICtrlSetColor(-1, 0xFFCC00)

    ; ══ BOUTONS DE CONTRÔLE ══
    $g_idBtnPause = GUICtrlCreateButton("⏸ Pause", 10, 74, 72, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0xFF9900)

    $g_idBtnPlay = GUICtrlCreateButton("▶ Play", 86, 74, 72, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0x00CC55)
    GUICtrlSetState(-1, $GUI_DISABLE)

    $g_idBtnSkip = GUICtrlCreateButton("⏭ Passer", 162, 74, 78, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0x3399FF)

    $g_idBtnStop = GUICtrlCreateButton("⏹ Stop", 244, 74, 66, 30)
    GUICtrlSetFont(-1, 10, 700)
    GUICtrlSetBkColor(-1, 0xFF4444)

    ; Liste des dossiers
    $g_idTrackLV = GUICtrlCreateListView("Statut|Numero J", 10, 110, 300, 340)
    GUICtrlSetBkColor(-1, 0x1E1E1E)
    GUICtrlSetColor(-1, 0xCCCCCC)
    GUICtrlSendMsg($g_idTrackLV, 0x101E, 0, 80)
    GUICtrlSendMsg($g_idTrackLV, 0x101E, 1, 190)
    For $i = 0 To $g_iTrackCount - 1
        $g_aTrackIDs[$i] = GUICtrlCreateListViewItem("Attente|" & $aList[$i], $g_idTrackLV)
        GUICtrlSetColor($g_aTrackIDs[$i], 0xAAAAAA)
    Next

    GUISetState(@SW_SHOWNOACTIVATE, $g_hTracker)
    Return True
EndFunc

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
Func _Tracker_PollButtons()
    If $g_hTracker = 0 Then Return
    Local $iMsg = GUIGetMsg()
    Switch $iMsg
        Case $g_idBtnPause
            $bFC_Pause = True
            $bCOMAT_Pause = True
            GUICtrlSetState($g_idBtnPause, $GUI_DISABLE)
            GUICtrlSetState($g_idBtnPlay, $GUI_ENABLE)
            GUICtrlSetData($g_idBatchInfo, "⏸ EN PAUSE — Cliquez Play pour reprendre")
        Case $g_idBtnPlay
            $bFC_Pause = False
            $bCOMAT_Pause = False
            GUICtrlSetState($g_idBtnPlay, $GUI_DISABLE)
            GUICtrlSetState($g_idBtnPause, $GUI_ENABLE)
        Case $g_idBtnSkip
            $bFC_Skip = True
            $bCOMAT_Skip = True
            $bFC_Pause = False
            $bCOMAT_Pause = False
            GUICtrlSetState($g_idBtnPlay, $GUI_DISABLE)
            GUICtrlSetState($g_idBtnPause, $GUI_ENABLE)
        Case $g_idBtnStop
            $bFC_Stop = True
            $bCOMAT_Stop = True
            $bFC_Pause = False
            $bCOMAT_Pause = False
    EndSwitch
EndFunc

Func _Tracker_End()
    If $g_hTracker Then
        GUICtrlSetData($g_idTrackProg, 100)
        GUICtrlSetData($g_idTrackLbl, "Traitement terminé !")
        GUICtrlSetData($g_idBatchInfo, "✓ Tout est fait")
        GUICtrlSetState($g_idBtnPause, $GUI_DISABLE)
        GUICtrlSetState($g_idBtnPlay, $GUI_DISABLE)
        GUICtrlSetState($g_idBtnSkip, $GUI_DISABLE)
        GUICtrlSetState($g_idBtnStop, $GUI_DISABLE)
        Sleep(2000)
        GUIDelete($g_hTracker)
        $g_hTracker = 0
    EndIf
EndFunc

; ==============================================================================
; COMAT
; ==============================================================================
Func _Batch_COMAT($sData)
    If $sData = "" Then Return
    $bCOMAT_Stop = False
    $bCOMAT_Pause = False
    $bCOMAT_Skip = False
    HotKeySet("{F9}", "_HK_COMAT_PauseToggle")
    HotKeySet("{ESCAPE}", "_HK_COMAT_Stop")

    Local $aJobs = StringSplit($sData, "|")
    Local $aValid[$aJobs[0]]
    For $i = 1 To $aJobs[0]
        Local $aInfos = StringSplit($aJobs[$i], ";")
        $aValid[$i-1] = $aInfos[1]
    Next
    _Tracker_Start("COMAT en masse", $aValid)
    Local $iDone = 0, $iStopped = 0
    Local $sRemaining = ""
    For $i = 1 To $aJobs[0]
        Local $aDetails = StringSplit($aJobs[$i], ";")
        If $aDetails[0] >= 1 Then
            Local $sNumJ = $aDetails[1]
            $bCOMAT_Skip = False
            _Tracker_Update($i-1, 1)
            _COMAT_WaitIfPaused2()
            If $bCOMAT_Stop Then
                _Tracker_Update($i-1, 3)
                $iStopped = 1
                ; Collecter les dossiers restants (celui-ci + suivants)
                For $r = $i To $aJobs[0]
                    Local $aR = StringSplit($aJobs[$r], ";")
                    If $aR[0] >= 1 Then $sRemaining &= $aR[1] & @CRLF
                Next
                ExitLoop
            EndIf
            If $bCOMAT_Skip Then
                _Tracker_Update($i-1, 4)
                $bCOMAT_Skip = False
                ContinueLoop
            EndIf
            _Run_COMAT_Single($sNumJ)
            If $bCOMAT_Stop Then
                _Tracker_Update($i-1, 3)
                $iStopped = 1
                ; Collecter les dossiers restants (suivants seulement, celui-ci peut être partiel)
                For $r = $i + 1 To $aJobs[0]
                    Local $aR2 = StringSplit($aJobs[$r], ";")
                    If $aR2[0] >= 1 Then $sRemaining &= $aR2[1] & @CRLF
                Next
                ExitLoop
            EndIf
            If $bCOMAT_Skip Then
                _Tracker_Update($i-1, 4)
                $bCOMAT_Skip = False
            Else
                _Tracker_Update($i-1, 2)
                $iDone += 1
            EndIf
            _Tracker_PollButtons()
            _COMAT_SmartSleep(300)
        EndIf
    Next
    HotKeySet("{F9}")
    HotKeySet("{ESCAPE}")
    _Tracker_End()
    ; Bilan final
    If $iStopped And $sRemaining <> "" Then
        ; Copier les dossiers restants dans le presse-papier pour reprise facile
        ClipPut(StringStripWS($sRemaining, 2))
        MsgBox(48+262144, "COMAT — Arrêté", _
            $iDone & " dossier(s) traité(s) sur " & $aJobs[0] & "." & @CRLF & @CRLF & _
            "Dossiers restants (copiés dans le presse-papier) :" & @CRLF & $sRemaining)
    ElseIf $iStopped Then
        MsgBox(48+262144, "COMAT — Arrêté", $iDone & " dossier(s) traité(s) sur " & $aJobs[0] & ".")
    EndIf
    $bCOMAT_Stop = False
    $bCOMAT_Pause = False
    $bCOMAT_Skip = False
EndFunc

Func _Run_COMAT_Single($Num)
    $Num = StringStripWS($Num, 8)
    If $Num = "" Then Return
    Local $hWnd = WinGetHandle("[CLASS:TfmBrowser]")
    If $hWnd = 0 Or Not WinExists($hWnd) Then
        _NotifyError("COMAT", "Fenêtre E.TMS introuvable.")
        $bCOMAT_Stop = True
        Return
    EndIf
    ; Mode arrière-plan — pas de WinActivate, tout via ControlSend

    _COMAT_Spinner("COMAT [" & $Num & "] 1/5 - LOG J...")
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG " & $Num)
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    _COMAT_SmartSleep($COMAT_DELAY_LOAD)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 2/5 - F3...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    Send("{F3}")
    _COMAT_SmartSleep($COMAT_DELAY_L)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 3/5 - F5 x4...")
    Local $k
    For $k = 1 To 4
        Send("{F5}")
        _COMAT_SmartSleep($COMAT_DELAY_M)
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Next
    _COMAT_SmartSleep($COMAT_DELAY_L)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 4/5 - F1 + TAB + C...")
    Send("{F1}")
    _COMAT_SmartSleep($COMAT_DELAY_L)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    For $k = 1 To 6
        Send("{TAB}")
        _COMAT_SmartSleep($COMAT_DELAY_S)
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Next
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Send("C")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    For $k = 1 To 4
        Send("{F5}")
        If $k = 4 Then
            _COMAT_SmartSleep($COMAT_DELAY_L)
        Else
            _COMAT_SmartSleep($COMAT_DELAY_M)
        EndIf
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    Next
    _COMAT_SmartSleep(800)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 5/5 - Retour LOG...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    _COMAT_SmartSleep(2000)
    ToolTip("")
EndFunc

Func _Action_COMAT_Solo($Num)
    $Num = StringStripWS($Num, 8)
    If $Num = "" Then
    MsgBox(48+262144, "Erreur", "Aucun numéro de dossier.")
    Return
EndIf
    If MsgBox(1+32+262144, "COMAT Solo", "Lancer COMAT sur le dossier : " & $Num & " ?") = 2 Then Return
    _Run_COMAT_Single($Num)
    ToolTip("")
    MsgBox(64+262144, "COMAT", "Dossier " & $Num & " traité.")
EndFunc

Func _GUI_COMAT_Multi()
    Local $hComat    = GUICreate("COMAT MULTI", 310, 440, -1, -1, -1, 0x00000008)
    GUISetBkColor(0x2D2D30)
    GUISetFont(9, 400, 0, "Segoe UI")
    GUICtrlCreateLabel("Numéros J (un par ligne) :", 10, 10, 290, 18)
    GUICtrlSetColor(-1, 0xFFFFFF)
    Local $idEdit    = GUICtrlCreateEdit("", 10, 30, 290, 310, BitOR(0x0004, 0x0040, 0x2000))
    GUICtrlSetBkColor(-1, 0x1E1E1E)
    GUICtrlSetColor(-1, 0xFFFFFF)
    Local $idBtnRun  = GUICtrlCreateButton("> LANCER COMAT MULTI", 10, 352, 290, 45)
    GUICtrlSetFont(-1, 10, 800)
    GUICtrlSetColor(-1, 0x007ACC)
    Local $idBtnStop = GUICtrlCreateButton("STOP", 10, 405, 290, 22)
    GUICtrlSetFont(-1, 8, 400)
    GUICtrlSetColor(-1, 0xCC0000)
    GUISetState(@SW_SHOW, $hComat)
    Local $msg_gui  = 0
    Local $sDataGui = ""
    Local $aLinesGui[1]
    Local $aValidGui[100]
    Local $iTotalGui = 0
    Local $sNumGui   = ""
    While 1
        $msg_gui = GUIGetMsg()
        Switch $msg_gui
            Case -3
                GUIDelete($hComat)
                Return
            Case $idBtnStop
                $bCOMAT_Stop = True
                GUIDelete($hComat)
                Return
            Case $idBtnRun
                $sDataGui = GUICtrlRead($idEdit)
                GUIDelete($hComat)
                If StringStripWS($sDataGui, 8) = "" Then
    MsgBox(48+262144, "Vide", "La liste est vide !")
    Return
EndIf
                $aLinesGui = StringSplit(StringStripCR($sDataGui), @LF)
                $iTotalGui = 0
                ReDim $aValidGui[$aLinesGui[0] + 1]
                For $j = 1 To $aLinesGui[0]
                    $sNumGui = StringStripWS($aLinesGui[$j], 8)
                    If $sNumGui <> "" Then
                        $aValidGui[$iTotalGui] = $sNumGui
                        $iTotalGui += 1
                    EndIf
                Next
                If $iTotalGui = 0 Then
    MsgBox(48+262144, "Vide", "Aucun numéro valide.")
    Return
EndIf
                ReDim $aValidGui[$iTotalGui]
                If MsgBox(1+32+262144, "Confirmation", $iTotalGui & " dossier(s) à traiter. GO ?") = 2 Then Return
                _Tracker_Start("COMAT Multi - Suivi", $aValidGui)
                HotKeySet("{F9}", "_HK_COMAT_PauseToggle")
                HotKeySet("{ESCAPE}", "_HK_COMAT_Stop")
                $bCOMAT_Stop = False
                $bCOMAT_Pause = False
                $bCOMAT_Skip = False
                Local $iDoneG = 0, $iStoppedG = 0
                Local $sRemainingG = ""
                For $j = 0 To $iTotalGui - 1
                    $bCOMAT_Skip = False
                    _Tracker_Update($j, 1)
                    _COMAT_WaitIfPaused2()
                    If $bCOMAT_Stop Then
                        _Tracker_Update($j, 3)
                        $iStoppedG = 1
                        For $rr = $j To $iTotalGui - 1
                            $sRemainingG &= $aValidGui[$rr] & @CRLF
                        Next
                        ExitLoop
                    EndIf
                    If $bCOMAT_Skip Then
                        _Tracker_Update($j, 4)
                        $bCOMAT_Skip = False
                        ContinueLoop
                    EndIf
                    _Run_COMAT_Single($aValidGui[$j])
                    If $bCOMAT_Stop Then
                        _Tracker_Update($j, 3)
                        $iStoppedG = 1
                        For $rr = $j + 1 To $iTotalGui - 1
                            $sRemainingG &= $aValidGui[$rr] & @CRLF
                        Next
                        ExitLoop
                    EndIf
                    If $bCOMAT_Skip Then
                        _Tracker_Update($j, 4)
                        $bCOMAT_Skip = False
                    Else
                        _Tracker_Update($j, 2)
                        $iDoneG += 1
                    EndIf
                    _Tracker_PollButtons()
                    _COMAT_SmartSleep(300)
                Next
                HotKeySet("{F9}")
                HotKeySet("{ESCAPE}")
                _Tracker_End()
                If $iStoppedG And $sRemainingG <> "" Then
                    ClipPut(StringStripWS($sRemainingG, 2))
                    MsgBox(48+262144, "COMAT — Arrêté", _
                        $iDoneG & " dossier(s) traité(s) sur " & $iTotalGui & "." & @CRLF & @CRLF & _
                        "Dossiers restants (copiés dans le presse-papier) :" & @CRLF & $sRemainingG)
                ElseIf $iStoppedG Then
                    MsgBox(48+262144, "COMAT — Arrêté", $iDoneG & " dossier(s) traité(s) sur " & $iTotalGui & ".")
                Else
                    MsgBox(64+262144, "Terminé", "Traitement COMAT terminé — " & $iDoneG & " dossier(s).")
                EndIf
                $bCOMAT_Stop = False
                $bCOMAT_Pause = False
                $bCOMAT_Skip = False
                Return
        EndSwitch
    WEnd
EndFunc

Func _COMAT_Spinner($sTxt)
    ToolTip($sTxt, 0, 0, "Robot E.TMS — COMAT", 1)
EndFunc

Func _COMAT_WaitIfPaused()
    While $bCOMAT_Pause And Not $bCOMAT_Stop And Not $bCOMAT_Skip
        _COMAT_Spinner("EN PAUSE — cliquez Play dans la fenêtre de contrôle")
        _Tracker_PollButtons()
        Sleep(80)
    WEnd
EndFunc

; Sleep réactif : découpe le sleep en tranches de 100ms et poll les boutons GUI
; Permet Pause/Skip/Stop instantané même pendant les longues attentes
Func _COMAT_SmartSleep($iMs)
    Local $iSlept = 0
    While $iSlept < $iMs
        _Tracker_PollButtons()
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
        _COMAT_WaitIfPaused()
        If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
        Sleep(100)
        $iSlept += 100
    WEnd
EndFunc

; ==============================================================================
; URI DECODE (pour les query params)
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
Func _GetPJConfig()
    Local $sIni  = @ScriptDir & "\dispatch_config.ini"
    Local $sPath = IniRead($sIni, "PJ", "Path",       "")
    Local $sRDV  = IniRead($sIni, "PJ", "RDV_Ext",    "pdf")
    Local $sUPS  = IniRead($sIni, "PJ", "UPS_Folder", "UPS")
    Local $sDGS  = IniRead($sIni, "PJ", "DGS_Folder", "DGS")
    Return StringSplit($sPath & "|" & $sRDV & "|" & $sUPS & "|" & $sDGS, "|", 1)
EndFunc

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
    Local $iSock = TCPConnect("127.0.0.1", 8888)
    Local $nConnect = TimerDiff($tNet)
    If $iSock >= 0 Then
        $sLog &= "  Connexion 127.0.0.1:8888   : " & Round($nConnect, 0) & " ms | OK" & @CRLF
        TCPCloseSocket($iSock)
    Else
        $sLog &= "  Connexion 127.0.0.1:8888   : ECHEC (err=" & @error & ")" & @CRLF
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
Func _CleanContactsFiles()
    Local $aFiles[4] = [$g_sContactsFile, "", "", ""]
    ; Chercher aussi les chunks
    For $i = 0 To 9
        Local $sChkF = @ScriptDir & "\dispatch_contacts_" & $i & ".json"
        If FileExists($sChkF) And $i < 3 Then $aFiles[$i + 1] = $sChkF
    Next

    Local $iCleaned = 0
    For $f = 0 To 3
        If $aFiles[$f] = "" Then ContinueLoop
        If Not FileExists($aFiles[$f]) Then ContinueLoop

        Local $hRead = FileOpen($aFiles[$f], 256)
        If $hRead = -1 Then ContinueLoop
        Local $sContent = FileRead($hRead)
        FileClose($hRead)

        Local $sBefore = $sContent

        ; Corriger les apostrophes mal encodées
        $sContent = StringReplace($sContent, "â€™", "'")       ; UTF-8 mojibake
        $sContent = StringReplace($sContent, "&#039;", "'")     ; HTML entity
        $sContent = StringReplace($sContent, "&apos;", "'")     ; XML entity
        $sContent = StringReplace($sContent, "&#x27;", "'")     ; Hex HTML entity
        $sContent = StringReplace($sContent, "'", "'")          ; Smart quote left
        $sContent = StringReplace($sContent, "'", "'")          ; Smart quote right
        $sContent = StringReplace($sContent, "Ã©", "é")         ; UTF-8 mojibake é
        $sContent = StringReplace($sContent, "Ã¨", "è")         ; UTF-8 mojibake è
        $sContent = StringReplace($sContent, "Ãª", "ê")         ; UTF-8 mojibake ê
        $sContent = StringReplace($sContent, "Ã ", "à")         ; UTF-8 mojibake à
        $sContent = StringReplace($sContent, "Ã§", "ç")         ; UTF-8 mojibake ç
        $sContent = StringReplace($sContent, "Ã¢", "â")         ; UTF-8 mojibake â
        $sContent = StringReplace($sContent, "Ã®", "î")         ; UTF-8 mojibake î
        $sContent = StringReplace($sContent, "Ã¼", "ü")         ; UTF-8 mojibake ü
        $sContent = StringReplace($sContent, "Ã¶", "ö")         ; UTF-8 mojibake ö
        $sContent = StringReplace($sContent, "&amp;", "&")       ; Double-encoded &
        $sContent = StringReplace($sContent, "&lt;", "<")        ; HTML <
        $sContent = StringReplace($sContent, "&gt;", ">")        ; HTML >
        $sContent = StringReplace($sContent, "&quot;", '"')      ; HTML "
        ; Supprimer les espaces multiples consécutifs
        While StringInStr($sContent, "  ")
            $sContent = StringReplace($sContent, "  ", " ")
        WEnd

        If $sContent <> $sBefore Then
            Local $hWrite = FileOpen($aFiles[$f], 2 + 256)
            FileWrite($hWrite, $sContent)
            FileClose($hWrite)
            $iCleaned += 1
        EndIf
    Next

    If $iCleaned > 0 Then
        TrayTip("Dispatch — Contacts", $iCleaned & " fichier(s) contact nettoyé(s).", 5, 1)
    Else
        TrayTip("Dispatch — Contacts", "Aucun nettoyage nécessaire.", 3, 1)
    EndIf
EndFunc

; ==============================================================================
; ESTIMATION STOCKAGE — Taille de tous les fichiers JSON
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
Func _NotifyError($sSource, $sMsg)
    TrayTip("Dispatch — Erreur " & $sSource, $sMsg, 10, 3)
    _AuditLog("ERREUR", $sSource & " : " & $sMsg)
EndFunc

; ==============================================================================
; AUDIT LOG — Journal d'erreurs en arrière-plan
; ==============================================================================
Func _AuditLog($sLevel, $sMsg)
    Local $sTs = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
    Local $sLine = "[" & $sTs & "] [" & $sLevel & "] " & $sMsg & @CRLF
    ConsoleWrite($sLine)
    ; Écrire dans le fichier audit
    Local $hAudit = FileOpen($g_sAuditLog, 1 + 256) ; 1=append, 256=UTF-8
    If $hAudit <> -1 Then
        FileWrite($hAudit, $sLine)
        FileClose($hAudit)
    EndIf
EndFunc

; ==============================================================================
; HEALTH CHECK SILENCIEUX — Tourne en arrière-plan toutes les minutes
; Détecte les erreurs silencieuses et les log
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

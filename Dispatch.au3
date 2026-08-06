; ╔══════════════════════════════════════════════════════════════════════════╗
; ║  DispatchMaster 2.0 - Serveur Local & Hub AutoIt                         ║
; ╚══════════════════════════════════════════════════════════════════════════╝
#include <File.au3>
#include <String.au3>
#include <Date.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <ListBoxConstants.au3>
#include <ComboConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>

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
Global Const $COMAT_DELAY_S    = 80
Global Const $COMAT_DELAY_M    = 150
Global Const $COMAT_DELAY_L    = 300
Global Const $COMAT_DELAY_LOAD = 1500

; ── CONSTANTES DE SÉCURITÉ (dispatch_patch) ──
Global Const $MAX_BODY_SIZE       = 2097152 ; 2 MB max par requête
Global Const $MAX_PATH_LENGTH     = 500     ; Longueur max d'un chemin réseau
Global Const $ALLOWED_NET_PREFIX  = "F:\"   ; Seul préfixe autorisé pour les chemins réseau
Global Const $ID_PATTERN          = "^[a-zA-Z0-9_\-\.\+\s]{1,200}$" ; Pattern valide pour un ID

; ── Checksum réseau (pour /api/net-check) ──
Global $g_sLastNetChecksum = ""
Global $g_sLastNetModified = ""

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
Global $g_sContactsFile       = @ScriptDir & "\dispatch_contacts.tsv"
Global $g_sContactsLegacyFile = @ScriptDir & "\dispatch_contacts.json"
Global $g_sHtmlFile = @ScriptDir & "\interface.html"
; CMR_FULL_FUSION_V1 - moteur CMR fusionné dans le serveur Dispatch
Global $g_sCMRStatus = "Prêt."
Global $g_bCMRRunning = False
Global $g_sCMRLastJSON = "{}"
Global $g_sCMRQueueFile = @ScriptDir & "\cmr_queue.txt"
Global $g_bCMRStop = False
Global $g_bCMRForce = False

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

    ; ── CONTACTS TSV : sauvegarde stable depuis Interface.html ──
    ElseIf $sURL = "/api/save-contacts" Then
        _BackupRotate($g_sContactsFile, 10)
        Local $hFileC = FileOpen($g_sContactsFile, 2 + 256)
        If $hFileC = -1 Then
            _SendHttpResponse($iSocket, 500, "application/json", '{"status":"error","reason":"cannot_write_contacts_tsv"}')
        Else
            FileWrite($hFileC, $sBody)
            FileClose($hFileC)
            _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","format":"tsv"}')
        EndIf

    ElseIf $sURL = "/api/load-contacts" Then
        Local $sContacts = ""
        If FileExists($g_sContactsFile) Then
            Local $hReadC = FileOpen($g_sContactsFile, 256)
            If $hReadC <> -1 Then
                $sContacts = FileRead($hReadC)
                FileClose($hReadC)
            EndIf
        ElseIf FileExists($g_sContactsLegacyFile) Then
            Local $hOldC = FileOpen($g_sContactsLegacyFile, 256)
            If $hOldC <> -1 Then
                $sContacts = FileRead($hOldC)
                FileClose($hOldC)
            EndIf
        EndIf
        If $sContacts = "" Then $sContacts = "#DISPATCH_CONTACTS_TSV" & @LF
        _SendHttpResponse($iSocket, 200, "text/plain", $sContacts)

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

    ElseIf StringLeft($sURL, 13) = "/api/net-list" Then
        ; /api/net-list?pattern=F:\...\dispatch_state_*.json — liste les fichiers correspondants
        Local $sPattern = StringMid($sURL, 22) ; après "/api/net-list?pattern="
        $sPattern = _URIDecode($sPattern)
        Local $sDir = StringRegExpReplace($sPattern, "\\[^\\]*$", "")
        Local $sGlob = StringRegExpReplace($sPattern, "^.*\\", "")
        Local $sFileList = '["' ; on construit un array JSON
        Local $hSearch2 = FileFindFirstFile($sDir & "\" & $sGlob)
        Local $bFirst = True
        If $hSearch2 <> -1 Then
            While True
                Local $sFound = FileFindNextFile($hSearch2)
                If @error Then ExitLoop
                If Not $bFirst Then $sFileList &= ',"'
                $sFileList &= StringReplace($sDir & "\" & $sFound, "\", "\\") & '"'
                $bFirst = False
            WEnd
            FileClose($hSearch2)
        EndIf
        If $bFirst Then
            $sFileList = "[]"
        Else
            $sFileList &= "]"
        EndIf
        _SendHttpResponse($iSocket, 200, "application/json", $sFileList)

    ; ── BKD CONFIG : sauvegarde globale BKD depuis Interface.html ──
    ElseIf $sURL = "/api/save-bkd-config" Then
        Local $sIniBKD = @ScriptDir & "\dispatch_config.ini"
        Local $sModeBKD = _GetJsonValue($sBody, "mode")
        Local $sCutoffBKD = _GetJsonValue($sBody, "cutoff")
        Local $sCustomBKD = _GetJsonValue($sBody, "customDate")

        If $sModeBKD = "" Then $sModeBKD = "auto"
        If $sCutoffBKD = "" Then $sCutoffBKD = "14:30"

        Local $bOkBKD = True
        If IniWrite($sIniBKD, "BKD", "Mode", $sModeBKD) = 0 Then $bOkBKD = False
        If IniWrite($sIniBKD, "BKD", "Cutoff", $sCutoffBKD) = 0 Then $bOkBKD = False
        If IniWrite($sIniBKD, "BKD", "CustomDate", $sCustomBKD) = 0 Then $bOkBKD = False

        If $bOkBKD Then
            _AuditLog("SAVE", "BKD config - mode=" & $sModeBKD & " cutoff=" & $sCutoffBKD & " custom=" & $sCustomBKD)
            _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
        Else
            _AuditLog("ERROR", "BKD config - impossible d'ecrire dispatch_config.ini")
            _SendHttpResponse($iSocket, 500, "application/json", '{"status":"error","reason":"cannot_write_bkd_config"}')
        EndIf

    ; ── BKD CONFIG : chargement global BKD vers Interface.html ──
    ElseIf $sURL = "/api/load-bkd-config" Then
        Local $sIniBKD2 = @ScriptDir & "\dispatch_config.ini"
        Local $sModeBKD2 = IniRead($sIniBKD2, "BKD", "Mode", "auto")
        Local $sCutoffBKD2 = IniRead($sIniBKD2, "BKD", "Cutoff", "14:30")
        Local $sCustomBKD2 = IniRead($sIniBKD2, "BKD", "CustomDate", "")

        Local $sRespBKD = '{"mode":"' & _JsonEscape($sModeBKD2) & _
                          '","cutoff":"' & _JsonEscape($sCutoffBKD2) & _
                          '","customDate":"' & _JsonEscape($sCustomBKD2) & '"}'

        _SendHttpResponse($iSocket, 200, "application/json", $sRespBKD)

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
        Local $sIndexes_a = ""

        _AuditLog("ACTION", $sAction & " — " & StringLeft($sBody, 200))

        ; ══ RÉPONSE IMMÉDIATE — l'HTML est débloqué avant l'exécution ══
        ; Pour les actions longues (ETMS, COMAT, FC), on répond OK tout de suite
        ; puis on exécute l'action. Le HTML n'attend plus.
        If $sAction = "ETMS_CMD" Or $sAction = "MAIL_RDV" Or $sAction = "KANBAN_2" Or _
           $sAction = "KANBAN_4" Or $sAction = "KANBAN_5" Or $sAction = "KANBAN_6" Or _
           $sAction = "COMAT_MULTI" Or $sAction = "COMAT_SOLO" Or $sAction = "BATCH_CP" Or _
           $sAction = "CMR_GENERATE" Then
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


            Case "CMR_GENERATE"
                $sData_a = _GetJsonValue($sBody, "data")
                _CMR_STABLE_RunFromData($sData_a)
            Case "CMR_STATUS"
                _SendHttpResponse($iSocket, 200, "application/json", _CMR_StatusJSON())
                TCPCloseSocket($iSocket)
                Return
            Case "CMR_OPEN_PDF"
                $sPath_a = _GetJsonValue($sBody, "path")
                If $sPath_a <> "" And FileExists($sPath_a) Then ShellExecute($sPath_a)
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
                TCPCloseSocket($iSocket)
                Return
            Case "CMR_OPEN_HTML"
                $sPath_a = _GetJsonValue($sBody, "path")
                If $sPath_a <> "" And FileExists($sPath_a) Then ShellExecute($sPath_a)
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')
                TCPCloseSocket($iSocket)
                Return
            Case "CMR_MAIL"
                $sData_a = _GetJsonValue($sBody, "index")
                _CMR_MailByIndex(Number($sData_a))
            Case "CMR_EDOC"
                $sData_a = _GetJsonValue($sBody, "index")
                _CMR_EdocByIndex(Number($sData_a))

            Case "CMR_MAIL_MANY"
                $sIndexes_a = _GetJsonValue($sBody, "indexes")
                _CMR_STABLE_MailMany($sIndexes_a)
            Case "CMR_EDOC_MANY"
                $sIndexes_a = _GetJsonValue($sBody, "indexes")
                _CMR_STABLE_EdocMany($sIndexes_a)
            Case "CMR_EDOC_MAIL_MANY"
                $sIndexes_a = _GetJsonValue($sBody, "indexes")
                _CMR_STABLE_EdocMailMany($sIndexes_a)
            Case "CMR_EDIT"
                $sData_a = _GetJsonValue($sBody, "index")
                _CMR_EditByIndex(Number($sData_a))
            Case "CMR_REBUILD_PDF"
                $sData_a = _GetJsonValue($sBody, "index")
                _CMR_RebuildPdfByIndex(Number($sData_a))
            Case "CMR_FORCE"
                $g_bForceNext = True
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","forced":true}')
                TCPCloseSocket($iSocket)
                Return
            Case "CMR_STOP"
                $g_bStopQueue = True
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","stop":true}')
                TCPCloseSocket($iSocket)
                Return
            Case "EDOC_MASTER_OPEN"
                _EDOCMasterGUI()
            Case "EDOC_MASTER_STATUS"
                _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","tool":"EDOC Master Bot"}')
                TCPCloseSocket($iSocket)
                Return

            Case "EDOC_WEB_INIT"
                _SendHttpResponse($iSocket, 200, "application/json", _EDOC_WebInit())
                TCPCloseSocket($iSocket)
                Return
            Case "EDOC_WEB_SCAN"
                Local $sEdocRespScan = _EDOC_WebScan($sBody)
                _SendHttpResponse($iSocket, 200, "application/json", $sEdocRespScan)
                TCPCloseSocket($iSocket)
                Return
            Case "EDOC_WEB_UPLOAD"
                Local $sEdocRespUpload = _EDOC_WebUpload($sBody)
                _SendHttpResponse($iSocket, 200, "application/json", $sEdocRespUpload)
                TCPCloseSocket($iSocket)
                Return
            Case "EDOC_WEB_RULE_SAVE"
                Local $sEdocRespRule = _EDOC_WebRuleSave($sBody)
                _SendHttpResponse($iSocket, 200, "application/json", $sEdocRespRule)
                TCPCloseSocket($iSocket)
                Return
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

    ; ── Servir les fichiers JS ──
    ElseIf $sURL = "/DataManager.js" Or $sURL = "/validate.js" Or $sURL = "/merge.js" Or $sURL = "/migrate.js" Then
        Local $sJsFile = @ScriptDir & "\" & StringMid($sURL, 2) ; enlever le / initial
        If FileExists($sJsFile) Then
            Local $hJs = FileOpen($sJsFile, 256)
            If $hJs <> -1 Then
                Local $sJsContent = FileRead($hJs)
                FileClose($hJs)
                _SendHttpResponse($iSocket, 200, "application/javascript", $sJsContent)
            Else
                _SendHttpResponse($iSocket, 500, "text/plain", "Cannot read file")
            EndIf
        Else
            _SendHttpResponse($iSocket, 404, "text/plain", "File not found: " & $sURL)
        EndIf

    ; ── /api/save-patch — Sauvegarde différentielle ──
    ElseIf $sURL = "/api/save-patch" Then
        If StringLen($sBody) > $MAX_BODY_SIZE Then
            _SendHttpResponse($iSocket, 400, "application/json", '{"success":false,"reason":"body_too_large"}')
        Else
            Local $sPatchResp = _API_SavePatch($sBody)
            _SendHttpResponse($iSocket, 200, "application/json", $sPatchResp)
        EndIf

    ; ── /api/net-check — Vérification légère réseau ──
    ElseIf StringLeft($sURL, 14) = "/api/net-check" Then
        Local $sCheckPath = StringMid($sURL, 21) ; après "/api/net-check?path="
        $sCheckPath = _URIDecode($sCheckPath)
        Local $sCheckResp = _API_NetCheck($sCheckPath)
        _SendHttpResponse($iSocket, 200, "application/json", $sCheckResp)

    ; ── /api/job-status — Statut d'un job en cours ──
    ElseIf StringLeft($sURL, 15) = "/api/job-status" Then
        Local $sJobResp = _API_JobStatus()
        _SendHttpResponse($iSocket, 200, "application/json", $sJobResp)

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
Func _JsonEscape($s)
    $s = StringReplace($s, "\", "\\")
    $s = StringReplace($s, '"', '\"')
    $s = StringReplace($s, @CRLF, "\n")
    $s = StringReplace($s, @CR, "\n")
    $s = StringReplace($s, @LF, "\n")
    Return $s
EndFunc

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

    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)

    Local $sInst = _GetETMSInstance($hWnd)
    Local $sCtrl = "[CLASS:TEIEdit; INSTANCE:" & $sInst & "]"

    ; ══ F8 seul : envoyer F8 sans écrire de commande ══
    If $sBouton = "F8" Then
        ControlFocus($hWnd, "", $sCtrl)
        ControlSend($hWnd, "", $sCtrl, "{F8}")
        _AuditLog("ETMS", "BG: F8 (execute)")
        Return
    EndIf

    ; Préparer la commande
    Send("{PGUP}")
    Sleep(150)
    Local $sCommande = $sBouton & " " & $sNumDossier
    If $sBouton = "LOG X" Then $sCommande = "LOG X"

    ControlFocus($hWnd, "", $sCtrl)
    ControlSetText($hWnd, "", $sCtrl, $sCommande)
    Sleep(150)
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
    $bFC_Skip = False
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
            _COMAT_SmartSleep(150)
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
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)

    _COMAT_Spinner("COMAT [" & $Num & "] 1/5 - LOG J...")
    Send("{PGUP}")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlFocus($hWnd, "", $COMAT_LOG_CTRL)
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
    _COMAT_SmartSleep(400)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 5/5 - Retour LOG...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlFocus($hWnd, "", $COMAT_LOG_CTRL)
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG")
    _COMAT_SmartSleep($COMAT_DELAY_M)
    If $bCOMAT_Stop Or $bCOMAT_Skip Then Return
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    _COMAT_SmartSleep(1000)
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
; CONTACTS TSV - Archivage des anciens JSON/chunks contacts
; ==============================================================================
Func _ArchiveOldContactJsonFiles()
    Local $sBackDir = @ScriptDir & "\backups\contacts_json_old"
    If Not FileExists(@ScriptDir & "\backups") Then DirCreate(@ScriptDir & "\backups")
    If Not FileExists($sBackDir) Then DirCreate($sBackDir)
    Local $iMoved = 0
    Local $aOld[3] = [@ScriptDir & "\dispatch_contacts.json", @ScriptDir & "\dispatch_contacts_meta.json", @ScriptDir & "\historique_contacts.json"]
    For $i = 0 To UBound($aOld) - 1
        If FileExists($aOld[$i]) Then
            Local $sName = StringRegExpReplace($aOld[$i], ".*\\", "")
            FileMove($aOld[$i], $sBackDir & "\" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_" & $sName, 9)
            $iMoved += 1
        EndIf
    Next
    For $c = 0 To 999
        Local $sChk = @ScriptDir & "\dispatch_contacts_" & $c & ".json"
        If FileExists($sChk) Then
            FileMove($sChk, $sBackDir & "\" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_dispatch_contacts_" & $c & ".json", 9)
            $iMoved += 1
        EndIf
    Next
    TrayTip("Dispatch - Contacts", $iMoved & " ancien(s) fichier(s) JSON contact archive(s).", 5, 1)
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


; ══════════════════════════════════════════════════════════════════════════
; FONCTIONS dispatch_patch — Endpoints v2.1
; ══════════════════════════════════════════════════════════════════════════

; ── Validation et sanitisation des chemins réseau ──
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
Func _SanitizeString($s)
    Local $sResult = StringRegExpReplace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    Return $sResult
EndFunc

; ── /api/save-patch — Sauvegarde différentielle ──
Func _API_SavePatch($sBody)
    Local $sId = _GetJsonValue($sBody, "id")
    Local $sVersion = _GetJsonValue($sBody, "v")
    Local $sUpdatedAt = _GetJsonValue($sBody, "updatedAt")
    Local $sUpdatedBy = _GetJsonValue($sBody, "updatedBy")
    Local $sChanges = _GetJsonValue($sBody, "changes")

    If Not _ValidateId($sId) Then
        Return '{"success":false,"reason":"invalid_id"}'
    EndIf

    $sId = _SanitizeString($sId)
    $sUpdatedBy = _SanitizeString($sUpdatedBy)

    If Not FileExists($g_sDataFile) Then
        Return '{"success":false,"reason":"no_data_file"}'
    EndIf

    Local $hRead = FileOpen($g_sDataFile, 256)
    If $hRead = -1 Then
        Return '{"success":false,"reason":"cannot_read_file"}'
    EndIf
    Local $sData = FileRead($hRead)
    FileClose($hRead)

    Local $sSearchPattern = '"file"\s*:\s*"' & StringRegExpReplace($sId, '([\.\+\*\?\[\]\(\)\{\}\^\$\\])', '\\\1') & '"'
    Local $iPos = StringRegExp($sData, $sSearchPattern)

    If Not $iPos Then
        $sSearchPattern = '"id"\s*:\s*"' & StringRegExpReplace($sId, '([\.\+\*\?\[\]\(\)\{\}\^\$\\])', '\\\1') & '"'
        $iPos = StringRegExp($sData, $sSearchPattern)
    EndIf

    If Not $iPos Then
        Return '{"success":false,"reason":"dossier_not_found","id":"' & $sId & '"}'
    EndIf

    Local $iNewV = Number($sVersion)
    If $iNewV <= 0 Then $iNewV = 1

    If $sChanges <> "" And $sChanges <> "{}" Then
        Local $aKeys = StringRegExp($sChanges, '"([^"]+)"\s*:', 3)
        If IsArray($aKeys) Then
            For $k = 0 To UBound($aKeys) - 1
                Local $sKey = $aKeys[$k]
                Local $sVal = _GetJsonValue($sChanges, $sKey)
                $sVal = _SanitizeString($sVal)
                _AuditLog("PATCH", "Dossier=" & $sId & " Clé=" & $sKey & " Val=" & $sVal)
            Next
        EndIf
    EndIf

    _AuditLog("PATCH", "Patch reçu pour " & $sId & " v" & $iNewV & " par " & $sUpdatedBy)
    Return '{"success":true,"v":' & $iNewV & ',"id":"' & $sId & '"}'
EndFunc

; ── /api/net-check — Vérification légère réseau ──
Func _API_NetCheck($sPath)
    If Not _ValidateNetPath($sPath) Then
        Return '{"error":"invalid_path","changed":false}'
    EndIf

    Local $sDir = StringRegExpReplace($sPath, "\\[^\\]*$", "")
    Local $sGlob = StringRegExpReplace($sPath, "^.*\\", "")
    $sGlob = StringReplace($sGlob, ".json", "_*.json")

    Local $hSearch = FileFindFirstFile($sDir & "\" & $sGlob)
    If $hSearch = -1 Then
        Return '{"lastModified":"","checksum":"none","changed":false,"files":0}'
    EndIf

    Local $sCheckData = ""
    Local $iFiles = 0
    Local $sLatest = ""

    While True
        Local $sFound = FileFindNextFile($hSearch)
        If @error Then ExitLoop
        Local $sFullPath = $sDir & "\" & $sFound
        Local $sSize = FileGetSize($sFullPath)
        Local $sTime = FileGetTime($sFullPath, 0, 1)
        $sCheckData &= $sFound & ":" & $sSize & ":" & $sTime & "|"
        If $sTime > $sLatest Then $sLatest = $sTime
        $iFiles += 1
    WEnd
    FileClose($hSearch)

    Local $iSum = 0
    For $c = 1 To StringLen($sCheckData)
        $iSum += Asc(StringMid($sCheckData, $c, 1))
    Next
    Local $sChecksum = Hex(Mod($iSum, 16777216), 6)

    Local $bChanged = ($sChecksum <> $g_sLastNetChecksum)
    $g_sLastNetChecksum = $sChecksum

    Local $sISO = ""
    If StringLen($sLatest) >= 14 Then
        $sISO = StringLeft($sLatest, 4) & "-" & StringMid($sLatest, 5, 2) & "-" & StringMid($sLatest, 7, 2) & "T" & _
                StringMid($sLatest, 9, 2) & ":" & StringMid($sLatest, 11, 2) & ":" & StringMid($sLatest, 13, 2) & "Z"
    EndIf

    Return '{"lastModified":"' & $sISO & '","checksum":"' & $sChecksum & '","changed":' & _
           StringLower(String($bChanged)) & ',"files":' & $iFiles & '}'
EndFunc

; ── /api/job-status — Statut d'un job E.TMS en cours ──
Func _API_JobStatus()
    If Not IsDeclared("g_bJobRunning") Or Not $g_bJobRunning Then
        Return '{"jobId":"","status":"idle","progress":0,"total":0,"current":"","done":true}'
    EndIf

    Local $sStatus = "running"
    If IsDeclared("g_bJobPaused") And $g_bJobPaused Then $sStatus = "paused"

    Local $sJobId = ""
    If IsDeclared("g_sCurrentJobId") Then $sJobId = $g_sCurrentJobId

    Local $iProgress = 0
    If IsDeclared("g_iJobProgress") Then $iProgress = $g_iJobProgress

    Local $iTotal = 0
    If IsDeclared("g_iJobTotal") Then $iTotal = $g_iJobTotal

    Local $sCurrent = ""
    If IsDeclared("g_sJobCurrent") Then $sCurrent = _SanitizeString($g_sJobCurrent)

    Local $sType = ""
    If IsDeclared("g_sJobType") Then $sType = $g_sJobType

    Local $bDone = ($iProgress >= $iTotal And $iTotal > 0)

    Return '{"jobId":"' & $sJobId & '","status":"' & $sStatus & '","progress":' & $iProgress & _
           ',"total":' & $iTotal & ',"current":"' & $sCurrent & '","type":"' & $sType & _
           '","done":' & StringLower(String($bDone)) & '}'
EndFunc

; ── Sécurité réseau — Remplacements sécurisés ──
Func _Net_SaveState_Secure($sPath, $sJSON)
    If Not _ValidateNetPath($sPath) Then
        _AuditLog("SECURITY", "Chemin réseau rejeté (save) : " & $sPath)
        Return False
    EndIf
    If StringLen($sJSON) > $MAX_BODY_SIZE Then
        _AuditLog("SECURITY", "Body trop large pour save : " & StringLen($sJSON) & " bytes")
        Return False
    EndIf
    Local $hFile = FileOpen($sPath, 2 + 256)
    If $hFile = -1 Then
        _NotifyError("Réseau", "Impossible d'écrire : " & $sPath)
        Return False
    EndIf
    FileWrite($hFile, $sJSON)
    FileClose($hFile)
    _AuditLog("NET", "Sauvegardé : " & $sPath & " (" & StringLen($sJSON) & " bytes)")
    Return True
EndFunc

Func _Net_LoadState_Secure($sPath)
    If Not _ValidateNetPath($sPath) Then
        _AuditLog("SECURITY", "Chemin réseau rejeté (load) : " & $sPath)
        Return "{}"
    EndIf
    If Not FileExists($sPath) Then Return "{}"
    Local $iSize = FileGetSize($sPath)
    If $iSize > $MAX_BODY_SIZE Then
        _AuditLog("SECURITY", "Fichier trop gros : " & $sPath & " (" & $iSize & " bytes)")
        Return "{}"
    EndIf
    Local $hFile = FileOpen($sPath, 256)
    If $hFile = -1 Then
        _NotifyError("Réseau", "Impossible de lire : " & $sPath)
        Return "{}"
    EndIf
    Local $sContent = FileRead($hFile)
    FileClose($hFile)
    _AuditLog("NET", "Chargé : " & $sPath & " (" & StringLen($sContent) & " bytes)")
    Return $sContent
EndFunc

; CMR_FULL_FUSION_MODULE_APPENDED


Opt("GUIOnEventMode", 0)

; ============================================================
; HPE BL QUEUE SNAPSHOT - VERSION PRO CLEAN OPTIMISÉE
; Base originale conservée, fonctions préservées.
; Optimisations majeures : cache CSV, throttling GUI hover, browser cache, snapshots plus rapides.
; ============================================================
; HPE BL QUEUE SNAPSHOT
; Correction majeure : chaque BL duplique ses outputs dans un dossier snapshot.
; Les HTML/PDF/mails sont générés depuis le snapshot du BL, jamais depuis OUTPUT live.
; ============================================================

Global Const $APP_TITLE = "HPE BL Queue Snapshot"
Global Const $BASE_PATH = "F:\Scripting\Export\EXPORT_HPE_CMR_001\"
Global Const $INPUT_PATH = $BASE_PATH & "Database\INPUT\"
Global Const $OUTPUT_PATH = $BASE_PATH & "Database\OUTPUT\"
Global Const $HTML_PATH = $BASE_PATH & "Database\HTML\"
Global Const $PDF_PATH = @ScriptDir & "\PDF\"
Global Const $SNAPSHOT_PATH = @ScriptDir & "\Snapshots\"
Global Const $INPUT_CSV = $INPUT_PATH & "input.csv"
Global Const $INPUT_GEN = $INPUT_PATH & "inputGEN.csv"
Global Const $LASTBL_MARKER = $INPUT_PATH & "last_bl_queue_snapshot.ini"
Global Const $EDS_PATH = $BASE_PATH & "EXPORT_HPE_CMR_006.eds"
Global Const $INI_PATH = @ScriptDir & "\Config\HPE_BL_QUEUE_SNAPSHOT.ini"
Global Const $MAIL_AUTOSEND = False
Global Const $OUTPUT_WAIT_MS = 180000
Global Const $OUTPUT_STABLE_MS = 1500

Global Const $ETMS_WINDOW = "[CLASS:TfmBrowser]"
Global Const $ETMS_TOOLBAR = "[CLASS:TRzToolbar; INSTANCE:1]"
Global Const $EDOC_UPLOAD_TITLE = "Upload Documents CDG"
Global Const $EDOC_UPLOAD_BTN = "TButton2"
Global Const $FILEOPEN_WIN = "[CLASS:TRzShellOpenSaveForm]"
Global Const $FILEOPEN_EDIT = "[CLASS:TRzEdit; INSTANCE:1]"
Global Const $TOOLBAR_X = 54
Global Const $TOOLBAR_Y = 9

Global $g_sOperator = ""
Global $g_bStopQueue = False
Global $g_aBL[0], $g_aHTML[0], $g_aPDF[0], $g_aCarrier[0], $g_aDelivery[0], $g_aCompany[0], $g_aSnap[0]
Global $g_aQueueBlocks[1] = [""]
Global $g_aQueueStatus[1] = [""]
Global $idQueue = 0, $idQueueList = 0, $idStatus = 0, $idResults = 0, $idLastBL = 0
Global $idSPCarrier = 0, $idSPTo = 0, $idSPCC = 0, $idSPBCC = 0, $idSPSubject = 0, $idSPPdf = 0, $idSPBody = 0, $idSPSign = 0, $idSPKnown = 0
; ============================================================
; UX / PROCESS CONTROL + HOVER BUTTONS
; ============================================================
Global $g_hGui = 0
Global $g_bGenerating = False
Global $g_bForceNext = False
Global $g_bForceSkipCurrent = False
Global $g_bHardClose = False
Global $idBtnStopQueue = 0
Global $idBtnForceNext = 0
Global $idBtnMail = 0
Global $idBtnEdocOnly = 0
Global $idBtnEdocMail = 0
Global $idBtnRefreshResults = 0
Global $idBtnRefreshResultsTop = 0
Global $g_aHoverIds[0], $g_aHoverNormal[0], $g_aHoverHover[0], $g_aHoverIsHover[0]

; ============================================================
; PRO CLEAN / PERFORMANCE CACHE
; ============================================================
Global $g_oCsvCache = ObjCreate("Scripting.Dictionary") ; clé = chemin fichier + date/taille, valeur = tableau de lignes non vides
Global $g_sBrowserCache = ""
Global $g_iLastHoverTick = 0
Global Const $HOVER_THROTTLE_MS = 60



Func _Main()
    DirCreate(@ScriptDir & "\Config")
    DirCreate($INPUT_PATH)
    DirCreate($OUTPUT_PATH)
    DirCreate($HTML_PATH)
    DirCreate($PDF_PATH)
    DirCreate($SNAPSHOT_PATH)
    DirCreate(@ScriptDir & "\HTML") ; fallback local PRO CLEAN
    _SP_InitDefaultsIfNeeded()

    $g_sOperator = IniRead($INI_PATH, "USER", "Operator", "")
    If StringStripWS($g_sOperator, 3) = "" Then
        $g_sOperator = InputBox($APP_TITLE, "Nom utilisateur / Operator", @UserName, "", 420, 140)
        If @error Or StringStripWS($g_sOperator, 3) = "" Then $g_sOperator = @UserName
        IniWrite($INI_PATH, "USER", "Operator", $g_sOperator)
    EndIf

    $g_hGui = GUICreate($APP_TITLE, 820, 610, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_MINIMIZEBOX))
    GUISetFont(9, 400, 0, "Segoe UI", $g_hGui)
    GUISetBkColor(0xF3F6FB, $g_hGui)
    Local $idTab = GUICtrlCreateTab(10, 10, 800, 590)

    GUICtrlCreateTabItem("Queue")
    Local $lblQueue = GUICtrlCreateLabel("Queue BL visuelle : 1 ligne dans la liste = 1 BL à générer", 24, 45, 520, 18)
    GUICtrlSetFont($lblQueue, 9, 700, 0, "Segoe UI")
    GUICtrlCreateLabel("Colle ici les J d’un BL, puis clique Ajouter BL. Pour plusieurs BL d’un coup : sépare par ligne vide ou --- puis Import auto.", 24, 66, 360, 34)
    $idQueue = GUICtrlCreateEdit("", 24, 103, 350, 95, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_WANTRETURN))
    Local $bAddQueue = GUICtrlCreateButton("Ajouter BL", 24, 204, 110, 28)
    Local $bImportQueue = GUICtrlCreateButton("Import auto", 144, 204, 110, 28)
    Local $bRemoveQueue = GUICtrlCreateButton("Retirer", 264, 204, 110, 28)
    $idQueueList = GUICtrlCreateListView("OK|#|Nb|J dans le BL|Statut", 24, 240, 350, 238, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_VSCROLL))
    _GUICtrlListView_SetExtendedListViewStyle($idQueueList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES))
    _GUICtrlListView_SetColumnWidth($idQueueList, 0, 34)
    _GUICtrlListView_SetColumnWidth($idQueueList, 1, 35)
    _GUICtrlListView_SetColumnWidth($idQueueList, 2, 35)
    _GUICtrlListView_SetColumnWidth($idQueueList, 3, 180)
    _GUICtrlListView_SetColumnWidth($idQueueList, 4, 70)
    Local $bPreviewQueue = GUICtrlCreateButton("Résumé", 24, 488, 110, 30)
    Local $bClearQueue = GUICtrlCreateButton("Vider queue", 144, 488, 110, 30)
    Local $bCheckQueue = GUICtrlCreateButton("Tout cocher", 264, 488, 110, 30)

    Local $lblWorkflow = GUICtrlCreateLabel("Workflow conseillé", 405, 45, 240, 18)
    GUICtrlSetFont($lblWorkflow, 9, 700, 0, "Segoe UI")
    Local $bGenerateAll = _CreateHoverButton("1. Générer tous les BL/PDF", 405, 72, 330, 44)
    $idBtnRefreshResultsTop = _CreateHoverButton("2. Rafraîchir résultats", 405, 126, 330, 34)
    $idBtnForceNext = _CreateHoverButton("Forcer output / passer au suivant", 405, 170, 330, 36)
    GUICtrlSetState($idBtnForceNext, $GUI_DISABLE)
    $idBtnStopQueue = _CreateHoverButton("Stop queue après BL en cours", 405, 214, 330, 34)
    Local $lblForceHint = GUICtrlCreateLabel("À utiliser si ETMS a fini mais que l’attente outputs reste bloquée.", 405, 250, 330, 16)
    GUICtrlSetFont($lblForceHint, 8, 400, 0, "Segoe UI")
    GUICtrlSetColor($lblForceHint, 0x64748B)
    GUICtrlSetBkColor($lblForceHint, 0xF3F6FB)
    Local $bUser = _CreateHoverButton("Utilisateur", 405, 270, 330, 30)
    $idLastBL = GUICtrlCreateLabel("Dernier BL : lecture...", 405, 310, 330, 42)
    GUICtrlSetColor($idLastBL, 0x344054)
    Local $lblStatus = GUICtrlCreateLabel("Résumé / progression", 405, 362, 240, 18)
    GUICtrlSetFont($lblStatus, 9, 700, 0, "Segoe UI")
    $idStatus = GUICtrlCreateEdit("Prêt.", 405, 384, 330, 103, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_READONLY))

    GUICtrlCreateTabItem("Résultats")
    Local $lblResults = GUICtrlCreateLabel("BL générés depuis snapshots", 24, 45, 260, 18)
    GUICtrlSetFont($lblResults, 9, 700, 0, "Segoe UI")
    $idResults = GUICtrlCreateListView("OK|BL|Entreprise|SP|Livraison", 24, 68, 485, 410, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $WS_VSCROLL))
    _GUICtrlListView_SetExtendedListViewStyle($idResults, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES))
    _GUICtrlListView_SetColumnWidth($idResults, 0, 38)
    _GUICtrlListView_SetColumnWidth($idResults, 1, 135)
    _GUICtrlListView_SetColumnWidth($idResults, 2, 145)
    _GUICtrlListView_SetColumnWidth($idResults, 3, 105)
    _GUICtrlListView_SetColumnWidth($idResults, 4, 95)
    GUICtrlCreateLabel("Sélection", 535, 45, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    $idBtnRefreshResults = _CreateHoverButton("Rafraîchir", 535, 68, 95, 28)
    Local $bClearResults = GUICtrlCreateButton("Vider", 640, 68, 95, 28)
    Local $bCheckResults = GUICtrlCreateButton("Tout cocher", 535, 102, 95, 28)
    Local $bUncheckResults = GUICtrlCreateButton("Décocher", 640, 102, 95, 28)

    GUICtrlCreateLabel("BL / PDF", 535, 145, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    Local $bEditBL = _CreateHoverButton("Modifier / enregistrer", 535, 168, 200, 34)
    Local $bOpenPdf = _CreateHoverButton("Ouvrir PDF", 535, 210, 200, 30)

    GUICtrlCreateLabel("Actions", 535, 262, 200, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    $idBtnMail = _CreateHoverButton("Créer mail", 535, 285, 200, 36)
    $idBtnEdocOnly = _CreateHoverButton("Upload EDOC seul", 535, 329, 200, 36)
    $idBtnEdocMail = _CreateHoverButton("EDOC + mail", 535, 373, 200, 40)
    Local $bSummaryResults = GUICtrlCreateButton("Résumé résultats", 535, 427, 200, 28)

    GUICtrlCreateLabel("Simple : coche les BL à traiter. Si rien n’est coché, l’action utilise la ligne sélectionnée. Modifier / enregistrer sauvegarde le HTML et recrée directement le PDF utilisé pour le mail/EDOC.", 24, 495, 710, 38)

    GUICtrlCreateTabItem("SP")
    GUICtrlCreateLabel("Transporteur / SP", 24, 45, 180, 18)
    $idSPCarrier = GUICtrlCreateCombo("", 24, 66, 300, 24, $CBS_DROPDOWN)
    GUICtrlSetData($idSPCarrier, _SP_ComboData(), "CARTAGE FLEX TRANSPORT")
    Local $bLoadSP = GUICtrlCreateButton("Charger", 340, 64, 85, 28)
    Local $bSaveSP = GUICtrlCreateButton("Sauver", 435, 64, 85, 28)
    Local $bDeleteSP = GUICtrlCreateButton("Supprimer", 530, 64, 95, 28)
    GUICtrlCreateLabel("To", 24, 102, 40, 18)
    $idSPTo = GUICtrlCreateInput("", 70, 99, 665, 22)
    GUICtrlCreateLabel("CC", 24, 129, 40, 18)
    $idSPCC = GUICtrlCreateInput("", 70, 126, 665, 22)
    GUICtrlCreateLabel("BCC", 24, 156, 40, 18)
    $idSPBCC = GUICtrlCreateInput("", 70, 153, 665, 22)
    GUICtrlCreateLabel("Objet", 24, 183, 45, 18)
    $idSPSubject = GUICtrlCreateInput("", 70, 180, 665, 22)
    GUICtrlCreateLabel("PDF", 24, 210, 45, 18)
    $idSPPdf = GUICtrlCreateInput("", 70, 207, 665, 22)
    GUICtrlCreateLabel("Corps mail - Entrée = saut de ligne", 24, 240, 420, 18)
    $idSPBody = GUICtrlCreateEdit("", 24, 262, 710, 105, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_WANTRETURN))
    GUICtrlCreateLabel("Signature additionnelle optionnelle", 24, 376, 560, 18)
    $idSPSign = GUICtrlCreateEdit("", 24, 398, 710, 70, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_WANTRETURN))
    GUICtrlCreateLabel("Variables : {J_SUBJECT}, {J_FILE}, {DELIVERY}, {DELIVERY_SHORT}, {DATE_JOUR}, {OPERATOR}, {CARRIER}, {NL}", 24, 478, 720, 18)
    $idSPKnown = GUICtrlCreateLabel("", 24, 502, 720, 20)
    GUICtrlCreateTabItem("")

    _SP_LoadToControls("CARTAGE FLEX TRANSPORT")
    _SP_UpdateKnownLabel()
    GUISetState(@SW_SHOW, $g_hGui)
    $g_iLastHoverTick = TimerInit()
    _LastBL_Refresh()

    While 1
        Local $msg = GUIGetMsg()
        _UpdateHoverButtons()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                If _HandleCloseRequest() Then ExitLoop
            Case $bAddQueue
                _QueueAddFromInput(False)
            Case $bImportQueue
                _QueueAddFromInput(True)
            Case $bRemoveQueue
                _QueueRemoveSelected()
            Case $bPreviewQueue
                _Status(_QueuePreviewText())
            Case $bClearQueue
                _QueueClearRows()
                GUICtrlSetData($idQueue, "")
                _Status("Queue vidée.")
            Case $bCheckQueue
                _QueueSetAllChecks(True)
            Case $bGenerateAll
                _RunGenerateQueue()
            Case $idBtnRefreshResultsTop, $idBtnRefreshResults
                _RefreshResultsList()
            Case $idBtnStopQueue
                $g_bStopQueue = True
                _Status("Stop demandé. Le BL en cours se termine puis la queue s’arrête.")
            Case $idBtnForceNext
                $g_bForceNext = True
                _Status("Forçage demandé : contrôle des outputs puis passage à l’étape suivante.")
            Case $bUser
                _ChangeUser()
            Case $bOpenPdf
                Local $idxOpen = _SelectedResultIndex()
                If $idxOpen >= 0 Then
                    _EnsureLatestPdfByIndex($idxOpen)
                    If FileExists($g_aPDF[$idxOpen]) Then ShellExecute($g_aPDF[$idxOpen])
                EndIf
            Case $bEditBL
                Local $idxEdit = _SelectedResultIndex()
                If $idxEdit >= 0 Then _OpenPreviewEditorByIndex($idxEdit)
            Case $idBtnMail
                _SendMailOnlySmart()
            Case $idBtnEdocOnly
                _UploadEdocOnlySmart()
            Case $idBtnEdocMail
                _SendEdocMailSmart()
            Case $bCheckResults
                _ResultsSetAllChecks(True)
            Case $bUncheckResults
                _ResultsSetAllChecks(False)
            Case $bSummaryResults
                _Status(_ResultsDetailedSummary())
            Case $bClearResults
                _ClearResults()
                _Status("Résultats vidés.")
            Case $bLoadSP
                _SP_LoadToControls(GUICtrlRead($idSPCarrier))
            Case $bSaveSP
                _SP_SaveFromControls()
                GUICtrlSetData($idSPCarrier, "")
                GUICtrlSetData($idSPCarrier, _SP_ComboData(), GUICtrlRead($idSPCarrier))
                _SP_UpdateKnownLabel()
                _Status("SP sauvegardé : " & GUICtrlRead($idSPCarrier))
            Case $bDeleteSP
                If MsgBox(52, $APP_TITLE, "Supprimer ce SP ?" & @CRLF & GUICtrlRead($idSPCarrier)) = 6 Then
                    _SP_Delete(GUICtrlRead($idSPCarrier))
                    GUICtrlSetData($idSPCarrier, "")
                    GUICtrlSetData($idSPCarrier, _SP_ComboData(), "Autre")
                    _SP_LoadToControls("Autre")
                    _SP_UpdateKnownLabel()
                EndIf
        EndSwitch
    WEnd
    GUIDelete($g_hGui)
EndFunc

Func _ChangeUser()
    Local $sNew = InputBox($APP_TITLE, "Nom utilisateur / Operator", $g_sOperator, "", 420, 140)
    If Not @error And StringStripWS($sNew, 3) <> "" Then
        $g_sOperator = StringStripWS($sNew, 3)
        IniWrite($INI_PATH, "USER", "Operator", $g_sOperator)
        _Status("Utilisateur sauvegardé : " & $g_sOperator)
    EndIf
EndFunc

Func _Status($txt)
    $g_sCMRStatus = $txt
    If $idStatus <> 0 Then GUICtrlSetData($idStatus, $txt)
    _AuditLog("CMR", StringLeft(StringReplace($txt, @CRLF, " | "), 500))
EndFunc

; ============================================================
; QUEUE + SNAPSHOT
; ============================================================
Func _RunGenerateQueue()
    Local $groups = _QueueGroups()
    If UBound($groups) = 0 Then
        _ModernInfo("Queue vide", "Colle au moins un J puis clique sur Ajouter BL ou Import auto.", "warning")
        Return False
    EndIf

    _ClearResults()
    $g_bStopQueue = False
    $g_bGenerating = True
    $g_bForceNext = False
    $g_bForceSkipCurrent = False
    $g_bHardClose = False
    _SetProcessingUi(True)

    For $i = 0 To UBound($groups) - 1
        If $g_bStopQueue Or $g_bHardClose Then ExitLoop
        Local $a = _QueueGroupToArray($groups[$i])
        If UBound($a) = 0 Then ContinueLoop

        $g_bForceNext = False
        $g_bForceSkipCurrent = False
        _QueueSetStatusByBlock($groups[$i], "En cours")
        _Status("Génération BL " & ($i + 1) & "/" & UBound($groups) & @CRLF & _JoinArray($a, " / "))

        Local $ok = _GenerateOneBL($a, $i + 1, UBound($groups))
        If $g_bHardClose Then ExitLoop

        If Not $ok Then
            If $g_bForceSkipCurrent Then
                _QueueSetStatusByBlock($groups[$i], "SKIP")
                _Status("BL ignoré par forçage utilisateur :" & @CRLF & _JoinArray($a, " / ") & @CRLF & @CRLF & "Passage au BL suivant.")
                _RefreshResultsList()
                ContinueLoop
            EndIf
            _QueueSetStatusByBlock($groups[$i], "KO")
            Local $r = _ModernErrorChoice("Erreur génération BL " & ($i + 1), _
                "BL concerné :" & @CRLF & _JoinArray($a, " / ") & @CRLF & @CRLF & "Que veux-tu faire ?", _
                "Continuer", "Relancer", "Arrêter")
            If $r = 3 Then ExitLoop
            If $r = 2 Then $i = $i - 1
        Else
            _QueueSetStatusByBlock($groups[$i], "OK")
        EndIf
        _RefreshResultsList()
    Next

    _SetProcessingUi(False)
    $g_bGenerating = False
    $g_bForceNext = False
    $g_bForceSkipCurrent = False

    If $g_bHardClose Then
        _Status("Fermeture demandée pendant traitement.")
        Return False
    EndIf

    _Status("Génération terminée." & @CRLF & _ResultsSummary() & @CRLF & "Chaque BL a son snapshot outputs.")
    Return True
EndFunc
Func _GenerateOneBL(ByRef $aNums, $index, $total)
    _ResetFiles()
    If Not _CreateInputFiles($aNums) Then Return False
    If Not _InputReady($aNums) Then Return False
    If Not _OpenEdsInEtms() Then Return False
    If Not _WaitOutputsReady($aNums, $index, $total) Then Return False

    Local $snap = _SnapshotOutputs($aNums, $index)
    If $snap = "" Then Return False

    Local $carrier = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 7)
    Local $delivery = _SnapGetDelivery($snap)
    Local $company = _SnapGetCompany($snap)
    Local $html = _GenerateHtmlFromSnapshot($aNums, $snap)
    If Not FileExists($html) Then Return False
    Local $pdf = _HtmlToPdf($html, $aNums, $carrier, $delivery)
    If Not FileExists($pdf) Then Return False
    _AddResult($aNums, $html, $pdf, $carrier, $delivery, $company, $snap)
    Return True
EndFunc

Func _WaitOutputsReady(ByRef $aNums, $index, $total)
    Local $t = TimerInit()
    While TimerDiff($t) < $OUTPUT_WAIT_MS
        _PumpGuiDuringProcess()
        If $g_bHardClose Then Return False
        If $g_bStopQueue Then Return False

        If $g_bForceNext Then
            $g_bForceNext = False
            If _OutputsOK() Then
                _Status("Forçage accepté." & @CRLF & "Les outputs existent : passage sans attendre la stabilité complète." & @CRLF & _JoinArray($aNums, " / "))
                Return True
            Else
                $g_bForceSkipCurrent = True
                _Status("Forçage demandé mais outputs incomplets." & @CRLF & "BL marqué SKIP puis passage au suivant :" & @CRLF & _JoinArray($aNums, " / "))
                Return False
            EndIf
        EndIf

        If _OutputsOK() And _OutputsStableFast() Then Return True
        _Status("BL " & $index & "/" & $total & " - attente outputs stables" & @CRLF & _JoinArray($aNums, " / ") & @CRLF & @CRLF & _
                "Si ETMS a terminé mais que l’outil reste bloqué, clique sur :" & @CRLF & "Forcer output / passer au suivant")
        Sleep(250)
    WEnd

    If _OutputsOK() Then
        _Status("Timeout atteint, mais outputs présents : passage à la suite.")
        Return True
    EndIf
    Return False
EndFunc
Func _OutputsStable()
    Local $s1 = _OutputsSignature()
    If $s1 = "" Then Return False
    Sleep($OUTPUT_STABLE_MS)
    Local $s2 = _OutputsSignature()
    Return ($s1 = $s2 And $s2 <> "")
EndFunc
Func _OutputsStableFast()
    Local $s1 = _OutputsSignature()
    If $s1 = "" Then Return False
    Local $t = TimerInit()
    While TimerDiff($t) < $OUTPUT_STABLE_MS
        _PumpGuiDuringProcess()
        If $g_bHardClose Or $g_bStopQueue Or $g_bForceNext Then Return False
        Sleep(100)
    WEnd
    Local $s2 = _OutputsSignature()
    Return ($s1 = $s2 And $s2 <> "")
EndFunc

Func _OutputsSignature()
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    Local $sig = ""
    For $i = 0 To 5
        Local $p = $OUTPUT_PATH & $a[$i]
        If Not FileExists($p) Or FileGetSize($p) <= 0 Then Return ""
        $sig &= $a[$i] & ":" & FileGetSize($p) & ":" & FileGetTime($p, 0, 1) & "|"
    Next
    Return $sig
EndFunc

Func _SnapshotOutputs(ByRef $aNums, $index)
    Local $name = @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_BL" & $index & "_" & _Safe(_JoinArray($aNums, "_"))
    Local $dir = $SNAPSHOT_PATH & $name & "\"
    DirCreate($dir)
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        Local $src = $OUTPUT_PATH & $a[$i]
        Local $dst = $dir & $a[$i]
        If Not FileExists($src) Or FileGetSize($src) <= 0 Then Return ""
        FileCopy($src, $dst, 1) ; PRO CLEAN : dossier déjà créé, overwrite simple plus rapide
        If Not FileExists($dst) Or FileGetSize($dst) <= 0 Then Return ""
    Next
    Return $dir
EndFunc

Func _AddResult(ByRef $aNums, $html, $pdf, $carrier, $delivery, $company, $snap)
    Local $n = UBound($g_aBL)
    ReDim $g_aBL[$n + 1]
    ReDim $g_aHTML[$n + 1]
    ReDim $g_aPDF[$n + 1]
    ReDim $g_aCarrier[$n + 1]
    ReDim $g_aDelivery[$n + 1]
    ReDim $g_aCompany[$n + 1]
    ReDim $g_aSnap[$n + 1]
    $g_aBL[$n] = _JoinArray($aNums, " / ")
    $g_aHTML[$n] = $html
    $g_aPDF[$n] = $pdf
    $g_aCarrier[$n] = $carrier
    $g_aDelivery[$n] = $delivery
    $g_aCompany[$n] = $company
    $g_aSnap[$n] = $snap
EndFunc

Func _ClearResults()
    ReDim $g_aBL[0]
    ReDim $g_aHTML[0]
    ReDim $g_aPDF[0]
    ReDim $g_aCarrier[0]
    ReDim $g_aDelivery[0]
    ReDim $g_aCompany[0]
    ReDim $g_aSnap[0]
    If $idResults <> 0 Then _GUICtrlListView_DeleteAllItems($idResults)
EndFunc

Func _CsvCacheClear($sPrefix = "")
    If Not IsObj($g_oCsvCache) Then Return False
    If $sPrefix = "" Then
        $g_oCsvCache.RemoveAll()
        Return True
    EndIf
    Local $aKeys = $g_oCsvCache.Keys
    For $i = 0 To UBound($aKeys) - 1
        If StringInStr($aKeys[$i], $sPrefix) = 1 Then $g_oCsvCache.Remove($aKeys[$i])
    Next
    Return True
EndFunc

Func _CsvLoadCached($file)
    If Not FileExists($file) Then Return SetError(1, 0, 0)
    Local $stamp = FileGetSize($file) & ":" & FileGetTime($file, 0, 1)
    Local $key = $file & "|" & $stamp
    If IsObj($g_oCsvCache) And $g_oCsvCache.Exists($key) Then Return $g_oCsvCache.Item($key)

    ; purge anciennes entrées du même fichier si le fichier a changé
    If IsObj($g_oCsvCache) Then
        Local $aKeys = $g_oCsvCache.Keys
        For $i = 0 To UBound($aKeys) - 1
            If StringInStr($aKeys[$i], $file & "|") = 1 Then $g_oCsvCache.Remove($aKeys[$i])
        Next
    EndIf

    Local $txt = FileRead($file)
    If @error Then Return SetError(2, 0, 0)
    $txt = StringReplace($txt, @CRLF, @LF)
    $txt = StringReplace($txt, @CR, @LF)
    Local $raw = StringSplit($txt, @LF, 2)
    Local $lines[0]
    For $i = 0 To UBound($raw) - 1
        If StringStripWS($raw[$i], 3) = "" Then ContinueLoop
        ReDim $lines[UBound($lines) + 1]
        $lines[UBound($lines) - 1] = $raw[$i]
    Next
    If IsObj($g_oCsvCache) Then $g_oCsvCache.Add($key, $lines)
    Return $lines
EndFunc

Func _RefreshResultsList()
    If $idResults = 0 Then Return False
    _GUICtrlListView_DeleteAllItems($idResults)
    For $i = 0 To UBound($g_aBL) - 1
        _GUICtrlListView_AddItem($idResults, "", $i)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aBL[$i], 1)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aCompany[$i], 2)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aCarrier[$i], 3)
        _GUICtrlListView_AddSubItem($idResults, $i, _ExtractDeliveryDateTime($g_aDelivery[$i]), 4)
        _GUICtrlListView_SetItemChecked($idResults, $i, False)
    Next
    Return True
EndFunc

Func _SelectedResultIndex()
    Local $aSel = _GUICtrlListView_GetSelectedIndices($idResults, True)
    If IsArray($aSel) And $aSel[0] > 0 Then
        Local $idx = Number($aSel[1])
        If $idx >= 0 And $idx < UBound($g_aBL) Then Return $idx
    EndIf
    MsgBox(48, $APP_TITLE, "Sélectionne une ligne dans l’onglet Résultats.")
    Return -1
EndFunc

Func _CheckedResultIndexes()
    Local $aIdx[0]
    For $i = 0 To UBound($g_aBL) - 1
        If _GUICtrlListView_GetItemChecked($idResults, $i) Then
            ReDim $aIdx[UBound($aIdx) + 1]
            $aIdx[UBound($aIdx) - 1] = $i
        EndIf
    Next
    Return $aIdx
EndFunc

Func _ActionIndexes()
    Local $aIdx = _CheckedResultIndexes()
    If UBound($aIdx) > 0 Then Return $aIdx
    Local $idx = _SelectedResultIndex()
    If $idx < 0 Then Return $aIdx
    ReDim $aIdx[1]
    $aIdx[0] = $idx
    Return $aIdx
EndFunc

Func _ResultsSummary()
    Return "BL générés : " & UBound($g_aBL)
EndFunc

Func _OpenPreviewEditorByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    _OpenPreviewEditor($g_aHTML[$idx], $a, $idx)
    Return True
EndFunc

Func _RebuildPdfByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _HtmlToPdf($g_aHTML[$idx], $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
    If FileExists($pdf) Then
        $g_aPDF[$idx] = $pdf
        ShellExecute($pdf)
        _Status("PDF recréé depuis HTML sélectionné :" & @CRLF & $pdf)
        Return True
    EndIf
    Return False
EndFunc

Func _CreateMailForResult($idx)
    If $idx < 0 Or $idx >= UBound($g_aPDF) Then Return ""
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _EnsureLatestPdfByIndex($idx)
    If FileExists($pdf) Then
        _CreateMailSP($pdf, $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
        _Status("Mail créé avec PDF :" & @CRLF & $g_aBL[$idx])
        Return $pdf
    EndIf
    MsgBox(48, $APP_TITLE, "PDF introuvable pour ce BL." & @CRLF & $g_aBL[$idx])
    Return ""
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

Func _SendMailOnlyResult($idx)
    Return (_CreateMailForResult($idx) <> "")
EndFunc

Func _SendOneResult($idx)
    ; Compatibilité : mail d'abord, puis EDOC pour ce BL.
    If _CreateMailForResult($idx) = "" Then Return False
    Return _UploadEdocForResult($idx)
EndFunc

Func _SendMailOnlySmart()
    Local $aIdx = _ActionIndexes()
    If UBound($aIdx) = 0 Then Return False
    For $i = 0 To UBound($aIdx) - 1
        _CreateMailForResult($aIdx[$i])
    Next
    _Status("Mails créés avec PDF : " & UBound($aIdx))
    Return True
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

Func _ResultsSetAllChecks($b)
    If $idResults = 0 Then Return False
    Local $count = _GUICtrlListView_GetItemCount($idResults)
    For $i = 0 To $count - 1
        _GUICtrlListView_SetItemChecked($idResults, $i, $b)
    Next
    Return True
EndFunc

Func _ResultsDetailedSummary()
    If UBound($g_aBL) = 0 Then Return "Aucun résultat BL."
    Local $s = "Résultats BL : " & UBound($g_aBL) & @CRLF & @CRLF
    For $i = 0 To UBound($g_aBL) - 1
        $s &= "#" & ($i + 1) & " - " & $g_aBL[$i] & @CRLF
        $s &= "Entreprise : " & $g_aCompany[$i] & @CRLF
        $s &= "SP : " & $g_aCarrier[$i] & @CRLF
        $s &= "Livraison : " & _ExtractDeliveryDateTime($g_aDelivery[$i]) & @CRLF
        $s &= "PDF : " & $g_aPDF[$i] & @CRLF & @CRLF
    Next
    Return $s
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

Func _SendAllResults()
    If UBound($g_aPDF) = 0 Then
        MsgBox(48, $APP_TITLE, "Aucun BL généré.")
        Return False
    EndIf
    For $i = 0 To UBound($g_aPDF) - 1
        _SendOneResult($i)
    Next
    _Status("EDOC + mails lancés pour tous les BL générés.")
    Return True
EndFunc

Func _StringToArrayBL($s)
    Local $parts = StringSplit(StringReplace($s, " / ", @LF), @LF, 2)
    Local $a[0]
    For $i = 0 To UBound($parts) - 1
        Local $j = _CleanJ($parts[$i])
        If $j = "" Then ContinueLoop
        ReDim $a[UBound($a) + 1]
        $a[UBound($a) - 1] = $j
    Next
    Return $a
EndFunc

Func _QueueAddFromInput($bImportMulti)
    Local $txt = GUICtrlRead($idQueue)
    If StringStripWS($txt, 3) = "" Then
        MsgBox(48, $APP_TITLE, "Colle au moins un J dans la zone de gauche.")
        Return False
    EndIf
    Local $groups
    If $bImportMulti Then
        $groups = _QueueGroupsFromText($txt)
    Else
        Local $a = _QueueGroupToArray($txt)
        If UBound($a) = 0 Then Return False
        Local $one[1]
        $one[0] = _JoinArray($a, @LF)
        $groups = $one
    EndIf
    For $i = 0 To UBound($groups) - 1
        _QueueAppendBlock($groups[$i])
    Next
    GUICtrlSetData($idQueue, "")
    _QueueRefreshList()
    _Status("BL ajoutés dans la queue : " & UBound($groups))
    Return True
EndFunc

Func _QueueCount()
    If Not IsArray($g_aQueueBlocks) Then Return 0
    If UBound($g_aQueueBlocks) = 1 And StringStripWS($g_aQueueBlocks[0], 3) = "" Then Return 0
    Return UBound($g_aQueueBlocks)
EndFunc

Func _QueueAppendBlock($sBlock)
    Local $a = _QueueGroupToArray($sBlock)
    If UBound($a) = 0 Then Return False

    Local $clean = _JoinArray($a, @LF)
    Local $n = _QueueCount()

    If $n = 0 Then
        ReDim $g_aQueueBlocks[1]
        ReDim $g_aQueueStatus[1]
        $g_aQueueBlocks[0] = $clean
        $g_aQueueStatus[0] = "Prêt"
    Else
        ReDim $g_aQueueBlocks[$n + 1]
        ReDim $g_aQueueStatus[$n + 1]
        $g_aQueueBlocks[$n] = $clean
        $g_aQueueStatus[$n] = "Prêt"
    EndIf

    Return True
EndFunc

Func _QueueRefreshList()
    If $idQueueList = 0 Then Return False
    _GUICtrlListView_DeleteAllItems($idQueueList)

    Local $count = _QueueCount()
    For $i = 0 To $count - 1
        Local $a = _QueueGroupToArray($g_aQueueBlocks[$i])
        _GUICtrlListView_AddItem($idQueueList, "", $i)
        _GUICtrlListView_AddSubItem($idQueueList, $i, $i + 1, 1)
        _GUICtrlListView_AddSubItem($idQueueList, $i, UBound($a), 2)
        _GUICtrlListView_AddSubItem($idQueueList, $i, _JoinArray($a, " / "), 3)
        _GUICtrlListView_AddSubItem($idQueueList, $i, $g_aQueueStatus[$i], 4)
        _GUICtrlListView_SetItemChecked($idQueueList, $i, True)
    Next
    Return True
EndFunc

Func _QueueRemoveSelected()
    Local $aSel = _GUICtrlListView_GetSelectedIndices($idQueueList, True)
    If Not IsArray($aSel) Or $aSel[0] = 0 Then
        MsgBox(48, $APP_TITLE, "Sélectionne une ligne dans la queue.")
        Return False
    EndIf

    Local $idx = Number($aSel[1])
    Local $count = _QueueCount()
    If $idx < 0 Or $idx >= $count Then Return False

    If $count <= 1 Then
        _QueueClearRows()
        Return True
    EndIf

    For $i = $idx To $count - 2
        $g_aQueueBlocks[$i] = $g_aQueueBlocks[$i + 1]
        $g_aQueueStatus[$i] = $g_aQueueStatus[$i + 1]
    Next

    ReDim $g_aQueueBlocks[$count - 1]
    ReDim $g_aQueueStatus[$count - 1]
    _QueueRefreshList()
    Return True
EndFunc

Func _QueueClearRows()
    ReDim $g_aQueueBlocks[1]
    ReDim $g_aQueueStatus[1]
    $g_aQueueBlocks[0] = ""
    $g_aQueueStatus[0] = ""
    If $idQueueList <> 0 Then _GUICtrlListView_DeleteAllItems($idQueueList)
EndFunc

Func _QueueSetAllChecks($b)
    If $idQueueList = 0 Then Return False
    Local $count = _GUICtrlListView_GetItemCount($idQueueList)
    For $i = 0 To $count - 1
        _GUICtrlListView_SetItemChecked($idQueueList, $i, $b)
    Next
    Return True
EndFunc

Func _QueueSetStatusByBlock($sBlock, $status)
    Local $aTarget = _QueueGroupToArray($sBlock)
    Local $target = _JoinArray($aTarget, @LF)
    Local $count = _QueueCount()

    For $i = 0 To $count - 1
        If $g_aQueueBlocks[$i] = $target Then
            $g_aQueueStatus[$i] = $status
            If $idQueueList <> 0 Then _GUICtrlListView_SetItemText($idQueueList, $i, $status, 4)
            Return True
        EndIf
    Next
    Return False
EndFunc

Func _QueueGroups()
    Local $count = _QueueCount()

    If $count > 0 Then
        Local $res[0]
        Local $hasChecked = False

        For $i = 0 To $count - 1
            If $idQueueList = 0 Or _GUICtrlListView_GetItemChecked($idQueueList, $i) Then
                $hasChecked = True
                ReDim $res[UBound($res) + 1]
                $res[UBound($res) - 1] = $g_aQueueBlocks[$i]
            EndIf
        Next

        ; Si aucune ligne cochee, on prend quand meme toute la queue pour eviter un blocage utilisateur.
        If Not $hasChecked Then
            ReDim $res[$count]
            For $i = 0 To $count - 1
                $res[$i] = $g_aQueueBlocks[$i]
            Next
        EndIf

        Return $res
    EndIf

    Local $txt = GUICtrlRead($idQueue)
    Return _QueueGroupsFromText($txt)
EndFunc

Func _QueueGroupsFromText($txt)
    $txt = StringReplace($txt, @CRLF, @LF)
    $txt = StringReplace($txt, @CR, @LF)
    Local $raw = StringSplit($txt, @LF, 2)
    Local $groups[0]
    Local $cur = ""
    For $i = 0 To UBound($raw) - 1
        Local $line = StringStripWS($raw[$i], 3)
        If $line = "" Or $line = "---" Then
            If StringStripWS($cur, 3) <> "" Then
                ReDim $groups[UBound($groups) + 1]
                $groups[UBound($groups) - 1] = $cur
                $cur = ""
            EndIf
        Else
            If $cur <> "" Then $cur &= @LF
            $cur &= $line
        EndIf
    Next
    If StringStripWS($cur, 3) <> "" Then
        ReDim $groups[UBound($groups) + 1]
        $groups[UBound($groups) - 1] = $cur
    EndIf
    Return $groups
EndFunc

Func _QueueGroupToArray($s)
    Local $norm = StringReplace(StringReplace($s, " / ", @LF), ";", @LF)
    Local $raw = StringSplit($norm, @LF, 2)
    Local $a[0]
    For $i = 0 To UBound($raw) - 1
        Local $j = _CleanJ($raw[$i])
        If $j = "" Then ContinueLoop
        ReDim $a[UBound($a) + 1]
        $a[UBound($a) - 1] = $j
    Next
    Return $a
EndFunc

Func _QueuePreviewText()
    Local $g = _QueueGroups()
    Local $s = ""
    If UBound($g) = 0 Then Return "Queue vide. Colle des J puis clique Ajouter BL ou Import auto."
    For $i = 0 To UBound($g) - 1
        Local $a = _QueueGroupToArray($g[$i])
        $s &= "BL " & ($i + 1) & " : " & _JoinArray($a, " / ") & @CRLF
    Next
    Return $s
EndFunc

; ============================================================
; SP
; ============================================================
Func _Enc($s)
    $s = StringReplace($s, @CRLF, "{BR}")
    $s = StringReplace($s, @CR, "{BR}")
    $s = StringReplace($s, @LF, "{BR}")
    Return $s
EndFunc

Func _Dec($s)
    Return StringReplace($s, "{BR}", @CRLF)
EndFunc

Func _SP_InitDefaultsIfNeeded()
    If IniRead($INI_PATH, "SP:CARTAGE FLEX TRANSPORT", "TO", "") <> "" Then Return
    _SP_WriteRule("CARTAGE FLEX TRANSPORT", "sensible@flex-transport.com", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Hello{NL}{NL}Pour le {DELIVERY_SHORT}", "Regards,{NL}{NL}{OPERATOR}")
    _SP_WriteRule("European Fret Distribution Services", "exploitation95@efds.eu", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance.", "Cordialement,{NL}{NL}{OPERATOR}")
    _SP_WriteRule("Transports Groussard SA", "expeditors@tps-groussard.fr", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance.", "Cordialement,{NL}{NL}{OPERATOR}")
    _SP_WriteRule("Autre", "", "", "", "{J_SUBJECT}", "{J_FILE}_Delivery Order", "Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance.", "Cordialement,{NL}{NL}{OPERATOR}")
EndFunc

Func _SP_Sec($carrier)
    Return "SP:" & StringUpper(StringStripWS($carrier, 3))
EndFunc

Func _SP_WriteRule($carrier, $to, $cc, $bcc, $subject, $pdf, $body, $signature)
    Local $sec = _SP_Sec($carrier)
    IniWrite($INI_PATH, $sec, "CARRIER", $carrier)
    IniWrite($INI_PATH, $sec, "TO", $to)
    IniWrite($INI_PATH, $sec, "CC", $cc)
    IniWrite($INI_PATH, $sec, "BCC", $bcc)
    IniWrite($INI_PATH, $sec, "SUBJECT", $subject)
    IniWrite($INI_PATH, $sec, "PDF", $pdf)
    IniWrite($INI_PATH, $sec, "BODY", _Enc($body))
    IniWrite($INI_PATH, $sec, "SIGNATURE", _Enc($signature))
EndFunc

Func _SP_ComboData()
    Local $sections = IniReadSectionNames($INI_PATH)
    If @error Then Return "CARTAGE FLEX TRANSPORT|European Fret Distribution Services|Transports Groussard SA|Autre"
    Local $s = ""
    For $i = 1 To $sections[0]
        If StringLeft($sections[$i], 3) = "SP:" Then
            Local $c = IniRead($INI_PATH, $sections[$i], "CARRIER", StringTrimLeft($sections[$i], 3))
            If $s <> "" Then $s &= "|"
            $s &= $c
        EndIf
    Next
    Return $s
EndFunc

Func _SP_LoadToControls($carrier)
    Local $sec = _SP_Sec($carrier)
    GUICtrlSetData($idSPCarrier, IniRead($INI_PATH, $sec, "CARRIER", $carrier))
    GUICtrlSetData($idSPTo, IniRead($INI_PATH, $sec, "TO", ""))
    GUICtrlSetData($idSPCC, IniRead($INI_PATH, $sec, "CC", ""))
    GUICtrlSetData($idSPBCC, IniRead($INI_PATH, $sec, "BCC", ""))
    GUICtrlSetData($idSPSubject, IniRead($INI_PATH, $sec, "SUBJECT", "{J_SUBJECT}"))
    GUICtrlSetData($idSPPdf, IniRead($INI_PATH, $sec, "PDF", "{J_FILE}_Delivery Order"))
    GUICtrlSetData($idSPBody, _Dec(IniRead($INI_PATH, $sec, "BODY", _Enc("Bonjour,{NL}{NL}Merci de mettre en place ce transport pour une livraison le {DELIVERY_SHORT} (En PJ le BL){NL}{NL}Merci d’avance."))))
    GUICtrlSetData($idSPSign, _Dec(IniRead($INI_PATH, $sec, "SIGNATURE", _Enc("Cordialement,{NL}{NL}{OPERATOR}"))))
EndFunc

Func _SP_SaveFromControls()
    Local $carrier = StringStripWS(GUICtrlRead($idSPCarrier), 3)
    If $carrier = "" Then Return False
    _SP_WriteRule($carrier, GUICtrlRead($idSPTo), GUICtrlRead($idSPCC), GUICtrlRead($idSPBCC), GUICtrlRead($idSPSubject), GUICtrlRead($idSPPdf), GUICtrlRead($idSPBody), GUICtrlRead($idSPSign))
    Return True
EndFunc

Func _SP_Delete($carrier)
    IniDelete($INI_PATH, _SP_Sec($carrier))
EndFunc

Func _SP_UpdateKnownLabel()
    GUICtrlSetData($idSPKnown, "SP connus : " & StringReplace(_SP_ComboData(), "|", " / "))
EndFunc

Func _SP_ReadRule($carrier)
    Local $sec = _SP_MatchSection($carrier)
    Local $r = ObjCreate("Scripting.Dictionary")
    $r.Add("CARRIER", IniRead($INI_PATH, $sec, "CARRIER", $carrier))
    $r.Add("TO", IniRead($INI_PATH, $sec, "TO", ""))
    $r.Add("CC", IniRead($INI_PATH, $sec, "CC", ""))
    $r.Add("BCC", IniRead($INI_PATH, $sec, "BCC", ""))
    $r.Add("SUBJECT", IniRead($INI_PATH, $sec, "SUBJECT", "{J_SUBJECT}"))
    $r.Add("PDF", IniRead($INI_PATH, $sec, "PDF", "{J_FILE}_Delivery Order"))
    $r.Add("BODY", _Dec(IniRead($INI_PATH, $sec, "BODY", "")))
    $r.Add("SIGNATURE", _Dec(IniRead($INI_PATH, $sec, "SIGNATURE", "")))
    Return $r
EndFunc

Func _SP_MatchSection($carrier)
    Local $c = StringUpper(StringStripWS($carrier, 3))
    Local $sections = IniReadSectionNames($INI_PATH)
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 3) = "SP:" Then
                Local $ruleCarrier = IniRead($INI_PATH, $sections[$i], "CARRIER", "")
                If $ruleCarrier <> "" And StringInStr($c, StringUpper($ruleCarrier)) Then Return $sections[$i]
            EndIf
        Next
    EndIf
    Return "SP:AUTRE"
EndFunc

Func _SP_GetRule($carrier)
    Return _SP_ReadRule($carrier)
EndFunc

Func _SP_BuildVariables(ByRef $aNums, $delivery, $carrier)
    Local $v = ObjCreate("Scripting.Dictionary")
    $v.Add("J_SUBJECT", _JoinArray($aNums, " - "))
    $v.Add("J_FILE", _JoinArray($aNums, "_"))
    $v.Add("DELIVERY", $delivery)
    $v.Add("DELIVERY_SHORT", _ExtractDeliveryDateTime($delivery))
    $v.Add("DATE_JOUR", @MDAY & "/" & @MON & "/" & @YEAR)
    $v.Add("OPERATOR", $g_sOperator)
    $v.Add("CARRIER", $carrier)
    $v.Add("NL", @CRLF)
    Return $v
EndFunc

Func _SP_ApplyTemplate($tpl, ByRef $vars)
    Local $s = $tpl
    For $k In $vars.Keys
        $s = StringReplace($s, "{" & $k & "}", $vars.Item($k))
    Next
    Return $s
EndFunc

; ============================================================
; INPUT / ETMS / LIVE OUTPUT
; ============================================================
Func _NowText()
    Return @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc

Func _LastBL_Write(ByRef $aNums)
    If Not FileExists($INPUT_PATH) Then DirCreate($INPUT_PATH)
    IniWrite($LASTBL_MARKER, "LAST", "Date", _NowText())
    IniWrite($LASTBL_MARKER, "LAST", "Operator", $g_sOperator)
    IniWrite($LASTBL_MARKER, "LAST", "WindowsUser", @UserName)
    IniWrite($LASTBL_MARKER, "LAST", "Machine", @ComputerName)
    IniWrite($LASTBL_MARKER, "LAST", "BL", _JoinArray($aNums, " / "))
    IniWrite($LASTBL_MARKER, "LAST", "InputCSV", $INPUT_CSV)
    IniWrite($LASTBL_MARKER, "LAST", "InputGEN", $INPUT_GEN)
    _LastBL_Refresh()
    Return True
EndFunc

Func _LastBL_Text()
    If Not FileExists($LASTBL_MARKER) Then Return "Dernier BL : aucune info"
    Local $d = IniRead($LASTBL_MARKER, "LAST", "Date", "")
    Local $op = IniRead($LASTBL_MARKER, "LAST", "Operator", "")
    Local $wu = IniRead($LASTBL_MARKER, "LAST", "WindowsUser", "")
    Local $pc = IniRead($LASTBL_MARKER, "LAST", "Machine", "")
    Local $bl = IniRead($LASTBL_MARKER, "LAST", "BL", "")
    If $op = "" Then $op = $wu
    If $pc <> "" Then $pc = " / " & $pc
    If $bl = "" Then Return "Dernier BL : " & $op & $pc & " - " & $d
    Return "Dernier BL : " & $bl & @CRLF & "Par " & $op & $pc & " le " & $d
EndFunc

Func _LastBL_Refresh()
    If $idLastBL <> 0 Then GUICtrlSetData($idLastBL, _LastBL_Text())
    Return True
EndFunc

Func _ResetFiles()
    _CsvCacheClear() ; PRO CLEAN : évite toute donnée CSV ancienne en mémoire
    If FileExists($INPUT_CSV) Then FileDelete($INPUT_CSV)
    If FileExists($INPUT_GEN) Then FileDelete($INPUT_GEN)
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        If FileExists($OUTPUT_PATH & $a[$i]) Then FileDelete($OUTPUT_PATH & $a[$i])
    Next
EndFunc

Func _CreateInputFiles(ByRef $aNums)
    Local $first = _CleanJ($aNums[0])
    If $first = "" Then
        _Status("INPUT KO : premier J vide.")
        Return False
    EndIf

    Local $tracking = $first & $g_sOperator
    Local $csv = ""
    For $i = 0 To UBound($aNums) - 1
        Local $sJ = _CleanJ($aNums[$i])
        If $sJ <> "" Then $csv &= $sJ & "," & $tracking & @CRLF
    Next

    If StringStripWS($csv, 3) = "" Then
        _Status("INPUT KO : aucun J valide à écrire.")
        Return False
    EndIf

    If Not _WriteAnsi($INPUT_CSV, $csv) Then
        _Status("INPUT KO : impossible de créer" & @CRLF & $INPUT_CSV)
        Return False
    EndIf

    If Not _WriteAnsi($INPUT_GEN, $first & "," & $tracking & @CRLF) Then
        _Status("INPUT KO : impossible de créer" & @CRLF & $INPUT_GEN)
        Return False
    EndIf

    _LastBL_Write($aNums)
    _Status("INPUT OK :" & @CRLF & $INPUT_CSV & @CRLF & $INPUT_GEN & @CRLF & _LastBL_Text())
    Return True
EndFunc

Func _InputReady(ByRef $aNums)
    Local $try = 0
    Local $last = ""
    While $try < 25
        $try += 1
        If FileExists($INPUT_CSV) And FileGetSize($INPUT_CSV) > 0 And FileExists($INPUT_GEN) And FileGetSize($INPUT_GEN) > 0 Then
            Local $sContent = FileRead($INPUT_CSV)
            Local $sGen = FileRead($INPUT_GEN)
            Local $first = _CleanJ($aNums[0])
            Local $tracking = $first & $g_sOperator
            Local $ok = True
            If Not StringInStr($sGen, $first & "," & $tracking) Then
                $ok = False
                $last = "inputGEN.csv ne contient pas : " & $first & "," & $tracking
            EndIf
            For $i = 0 To UBound($aNums) - 1
                Local $sJ = _CleanJ($aNums[$i])
                If $sJ <> "" And Not StringInStr($sContent, $sJ & "," & $tracking) Then
                    $ok = False
                    $last = "input.csv ne contient pas : " & $sJ & "," & $tracking
                    ExitLoop
                EndIf
            Next
            If $ok Then Return True
        Else
            $last = "Fichiers INPUT absents ou vides :" & @CRLF & $INPUT_CSV & " size=" & FileGetSize($INPUT_CSV) & @CRLF & $INPUT_GEN & " size=" & FileGetSize($INPUT_GEN)
        EndIf
        Sleep(250)
    WEnd
    _Status("INPUT KO après contrôle :" & @CRLF & $last)
    Return False
EndFunc

Func _OpenEdsInEtms()
    Local $hWnd = WinGetHandle($ETMS_WINDOW)
    If $hWnd = 0 Then
        MsgBox(16+262144, $APP_TITLE, "ETMS introuvable : [CLASS:TfmBrowser].")
        Return False
    EndIf
    If Not FileExists($EDS_PATH) Then
        MsgBox(16+262144, $APP_TITLE, "EDS introuvable :" & @CRLF & $EDS_PATH)
        Return False
    EndIf
    Local $ok = False, $try = 0
    While Not $ok And $try < 3
        $try += 1
        If $try > 1 And WinExists($FILEOPEN_WIN) Then WinClose($FILEOPEN_WIN)
        WinActivate($hWnd)
        WinWaitActive($hWnd, "", 3)
        Sleep(500)
        ControlClick($hWnd, "", $ETMS_TOOLBAR, "LEFT", 1, $TOOLBAR_X, $TOOLBAR_Y)
        Local $t = TimerInit()
        While Not WinExists($FILEOPEN_WIN)
            Sleep(100)
            If _IsPressed("1B") Then Return False
            If TimerDiff($t) > 10000 Then ExitLoop
        WEnd
        If Not WinExists($FILEOPEN_WIN) Then ContinueLoop
        WinActivate($FILEOPEN_WIN)
        WinWaitActive($FILEOPEN_WIN, "", 3)
        ControlSetText($FILEOPEN_WIN, "", $FILEOPEN_EDIT, "")
        Sleep(100)
        ControlSetText($FILEOPEN_WIN, "", $FILEOPEN_EDIT, $EDS_PATH)
        Sleep(200)
        Send("{ENTER}")
        Sleep(800)
        If Not WinExists($FILEOPEN_WIN) Then $ok = True
    WEnd
    Return $ok
EndFunc

Func _OutputsOK()
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        Local $p = $OUTPUT_PATH & $a[$i]
        If Not FileExists($p) Then Return False
        If FileGetSize($p) <= 0 Then Return False
    Next
    Return True
EndFunc

Func _GetCarrierLive()
    Return _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 7)
EndFunc

Func _GetDeliveryLive()
    Local $d = _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 8)
    If _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 9) <> "" Then $d &= " " & _CsvGet($OUTPUT_PATH & "outputEDICEC.csv", 0, 9)
    Return $d
EndFunc

; ============================================================
; HTML / PDF FROM SNAPSHOT
; ============================================================
Func _OpenPreviewEditor($htmlFile, ByRef $aNums, $idx)
    If Not FileExists($htmlFile) Then Return False
    Local $hPrev = GUICreate("Modifier BL - " & _JoinArray($aNums, " / "), 1120, 820, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_SIZEBOX))
    GUICtrlCreateLabel("Modifie le BL puis clique Enregistrer. Enregistrer sauvegarde le HTML ET régénère le PDF utilisé pour le mail.", 10, 8, 970, 22)
    Local $bSave = GUICtrlCreateButton("Enregistrer", 10, 35, 150, 32)
    Local $bPdfOpen = GUICtrlCreateButton("Ouvrir PDF", 170, 35, 140, 32)
    Local $bHtml = GUICtrlCreateButton("Ouvrir HTML", 320, 35, 140, 32)
    Local $bClose = GUICtrlCreateButton("Fermer", 970, 35, 130, 32)
    Local $oIE = ObjCreate("Shell.Explorer.2")
    If Not IsObj($oIE) Then
        MsgBox(16+262144, $APP_TITLE, "Contrôle aperçu HTML indisponible.")
        GUIDelete($hPrev)
        Return False
    EndIf
    GUICtrlCreateObj($oIE, 10, 75, 1090, 730)
    $oIE.Silent = True
    $oIE.Navigate("file:///" & StringReplace($htmlFile, "\", "/"))
    While $oIE.Busy Or $oIE.ReadyState <> 4
        Sleep(100)
    WEnd
    GUISetState(@SW_SHOW, $hPrev)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $bClose
                ExitLoop
            Case $bSave
                Local $pdfSaved = _SavePreviewAndRebuildPdf($oIE, $aNums, $idx, True)
                If FileExists($pdfSaved) Then
                    MsgBox(64, $APP_TITLE, "Modifications enregistrées et PDF régénéré." & @CRLF & @CRLF & $pdfSaved)
                Else
                    MsgBox(16, $APP_TITLE, "Impossible d’enregistrer ou de régénérer le PDF.")
                EndIf
            Case $bHtml
                ShellExecute($htmlFile)
            Case $bPdfOpen
                If $idx >= 0 Then
                    _EnsureLatestPdfByIndex($idx)
                    If $idx < UBound($g_aPDF) And FileExists($g_aPDF[$idx]) Then ShellExecute($g_aPDF[$idx])
                EndIf
        EndSwitch
    WEnd
    GUIDelete($hPrev)
    Return True
EndFunc

Func _SaveEditedPreview(ByRef $oIE, ByRef $aNums)
    Local $html = ""
    If IsObj($oIE) And IsObj($oIE.Document) Then $html = "<!doctype html>" & $oIE.Document.documentElement.outerHTML
    If $html = "" Then Return ""
    Local $name = "Printed_Delivery_Order_" & _Safe(_JoinArray($aNums, "_")) & "_MODIFIE_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".html"
    Local $file = $HTML_PATH & $name
    If Not _WriteUtf8($file, $html) Then
        Local $fallback = @ScriptDir & "\HTML\"
        If Not FileExists($fallback) Then DirCreate($fallback)
        $file = $fallback & $name
        If Not _WriteUtf8($file, $html) Then Return ""
    EndIf
    Return $file
EndFunc

Func _SavePreviewAndRebuildPdf(ByRef $oIE, ByRef $aNums, $idx, $bOpenPdf = False)
    Local $edited = _SaveEditedPreview($oIE, $aNums)
    If Not FileExists($edited) Then Return ""
    Local $carrier = ""
    Local $delivery = ""
    If $idx >= 0 And $idx < UBound($g_aCarrier) Then
        $carrier = $g_aCarrier[$idx]
        $delivery = $g_aDelivery[$idx]
    EndIf
    Local $pdf = _HtmlToPdf($edited, $aNums, $carrier, $delivery)
    If FileExists($pdf) Then
        If $idx >= 0 And $idx < UBound($g_aPDF) Then
            $g_aHTML[$idx] = $edited
            $g_aPDF[$idx] = $pdf
        EndIf
        _Status("Modifications enregistrées + PDF régénéré :" & @CRLF & $pdf)
        If $bOpenPdf Then ShellExecute($pdf)
        Return $pdf
    EndIf
    Return ""
EndFunc

Func _EnsureLatestPdfByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return ""
    If Not FileExists($g_aHTML[$idx]) Then Return ""
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _HtmlToPdf($g_aHTML[$idx], $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
    If FileExists($pdf) Then
        $g_aPDF[$idx] = $pdf
        Return $pdf
    EndIf
    Return $g_aPDF[$idx]
EndFunc

Func _GenerateHtmlFromSnapshot(ByRef $aNums, $snap)
    Local $rows = "", $totalPcs = 0, $totalWeight = 0, $totalVol = 0
    For $r = 0 To 39
        Local $j = _SnapCsvGet($snap, "outputDIMS.csv", $r, 0)
        If $j = "" Then ContinueLoop
        Local $dim = _SnapCsvGet($snap, "outputDIMS.csv", $r, 1)
        Local $pcs = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 1)
        Local $wgt = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 2)
        Local $vol = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 3)
        Local $packid = _JoinUniqueColsSnap($snap, "outputPACKID.csv", $r, 2, 12, " - ")
        Local $ref = _JoinUniqueColsSnap($snap, "outputREFS.csv", $r, 2, 12, " - ")
        $rows &= "<tr><td contenteditable='true'>" & _H($j) & "</td><td contenteditable='true'>" & _H($ref) & "</td><td contenteditable='true'>" & _H($packid) & "</td><td contenteditable='true'>" & _H($pcs) & "</td><td contenteditable='true'>" & _H($wgt) & "</td><td contenteditable='true'>" & _H($dim) & "</td></tr>"
        $totalPcs += Number(_CleanNum($pcs))
        $totalWeight += Number(_CleanNum($wgt))
        $totalVol += Number(_CleanNum($vol))
    Next
    Local $deliver = _FormatDeliverTo(_DefaultDeliverToSnap($snap))
    Local $ready = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 1)
    Local $delivDate = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 5)
    Local $delivTime = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 6)
    Local $carrier = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 7)
    Local $instr = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 8)
    If _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9) <> "" Then $instr &= " " & _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9)
    $instr = _FormatDeliveryInstructions($instr)
    Local $first = _CleanJ($aNums[0])
    Local $tracking = $first & $g_sOperator
    Local $html = _HtmlTemplate($tracking, $g_sOperator, $deliver, $ready, $delivDate, $delivTime, $carrier, $instr, $rows, _TrimNum($totalPcs), _TrimNum($totalWeight), _TrimNum($totalVol))
    Local $name = "Printed_Delivery_Order_" & _Safe(_JoinArray($aNums, "_")) & "_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".html"
    Local $file = $HTML_PATH & $name
    If Not _WriteUtf8($file, $html) Then
        Local $fallback = @ScriptDir & "\HTML\"
        If Not FileExists($fallback) Then DirCreate($fallback)
        $file = $fallback & $name
        If Not _WriteUtf8($file, $html) Then Return ""
    EndIf
    Return $file
EndFunc

Func _DefaultDeliverToSnap($snap)
    Local $s = _SnapCsvGet($snap, "outputGEN.csv", 0, 1)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 2) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 2)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 3) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 3)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 4) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 4)
    Return $s
EndFunc

Func _SnapGetCompany($snap)
    Local $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 1)
    If $company = "" Then $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 2)
    If $company = "" Then $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 3)
    $company = StringReplace($company, @CRLF, " ")
    $company = StringReplace($company, @CR, " ")
    $company = StringReplace($company, @LF, " ")
    While StringInStr($company, "  ")
        $company = StringReplace($company, "  ", " ")
    WEnd
    Return StringStripWS($company, 3)
EndFunc

Func _SnapGetDelivery($snap)
    Local $d = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 8)
    If _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9) <> "" Then $d &= " " & _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9)
    Return $d
EndFunc

Func _FormatDeliveryInstructions($s)
    Local $x = StringStripWS($s, 3)
    Local $rep = "$1" & @CRLF
    $x = StringRegExpReplace($x, "(?i)(entre\s+[0-9]{1,2}h(?:[0-9]{0,2})?\s+et\s+[0-9]{1,2}h(?:[0-9]{0,2})?)\s+", $rep)
    Return $x
EndFunc

Func _FormatDeliverTo($s)
    $s = StringReplace($s, " Tél", @CRLF & "Tél")
    $s = StringReplace($s, " Tel", @CRLF & "Tel")
    $s = StringReplace($s, " TEL", @CRLF & "TEL")
    $s = StringReplace($s, " Phone", @CRLF & "Phone")
    $s = StringReplace($s, " Mobile", @CRLF & "Mobile")
    $s = StringReplace($s, " +33", @CRLF & "+33")
    Return _H($s)
EndFunc

Func _HtmlTemplate($tracking, $operator, $deliver, $ready, $delivDate, $delivTime, $carrier, $instr, $rows, $totalPcs, $totalWeight, $totalVol)
    Local $date = @MDAY & "/" & @MON & "/" & @YEAR
    Local $s = "<!doctype html><html lang='fr'><head><meta charset='utf-8'><title>DELIVERY ORDER - " & _H($tracking) & "</title>" & _Css() & "</head><body>"
    $s &= "<main class='page'>"
    $s &= "<section class='header'><div class='logo'><img src='" & _LogoDataUri() & "'></div><div class='head-right'><div class='doc-title'>DELIVERY ORDER</div><div class='datebox'><span class='label'>Date :</span> " & _H($date) & "</div></div></section>"
    $s &= "<section class='topline'><div class='line'><span class='label'>Tracking# :</span><span class='value'>" & _H($tracking) & "</span></div><div class='line'><span class='label'>Operator :</span><span class='value'>" & _H($operator) & "</span></div></section>"
    $s &= "<section class='boxgrid'><div class='box'><div class='boxtitle'>DELIVER TO:</div><div class='addr' contenteditable='true'>" & $deliver & "</div></div><div class='box'><div class='boxtitle'>Pick Up From:</div><div contenteditable='true'>Expeditors International France SAS<br>La Porte des Champs<br>Batiment B - Quais 21 a 27<br>95470 Survilliers CEDEX</div></div></section>"
    $s &= "<section class='sectiongrid'><div class='info'><span class='label'>DELIVERING CARRIER:</span><br><span contenteditable='true'>" & _H($carrier) & "</span></div><div class='info'><span class='label'>Commodity:</span><br><span contenteditable='true'>IT spare parts</span></div></section>"
    $s &= "<section class='schedule'><div><span class='label'>Ready:</span><br><span contenteditable='true'>" & _H($ready) & "</span></div><div><span class='label'>Delivery date:</span><br><span contenteditable='true'>" & _H($delivDate) & "</span></div><div><span class='label'>Delivery Time:</span><br><span contenteditable='true'>" & _H($delivTime) & "</span></div></section>"
    $s &= "<section class='instructions'><span class='label'>Delivery Instructions:</span><br><div contenteditable='true'>" & _H($instr) & "</div></section>"
    $s &= "<table><thead><tr><th>Expeditors Ref</th><th>W/o &amp; S/o Ref</th><th>Pack ID</th><th>Pcs count</th><th>Weight</th><th>Dims</th></tr></thead><tbody>" & $rows
    $s &= "<tr class='total'><td>Total:</td><td></td><td></td><td>" & _H($totalPcs) & "</td><td>" & _H($totalWeight) & "</td><td>" & _H($totalVol) & "</td></tr></tbody></table>"
    $s &= "<section class='signatures'><div><div class='sigline'>X</div><div class='sublabel'>(signature)</div><div class='sigline second'>X</div><div class='sublabel'>(name)</div></div><div class='received'><div class='received-row'><span>DATE RECEIVED:</span><span class='received-line'></span></div><div class='received-row'><span>TIME RECEIVED:</span><span class='received-line'></span></div></div></section>"
    $s &= "</main></body></html>"
    Return $s
EndFunc

Func _Css()
    Return "<style>*{box-sizing:border-box}body{margin:0;background:#e9e9e9;color:#111;font-family:Calibri,Arial,sans-serif;font-size:11pt}.page{width:210mm;min-height:297mm;margin:0 auto;background:#fff;padding:12mm}.header{display:grid;grid-template-columns:62mm 1fr 58mm;gap:8mm;margin-bottom:8px}.logo img{width:58mm}.head-right{grid-column:3;padding-top:2mm}.doc-title{font-size:16pt;font-weight:700;margin:0 0 7mm 0}.topline{display:grid;grid-template-columns:1fr 1fr;gap:18mm;margin-top:6px}.line{display:flex;gap:8px;margin:4px 0}.label{font-weight:700;white-space:nowrap}.boxgrid,.sectiongrid{display:grid;grid-template-columns:1fr 1fr;gap:12mm;margin-top:10px}.box{border:1.5px solid #111;min-height:34mm;padding:7px}.boxtitle{font-weight:700;margin-bottom:6px;text-transform:uppercase}.addr{white-space:normal;word-break:break-word}.info,.schedule,.instructions{border:1px solid #111;padding:7px}.schedule{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:10px}.instructions{margin-top:10px;min-height:14mm}table{width:100%;border-collapse:collapse;margin-top:12px;table-layout:fixed}th,td{border:1px solid #111;padding:5px 4px;vertical-align:top;word-wrap:break-word;font-size:11pt}th{background:#f2f2f2;text-align:left}th:nth-child(4),th:nth-child(5),th:nth-child(6),td:nth-child(4),td:nth-child(5),td:nth-child(6){text-align:center}.total td{font-weight:700;background:#fafafa}.signatures{display:grid;grid-template-columns:1fr 1fr;gap:14mm;margin-top:18mm}.sigline{border-bottom:1px solid #111;height:15mm;font-size:18pt;padding-left:8px}.second{margin-top:8mm}.sublabel{text-align:center}.received-row{display:flex;align-items:flex-end;gap:8px;margin-bottom:13mm;font-weight:700}.received-line{border-bottom:1px solid #111;height:8mm;flex:1}[contenteditable='true']:focus{outline:1px dashed #777;background:#fffbe6}@media print{body{background:#fff}.page{margin:0;width:auto;padding:10mm}@page{size:A4 portrait;margin:8mm}}</style>"
EndFunc

Func _HtmlToPdf($sHtml, ByRef $aNums, $carrier, $delivery)
    If Not FileExists($sHtml) Then Return ""
    DirCreate($PDF_PATH)
    If $carrier = "" Then $carrier = "Autre"
    Local $rule = _SP_GetRule($carrier)
    Local $vars = _SP_BuildVariables($aNums, $delivery, $carrier)
    Local $pdfBase = _SP_ApplyTemplate($rule.Item("PDF"), $vars)
    Local $pdf = $PDF_PATH & _Safe($pdfBase) & ".pdf"
    If FileExists($pdf) Then FileDelete($pdf)
    Local $browser = _FindBrowser()
    If $browser = "" Then Return ""
    Local $url = "file:///" & StringReplace($sHtml, "\", "/")
    Local $cmd = '"' & $browser & '" --headless --disable-gpu --print-to-pdf="' & $pdf & '" "' & $url & '"'
    RunWait($cmd, @ScriptDir, @SW_HIDE)
    If FileExists($pdf) And FileGetSize($pdf) > 0 Then Return $pdf
    Return ""
EndFunc

Func _FindBrowser()
    If $g_sBrowserCache <> "" And FileExists($g_sBrowserCache) Then Return $g_sBrowserCache
    Local $a[6] = [@ProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe", @ProgramFilesDir & " (x86)\Microsoft\Edge\Application\msedge.exe", @LocalAppDataDir & "\Microsoft\Edge\Application\msedge.exe", @ProgramFilesDir & "\Google\Chrome\Application\chrome.exe", @ProgramFilesDir & " (x86)\Google\Chrome\Application\chrome.exe", @LocalAppDataDir & "\Google\Chrome\Application\chrome.exe"]
    For $i = 0 To UBound($a) - 1
        If FileExists($a[$i]) Then
            $g_sBrowserCache = $a[$i]
            Return $g_sBrowserCache
        EndIf
    Next
    Return ""
EndFunc
; ============================================================
; MODERN UX HELPERS - HOVER SOFT / NON BLOQUANT
; ============================================================
Func _CreateHoverButton($sText, $x, $y, $w, $h, $normalColor = 0xF8FAFC, $hoverColor = 0xE2E8F0, $textColor = 0x111827)
    Local $idButton = GUICtrlCreateButton($sText, $x, $y, $w, $h)
    GUICtrlSetFont($idButton, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($idButton, $normalColor)
    GUICtrlSetColor($idButton, $textColor)
    _RegisterHoverButton($idButton, $normalColor, $hoverColor)
    Return $idButton
EndFunc
Func _RegisterHoverButton($idButton, $normalColor, $hoverColor)
    Local $n = UBound($g_aHoverIds)
    ReDim $g_aHoverIds[$n + 1]
    ReDim $g_aHoverNormal[$n + 1]
    ReDim $g_aHoverHover[$n + 1]
    ReDim $g_aHoverIsHover[$n + 1]
    $g_aHoverIds[$n] = $idButton
    $g_aHoverNormal[$n] = $normalColor
    $g_aHoverHover[$n] = $hoverColor
    $g_aHoverIsHover[$n] = False
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
Func _SetProcessingUi($bProcessing)
    If $idBtnForceNext <> 0 Then
        If $bProcessing Then
            GUICtrlSetState($idBtnForceNext, $GUI_ENABLE)
        Else
            GUICtrlSetState($idBtnForceNext, $GUI_DISABLE)
        EndIf
    EndIf
EndFunc
Func _PumpGuiDuringProcess()
    Local $msg = GUIGetMsg()
    _UpdateHoverButtons()
    Switch $msg
        Case 0
            Return
        Case $GUI_EVENT_CLOSE
            _HandleCloseRequest()
        Case $idBtnStopQueue
            $g_bStopQueue = True
            _Status("Stop demandé. Le BL en cours se termine puis la queue s’arrête.")
        Case $idBtnForceNext
            $g_bForceNext = True
            _Status("Forçage demandé : contrôle des outputs puis passage à l’étape suivante.")
        Case $idBtnMail
            _SendMailOnlySmart()
        Case $idBtnEdocOnly
            _UploadEdocOnlySmart()
        Case $idBtnEdocMail
            _SendEdocMailSmart()
        Case $idBtnRefreshResults, $idBtnRefreshResultsTop
            _RefreshResultsList()
    EndSwitch
EndFunc
Func _HandleCloseRequest()
    If Not $g_bGenerating Then Return True
    Local $choice = _ModernCloseDialog()
    Switch $choice
        Case 1
            Return False
        Case 2
            $g_bStopQueue = True
            _Status("Fermeture demandée : arrêt propre après le BL en cours.")
            Return False
        Case 3
            $g_bHardClose = True
            $g_bStopQueue = True
            Return True
    EndSwitch
    Return False
EndFunc
Func _ModernCloseDialog()
    Local $r = MsgBox(35 + 262144, $APP_TITLE, "Un BL est encore en cours." & @CRLF & @CRLF & "Oui = continuer" & @CRLF & "Non = arrêt propre après BL en cours" & @CRLF & "Annuler = fermer maintenant")
    If $r = 6 Then Return 1
    If $r = 7 Then Return 2
    Return 3
EndFunc
Func _ModernInfo($sTitle, $sMessage, $sType = "info")
    If $g_hGui = 0 Then
        _Status($sTitle & " - " & $sMessage)
        Return
    EndIf
    MsgBox(64, $sTitle, $sMessage)
EndFunc
Func _ModernErrorChoice($sTitle, $sMessage, $sBtn1, $sBtn2, $sBtn3)
    If $g_hGui = 0 Then
        _Status($sTitle & " - " & $sMessage & " / action auto: continuer")
        Return 1
    EndIf
    Local $r = MsgBox(51 + 262144, $sTitle, $sMessage & @CRLF & @CRLF & "Oui = " & $sBtn1 & @CRLF & "Non = " & $sBtn2 & @CRLF & "Annuler = " & $sBtn3)
    If $r = 6 Then Return 1
    If $r = 7 Then Return 2
    Return 3
EndFunc

; ============================================================
; CSV / SNAP / MAIL / EDOC / UTILS
; ============================================================
Func _SnapCsvGet($snap, $fileName, $rowIndex, $colIndex)
    Return _CsvGet($snap & $fileName, $rowIndex, $colIndex)
EndFunc

Func _CsvGet($file, $rowIndex, $colIndex)
    Local $lines = _CsvLoadCached($file)
    If Not IsArray($lines) Then Return ""
    If $rowIndex < 0 Or $rowIndex >= UBound($lines) Then Return ""
    Local $a = _ParseCsvLine($lines[$rowIndex])
    If $colIndex < 0 Or $colIndex >= UBound($a) Then Return ""
    Return _CleanJ($a[$colIndex])
EndFunc

Func _GetLine($file, $rowIndex)
    Local $lines = _CsvLoadCached($file)
    If Not IsArray($lines) Then Return ""
    If $rowIndex < 0 Or $rowIndex >= UBound($lines) Then Return ""
    Return $lines[$rowIndex]
EndFunc

Func _ParseCsvLine($line)
    Local $delim = ","
    If StringInStr($line, ";") And Not StringInStr($line, ",") Then $delim = ";"
    If StringInStr($line, @TAB) Then $delim = @TAB
    Local $a[0], $field = "", $inQ = False
    For $i = 1 To StringLen($line)
        Local $ch = StringMid($line, $i, 1)
        If $ch = '"' Then
            If $inQ And $i < StringLen($line) And StringMid($line, $i + 1, 1) = '"' Then
                $field &= '"'
                $i += 1
            Else
                $inQ = Not $inQ
            EndIf
        ElseIf $ch = $delim And Not $inQ Then
            ReDim $a[UBound($a) + 1]
            $a[UBound($a) - 1] = $field
            $field = ""
        Else
            $field &= $ch
        EndIf
    Next
    ReDim $a[UBound($a) + 1]
    $a[UBound($a) - 1] = $field
    Return $a
EndFunc

Func _JoinUniqueColsSnap($snap, $fileName, $row, $c1, $c2, $sep)
    Local $d = ObjCreate("Scripting.Dictionary")
    Local $res = ""
    For $c = $c1 To $c2
        Local $v = _SnapCsvGet($snap, $fileName, $row, $c)
        If $v = "" Then ContinueLoop
        If Not $d.Exists($v) Then
            $d.Add($v, 1)
            If $res <> "" Then $res &= $sep
            $res &= $v
        EndIf
    Next
    Return $res
EndFunc

Func _CreateMailSP($sPdf, ByRef $aNums, $carrierOverride = "", $deliveryOverride = "")
    Local $carrier = $carrierOverride
    If $carrier = "" Then $carrier = "Autre"
    Local $delivery = $deliveryOverride
    Local $rule = _SP_GetRule($carrier)
    Local $vars = _SP_BuildVariables($aNums, $delivery, $carrier)
    Local $subject = _SP_ApplyTemplate($rule.Item("SUBJECT"), $vars)
    Local $body = _SP_ApplyTemplate($rule.Item("BODY"), $vars)
    Local $signatureAdd = _SP_ApplyTemplate($rule.Item("SIGNATURE"), $vars)
    If $signatureAdd <> "" Then $body &= @CRLF & @CRLF & $signatureAdd
    Local $oOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($oOutlook) Then Return False
    Local $mail = $oOutlook.CreateItem(0)
    $mail.To = $rule.Item("TO")
    $mail.CC = $rule.Item("CC")
    $mail.BCC = $rule.Item("BCC")
    $mail.Subject = $subject
    If FileExists($sPdf) Then $mail.Attachments.Add($sPdf)
    $mail.Display
    Sleep(700)
    Local $sig = $mail.HTMLBody
    $mail.HTMLBody = _TextToHtml($body) & "<br><br>" & $sig
    If $MAIL_AUTOSEND Then $mail.Send
    Return True
EndFunc

Func _TextToHtml($s)
    $s = _H($s)
    $s = StringReplace($s, @CRLF, "<br>")
    $s = StringReplace($s, @CR, "<br>")
    $s = StringReplace($s, @LF, "<br>")
    Return "<div style='font-family:Calibri,Arial;font-size:11pt'>" & $s & "</div>"
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
Func _EdocValidateUploadWindow($sDocType, ByRef $aNums)
    If Not WinWait("Upload Documents CDG", "", 35) Then
        _Status("EDOC KO : la fenêtre 'Upload Documents CDG' n'est pas apparue.")
        Return False
    EndIf
    WinActivate("Upload Documents CDG")
    If Not WinWaitActive("Upload Documents CDG", "", 8) Then Return False
    Sleep(3500)

    ; 1 seul dossier : TAB -> Delivery Order -> ENTER.
    ; Plusieurs dossiers : TAB -> Delivery Order -> TAB 1 -> TAB 2 -> collage numéros -> validation.
    Send("{TAB}")
    Sleep(900)
    Send($sDocType)
    Sleep(1200)

    If UBound($aNums) <= 1 Then
        Send("{ENTER}")
    Else
        Send("{TAB}")
        Sleep(900)
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
Func _ArrayNumsClipboard(ByRef $aNums)
    Local $s = ""
    For $i = 0 To UBound($aNums) - 1
        Local $j = _CleanJ($aNums[$i])
        If $j <> "" Then $s &= $j & @CRLF
    Next
    Return $s
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

Func _CleanJ($s)
    $s = String($s)
    $s = StringReplace($s, Chr(239) & Chr(187) & Chr(191), "")
    $s = StringReplace($s, "ï»¿", "")
    $s = StringReplace($s, "Ï»¿", "")
    $s = StringReplace($s, ChrW(65279), "")
    Return StringStripWS($s, 3)
EndFunc

Func _CleanNum($s)
    Return StringReplace($s, ",", ".")
EndFunc

Func _TrimNum($n)
    Local $s = StringFormat("%.3f", $n)
    While StringRight($s, 1) = "0"
        $s = StringTrimRight($s, 1)
    WEnd
    If StringRight($s, 1) = "." Then $s = StringTrimRight($s, 1)
    Return $s
EndFunc

Func _JoinArray(ByRef $a, $sep)
    Local $s = ""
    For $i = 0 To UBound($a) - 1
        If $s <> "" Then $s &= $sep
        $s &= _CleanJ($a[$i])
    Next
    Return $s
EndFunc

Func _H($s)
    $s = String($s)
    $s = StringReplace($s, "&", "&amp;")
    $s = StringReplace($s, "<", "&lt;")
    $s = StringReplace($s, ">", "&gt;")
    $s = StringReplace($s, '"', "&quot;")
    $s = StringReplace($s, @CRLF, "<br>")
    $s = StringReplace($s, @CR, "<br>")
    $s = StringReplace($s, @LF, "<br>")
    Return $s
EndFunc

Func _WriteAnsi($file, $txt)
    Local $sDir = StringLeft($file, StringInStr($file, "\", 0, -1))
    If $sDir <> "" And Not FileExists($sDir) Then DirCreate($sDir)
    Local $h = FileOpen($file, 2)
    If $h = -1 Then Return False
    FileWrite($h, $txt)
    FileClose($h)
    Return FileExists($file)
EndFunc

Func _WriteUtf8($file, $txt)
    Local $sDir = StringLeft($file, StringInStr($file, "\", 0, -1))
    If $sDir <> "" And Not FileExists($sDir) Then DirCreate($sDir)
    Local $o = ObjCreate("ADODB.Stream")
    If Not IsObj($o) Then Return False
    $o.Type = 2
    $o.Charset = "utf-8"
    $o.Open
    $o.WriteText($txt)
    $o.SaveToFile($file, 2)
    $o.Close
    If Not FileExists($file) Then Return False
    If FileGetSize($file) <= 0 Then Return False
    Return True
EndFunc

Func _Safe($s)
    Local $a[12] = ["\", "/", ":", "*", "?", '"', "<", ">", "|", @CR, @LF, @TAB]
    For $i = 0 To UBound($a) - 1
        $s = StringReplace($s, $a[$i], "_")
    Next
    Return StringStripWS($s, 3)
EndFunc

Func _LogoDataUri()
    Local $s = ""
    $s &= "iVBORw0KGgoAAAANSUhEUgAAAQUAAABhCAYAAADfhLRpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAACbVSURBVHhe7Z0JnI3VG8epkH1nGDN2svwrZSdkCZU1a4k2lUhSISkkUvyzVkryV6mUiopKpZJISouKKMqaMGPfef6f7zOd671n7szcOzN3zOX8Pp/zmbnvPe/7nuV5fuc5z3nOuVnEwcHBwYMs9gUHB4dzG44UHBwc/OBIwcHBwQ+OFBwcHPzgSMHBwcEPjhQcHBz84EjBwcHBD44UHBwc/OBIwSHT4uTGv+TYh5/KsU8+z3zpo0/l+LJv7CKfFXCk4JApceyzpRIfe7HER1WR+BJVM1cqVlniil8kR1+daxf7rIAjBYdMh+NLlkv8RXUlLneMxOUvk7lSzpISF/sfOTxxmsiJE3bRzwo4UnDIVDg6+02JK1JR4vLGSlzBcpkr5YmR+Ao15diCRXaxzyo4UnDINDj89AyJK1lN4vKWTqyQZzrliZU9F9WVYx9/Zhf7rIMjBYczj+PH5ciU6ap4u7OXSJg22OlMWg55YtR6OfruB3bJz0o4UnA44zgybabsbdJG9rXuIvuu7po4te8hextdI3GFK0hcgbKJlTacKXeMxMdeIkf+95pd7LMWjhQczhhOHTosh4aPlX3NOogcOSqn9u5LnI4eleNffyt7G10rcfkyeFqRM1r2Nu8gx7/9wS76WQ1HCg5nBidPysG7h8juLPkTSCEJnFj9q+xt0DrBhLeVNpwpZ7RaLsRKnGtwpOCQ4Th15Ijs73KLjvy7c0XLvuYd7SyK48tXyp5azSQuV6nEShvOlLuU7GvTXU5u2W4X6ZyAIwWHjMXx"
    $s &= "43LwgeE6EsflL5skKRxbvET2VKsncbmiEyttuBL+ihwlZE/1BnJi/Qa7SOcMHCk4ZBhO7dsvBx96TOIKl08IBCpYLjEpHD8ux5Z+LfFlLpXdWQrJ7mzFT6ccJRIrcnql/GUlLl8ZnTKcWLvOW+xzDo4UHDIEJ9aul32tOieM/P8SAskmhVM7d8vh/06Vg/cPl0NDHzudHnlcDg4YqoqbSKHTmihP3li1YE7tP+BX7nMRjhQcMgQsO+7OUjCRQtqkkBzUykhvhyNThgJl5eDDY+TU4SP2K89JOFJwyBAcmfGKTgFspQyKFA4fkcNjJiSY+OkZp8ASZ65ScnDYaF0NcUiAIwWHDEFqSeHkjp1yoFdfiSdwyTPtSHMilLpoJTk84Vk3ZbDgSMEhQ5AaUji5ZZvsrdVMdmcrlqDEWArpkfLESnypanJk5mz7lQ6OFBxSg0ULF0mX1l3kxg43+qXubbrL4LsH29kVoZLCyc1bZW+t5hJXpIKO6OmWClfQlY2jc96xX+nwL8JGCpv+2iRzX50rc16eIwvmLbC/Dgnbt26X12a9Jm/OftOX/joHI80yA3b+s1NK5CghhbIUkqjsUVLywpK+VPS8otL40sb2LYpQSeFU/B45+uY8Ofreh+mb5i+UY59/Zb/OwYOwkMJniz6Thv9pKPmy5JOC5xWUMgXKyC1dbpGbO98ccrqt221yTaNrpMh5RaR49uKaCmctLL069pJ9e/fZr3YIM44fPy7vv/O+jBo6SmpWqCnlCpbzpZjcMXJV3avsWxShkoLDmUNYSOGPdX/I5Ccmy6C+g6RxjcaqxPaoYlLxbMWlQJYCAVOhrIUSyCBbcSmTv4yULVBWhS82b6zUqlhLflz1o/1qhwzEgQMHpN/N/aRkzpKOFM4i"
    $s &= "hIUUvFi/dr289dpbcsXFV0iZfGX8RhY+X3nZlfL+2+/LJx98Ih8v/NiXsDamT50uA+8YqILXoHoDKV+ovI8U6lap60ghE+Dp/z4txS4o5kjhLELYScFg9szZOuJ7SYHP7wTh8Dl16pScPHlSZr84Wy2O0nlLO1LIJEgXUmjSxs7ucAaRYaTw+qzXA5LC/Dfn21mTxIkTJ+SBvg+oEDpSyBxIEykQd3BhSTk04gk7u0MIYMA8duyYfTnViChSACu+WiGXlL5Ealas6UghEyDVpEA0YdGKcqDP/XrAikPoOHLkiCz5dImu8j391NPy0L0PyW9rfrOzhYyIIwXQs2NPubTMpfLjd44UzjRSRQoEIhUqL4cmPKMnKzmkDt8s+0bKFy4vF0VdJG2vbCsTx06Uzq06y58b/rSzhoSIJIW9e/dq7MPunbvtrxwyGCGTQr4yEl+ymhx+fpadxSFE3Hv7vfL4I49rm7es11Ju7Xqr3N/nfnngrgfsrCEhIknBIfMgJFLIWlh254mVwxOm2V8HjT1xe2TxosWy6c9N9lfnFDb/tVluaHeDrPl5jdzR4w6pXbm2dG7dWT56/yNpclkTdc6nFhFFCgcPHJQ98Xvsyw5nEEGTwrMzZXfWInJw8Ej7q5AwdfxUickbo1Gt5zqG3jNU3n3rXWnbpK1OH5rWbCrXNr5Wbup0k501JEQMKaxbs04ri3mUGSMZN23cJMMfGC6vzHjF/irsMEu23pRRCJYUTqxZJ0fnLRQ5kbhslDeYkY14lsvKXyal85VON1Kw2y2YciSFnTt2ysP3Pywvv/CynDju/5NyqXm+XTa7X4ntaVarmbbJgN4DpFX9VtKmcRuJj4v3yxcqIoYUHh3yqOTJkkeFDtMpEPA1bN+2Xf7e/neS"
    $s &= "ie937dzlu4dlzl27diV7H9/t3hXYf3Hq5CnZ8fcOnc8VzFJQTbeMAntCvv/2e/nv6P/qikztSrU19Lh5reay7Itl8sf6P0ISwtQgWFIIhMOHD2twW52L6sjzk5+3v/YDCle/an2JzhWt73rh6Re0T+y+Iv2z4x/7dj/s2L5DndRMQ+pXqy+1KtXStkO5+t3ST7764iv14ttKmBywYjHj82bJK1dfcbXKhF4/eFAWzl+oz768wuXSok4L3cdDfhv0FdHAy79crhvLLi9/uU4LSKy2Nbq0kcx8bqb89utvvn5FF54Y8YTK38A7B8qvq3+1HxsyIoIUaECsBEKeafAtm7bYWRRTxk3RsOpWDVrpfolACWbt1LKTNj745+9/pM+NfaRJjSbSumHrRPlbN2itSsYzibL0YtGCRcrQbZq00ShLlIJODzcYhYYPGq7OJdoQpURZCOoixeaJVYJCEJ956hn79pDw3YrvVECHDhgqQ+4eIls3b/X7PrWkgMLRX4Srs7lq/KjxdhYFwj9vzjw1kbEQjOw0rN5QTWW7v5AP0tuvvW0/SjHvjXnSsn5LqVikosoTIdoxeWK03fjL3hraDoV8+L6H1UINBihm0fOLSqncpaRD8w4StztO/vzjT+nRrofKBony85f2QuZs0FcQAOUyfUqijGwJ4N4SF5aQqiWqyqznwueoPbOkkL24vP164M7zYsE7C+Q/Mf/R+++88c4kAzWen/K8KmjVklWlVK5SKqTehLJccckVyqqGWA4fOqybezC9qkdX97uPTq5YtKI0r91crm97vaz6ZpVaBs9OeFauu+o6XQ4iwpK81IdOGzlkpI5EKNOw+4ap4jISMsp5U51KdeS6FtfJjGdmyPcrv5dDhw7Z1QkIdiminLQduxWx"
    $s &= "EPr07KPTFnamMkJt3rRZvdJ8F5UtSu6+5e6gHHNHDh+R1T+sllnTZ0n3a7vr6FmucDm10KqVqibReaOlQtEK0viSxvLgPQ/Kpx9+qu1g6p8SKWCVQSo4w65ueLXuZUGJKhSuoKHwXjBFfOOVN/R55DMWgkmqxFb/msS+mdt73u73vI1/bJS+N/WVsvnLah/TduzoxEJZsWyFxMXF6RLfyMEjtc/M3pxqJarJ6GGjA5rkWDorl6+UEYNHSOVilbWcyBiyhBxcVu4yicoRJeUKnC63kROsEwOsCVYMjCwh62wgnPX8LLUMIF76g3KbwZTP4cIZJQU6tl3TdqqUjz74aKLEdUZEYhJgWe5n3p4SEDoa3mygMh1R96K68kMSv/YDOdzW/TYVUjZfITRPjnxSln621M+HgWCjZBAGz/TWh/dhjja6pJFE54zWTuZ5JCNkMD0jkloWeWL0Gvf27t5bpwPJgSXYZjWbqaAhfAj1siXLVDgDARO5SlQVJRAEdc3qNXYWH3DgsscE4aadqR/k+uIzL8r/pv1Pfvr+J23Xmc/OlP639tfRnbJ7R++USGHyk5OVaBgFTf+g7B0D7H3AochGOr6nrt52Nm1NP9mJ57Ju/+XiL33PWvr5Ut21y4hr3tn16q5Jbr9f+8taNfXVgshXWmUBJd22dZtfvulTpqsC0w5eWeN/2pD3UCb63ysrlIOBwmD196uVOIzcYUHZIFBpypNTfLKUVmdicjijpMAGJ5QHc43RiMScjM8IHQnBMIIXLCmAhfMWaoeZe2nwatHV1FsbCCuWrlDBQRDIP/6x8XI0icCaQwcPqQK/OvNV3yhphIGRqEf7HjLh8QlqqmI1MG9+bvJzusHrrp53SdmCZbXe3vtQlG7XdJO/NgQWVEZx"
    $s &= "TGiEhnsQDkaSlMCIjEAiSFgPNjDP58+drxvWaF+IGtMaS2DD74F/+4Dt0y9MfSGhvh5lSIkUnpv0nE7fKhWrpP3BPUmRQvzueLXmtmzeotMX6mvair+Y2rSt5rGSl1yZehLsZqY4tAVTg6QIwYBpLe1gdudyP33qxYvPvig1ytRINDjwmWkk7frzjz/LsIHDfNMCEu//YP7pH6sdO2KsygP3MaCwzBgI1AsZpZ/OWlKgsTGx7rjhDnUcwbyTnpgkt99wu9zQ9gbp0aGHnuZjOiYUUsAku+e2e3wjsQptnhidXuzb4796geDA3JhnjEpMD4IBTh2vWcvz7+l9j53ND1gknEeAoJqRxNxfLFsxad+svezds9e+TRWK55uRDi/z/n377WyJAEEZq4SdprYjCoca+0h8o2jOaJk2Mbg4AgTTKKvWPwVSABAtIyEWCXVPihS8YNcsVpeRGe795adf7GwBwRQHS8nbR1g7wQD5MSY7gwtOXDta8N257/qRHIMC5WNK6MUvP/4iT49/WpMdnv/I/Y8kTH3yxEiL2i2SJGOA5czgyUlX4cIZJQU6+uUZL9tZ5eiRo3Jw/0EdkVFYRi0zfSC+O1js379fOjTr4EcMCFXX1l39FOqj9z5S5w0szvw8WK+zTQpKKBODIxSmIa/+71UVIi8x8Lw3X0m83Eb0mmk/FBFTPxAgHZxcnHXAPB1rTEfXQuW0ft7yUU+eq9bUv/Ne5tNYJcGAKYWZS6vCBUEKgDl8jbI19L2pIYVqJasF7YMZ3G+wz/rjb7druwX0/AcCBMC7TP0giGsaXqOWkhfImJED2gDCTmkq6AWKzn28hzbBN5YUGDBYVcJKChfOKCnwOZjVh61btmpDY8IxF0xqeTAQWF1A0I3ZTcMzKg7q"
    $s &= "N0g7l7k08eKU5cb2N8rPP/1sPyJJBCKFQPPBpIDpjtcaJTeCZ8oGaRiwBIXZaMxUFL1Diw7S67peMubhMer74Dl48HGmYQ3hAKweU12fjaCRIEfugTgAzkljfZCPNqatgwVth5IZayFYUsAPYvxE4SQF2teQD/fizxk3apydLUnQB/SFmepB3hcVv0hJzYtOV3XyyQFt0b5p+5CUllUOY5FQP8rJlCOYOoYDEUEKgIbDx8AcGe96KDiw/4D07NDTN2IYBbyvz33S9PKmUjJHSV1hCHYEMUgrKQCchESjMZc3ZcNqgcwMWAr1+iAYnXEC4kRk3utNCC3fkS6OvViVzySuIeRg4+8bddpgfCI4FhlVQwVKZvo1s5HC6IdG+/l8qOMbL79hZ0sWEKepH+9mpYRVES9sUsB5TrxEsMBX4T2AiPdAQDitP3z3Qzl2NPBqW7gQMaTAtIG5FKN+UnEKyWHjho2+EdM0PAKDsuHc+fKz097qYJEepAB++O4HnyWjz8kb62d+Erlm5sW8g7MrGa1Wfr1Svln+TdBp+dLlGgQEWOY1c2FS9VLV5YtPvvCUKjikJk4ho0jhrl53+e5D6aijd1UiGHA4UJHzi/j6hucRm+JFWkkBeb7y8it9x9p55ZP+ZuXD9gWFExFDCqzPYxKzbBhKg3vBkhpLYt4lLjp50thJdtagkF6k8Pe2v9UCMpYMSmpIAVMfa8aQGUKX1DHqoQBnIqcv80zey7F4qUFmJgWWTjVO4N86svJhgtaCBUuB17e53qewlIMjAr1IKymAzz/+XC0DdMJYsj49uaC4BtHhjMdfFG5EDCmkF+ylNASTaEavuR4s0osUmLZwarV37mpIYd++fRokZcqL4hHoktbwZVZ6fMqcyUkBc93rE0ot"
    $s &= "KUCu639bb2dLEXf2uNP3/nCRAmBVgjJTR2TBDAQqp3lLKzm0u7KdrAvzr2Kfk6RgBBjPP155nDx4kIMNaTVIL1IAEx+fqEuS+hzP9IHAKV2d+Hd1IL0shUghBcgPP4g3TiE1pECbsiRL3EAoIMYE095YceGYPnjBUjrLxPhDqpSoohGpRr5MXA/7III52zS1iFhSYLmShnlp+kv2V0li1cpVGvqLgOBYxBFnlgNZriOqMKkQ6kBIL1JA8AldZjRC6CmXmfvjAWf50TjMEE7K6V2dSA2Ie8DxxjOZb/PO5UuW29lSRLhJgWVT4gVM/UMhBVZizLSDRDkDLfcmBwLUTDuRKC9hzV6kJyl4wcao26+/XZ/ttRp4B/4gfE3hQMSSwrcrvtUOJzosGLAngBDhIlmLSMcWHXW9t3e33lIi+2mztEKRCrpmHKzCpRcp4C8haEkdn7mi1ZHojRVgTl3s/ATFg8QgtrSeK0Ho76VlExST5yL4bEsPFRlBCgNuH5AqS4FAN2+MCvJmrxykBE4Q964+sBfGHqXDRQoAWVz80WKNkDTWCokBhOjYpKJu04KIJQVGT+4PRpA57JWOwjzHoWcUihBadkAagUPhUOy5s+fajwiI9CIFdlviGUdJKINt/bC5ipBtY9XwHkarYMkrEDas36Cbw0zdeTdm6c8/hGZeZwgp9B6QKksB4Fz0OWlzlVLrIdglPt79YP8H/Xw9bAxb++tav3xpIQX6EJJJabv310u/1iVlIwO8h7galtvTGxFJCiMGjVCmhDmfGvOU/bUfMMPZysq7GDUohxfsCkRwvCMRChrMiBKIFIKNaDTAyYijk3spA/Pe33/73S8PAoOgGeEkb6i/kBVoWoQDzRv/gOLh5WbfQbDICFLgnABv/6Ac"
    $s &= "wSoD8S2+EbZAObUG1/wSeG+BDeJhIGPj5IVUWtZtaWdLEykQQEfsw7RJKYeWE2RnAth4D6HywbZDKIgoUmB5jn39dBKMScMsWbzEzubDti3bpOd1CUFLdCgmeqDDXtmsZCL7jHATQMR++ORgkwLPQIBDAbtBzW5DlN0mLQP2Y5DHJ6C5S6kwBjONQMlx1gU6FYr4D6+JDUnw3GCdrpw0ZDzzwZICo56ZuqRECgAT3rv6QJAWu1eDAcvQXrmj3ThzI25Xykt7BLeZ/qWsREcS9GXDJgWc1uaQlZSApQDx1KtWL8nDgwzY5m0IDnlmf1Aw+19CRYaQAhuQHnvoMT/hMxXjx10ICU2qcgg9HUhwkcb/eza3EA5qm3IG7NtnHwN5NG+OEjoyBgLzMmVhT/wCjc+uxeQ2p9ikoApboJyGHKcU+44FQz5GTAQJz3JyUw/OA0B4vPNKVeCWnTToCMW3lykZZZk6MdUiGnTs8LF+3wP2L2AZefuG/+tVrac7EVk/12efPP1s3sOqCNuq2V1pzPNgSQErjLm5Ttfyxkq9KvXk15+SDs754N0PdLcsK0W8h/fhF2IbPGWjLQl9Jxyc8ym8hIFcEWdgnIX0EeUcNzL5cGfe6d3oVPSCojLu0cT34JtiW7pvGvbvr5cFS6qQgp67kD1K7rvzvmSjarEOjaVA308dN9XOki4ICykQYIHCfLzgYw0rRcDs7aUmmYg61pBnPD1D8zNaktjQc1Wdq9Q7Tmd6lZZrNKb38BCUgD3vjCxGeMzIakakQNGQCBZLdHYZMYsRPsKMA+1ctEmBhNJCXIQis3Ua4WCDFzDHbWHd0CY8Hz8HdSFvSuDgDiMQ5n0II+2HAM+cNlPmvDRH244zDDGd+Q4lQOCSwntvv6eHfjKi"
    $s &= "ciiqUVYS7VepaCXdWciz8d5DXvgfKL+3zegfpkLUke2//DUOUw4poS3ML5J77yMwiKg9/BmByJRlQQJ7vBF/5n5Mb5PYkk572kfi8U4Ul4GBulEnnsUKTCAzn/waAv7vuRcoLPN370iO1crUk+3oXn+P9kmuUnowCwewQFS0BVu1A/mAuIYFwv0sP7K3x7b+kB9+sNls5aYeTDPZGBUOpDspEAGG8w9nUHTuaA1NpqNo2OQSTM5P13OegjfB0HZeGo8zF2x/AtMAIhaLZE/4pepE9+WI0g4YcNvpdWYEF+djwfMKJrqHlYnC5xeWqJxRevyX7YC0SYEOI+qSuR6jbVSuKKlTuY6O5ph6nJajBFkgVuvFKMuGJnOiU0rAouHgFIJnKKtu9f5XORB0ysx1RlUsA04g4pQnzgQM5FPwgqkWoyP+F/ZP8Bz12+SOUeLhF8DpD+3PC4upknEIDnX1bubhL3VmhYTYB0ZzhJyVH6wi9muw/Gu3M+1fOaqylpcR2S4vy3MckoOccI8pF4k+4BrE0KV1F1VGG1gSHGHGahV5uY/8DEZsKGMAQsE5YYs8RmbZXMbUzd7gNGbYGA0qQ2ZtuSEVyFpAv2fgMntPsNpsnDxxUo/NM4fv0Ib4dYixgIBxsjIwUVbyYFEyuK1asUrvDQfSnxQOH1GzE5MV4YUg0jsNvH2gjpr2QRmYpXiXGRXte0yiXExlDFBsphnJ3cN3rBdzAKcXNikwUnJiEMDyYGmRUYsTkthfAQlgXvOsUQ+OCnqLtg0UhqhGotuYH3uFkqg33sMSFlGSqTn5mpOHiA2459Z79Dns8ccfM6jvIOl/S389rRiwAkAdIR9DSFgKBPv06tjLdy4DS670CwLOFNBuX5P4rm+vvjp3trcnA6wI"
    $s &= "8iBXkBJtSfk4fIbTsFhmTQmM7pSD+1BYyK7weYW13fiLkkPuKB4ylhQ44Ia9FXYd7PrwLhJnMiZ1BsScl+doHSBR2pEyFcxaUC0x/ubPkl/LSmxNIL9QeiPdSeFcQiBS8PoFIEhMTnMiEP+zz+H4icQCHyqYinAqNY4vDhJhLZvE/1g/vI8Ar7QA09aUO6kwcAiEOArzbnwfW/7aEvSZDKkBdcdnZMoG6YRCsJAq93LwiWk3b/thlcTHx8spSdl6Sw9QHywZHNt2eUgEKVFWpjrHjidv8aUHHCmkASmRgoNDJMKRQhrgSMHhbIQjhTTAkYLD2QhHCmmAIwWHsxGOFNIARwoOZyMcKaQBjhQczkY4UkgDApFCqBuiHBwyGxwppAEEqHj3IvA/4brA3ofg4BApcKQQAji5if0E/KKP+Qk3wlLtRCg1m6/IN6T/kJDPKHBwOJNwpBAC2BVI/Dlx82y2Ig49UDIbirAc2G8RzG8+OjhkFjhSCAHsVmObLiG/Qaft/wR90o+DQ2aAIwUHBwc/OFJwcHDwQ1hIAc+7vWuNa+eKR37z5s0ycuRIWb489CPT0wOBth3boC+8+Uyfkbge6ECQSAY/qnP33XdL3759ZffuhCP5Vq1alWxbPffcc0HJLL9u3r9/f7nrrrtk165deu2FF14I+nDZzIawkML69etl1qxZUrlyZSlTpoyULl1aevfurZ0QTCN7gXBu27bNvpyp8Ntvv/mUCAFp166dZMmSRVq1amVnDTvi4uLkjjvukJiYGClbtqzs3Jnw+xE2fv/9d7ntttu0f2JjY6Vly5by1ltvyRtvvCEPPvigNGvWTIoUKSKDBiX8IG2kAjK4/vrrVf5GjBghw4cPl0mTJknnzp2lRo0aSSruJ598ou2H"
    $s &= "zCaFtWvXSocOHWTcuHE6CPDsgQMHSoMGDVTm+WGXSERYSAHQ2Llz51blyJkzp7z33nt2lqAwfvx46datm30502D+/PkSHR2tvx5twGhEve+9916/vBmF7du36/tJf/+d+LgxAwjswgsv1HyPPPKI33dYO1dccYX07NnT73qkgdGe+n399dd+1z/66COpUqWKHDgQ+DRkCIT77rsv6WPs6tSpI7Vr15Z//jl91gSDw2OPPaaE6kjBQlpJATP22WeflQsuuEBHscyIadOmSZ48eSRv3rx+x4chaFg3oVpF6YXUkAKjHMCc/uGHH/T/d955R3r16mXdFVkwBD12bOJDa7Eg9u5NfPYm1lWjRo30PuQPqyEQChcurMq/YMECv+tr1qyRSy+9VK2USESGkwLz7AceeEDNUsxVTrh5+OGHZfDgwfLoo4/6zN2XXnrJJ9jlypVTk3bRokW+53Pv888/L++//76yOiMbeO211/T55Of/efPm6bsQ+hUrVuhzeRcJE/rtt9+W+++/Xx566CE/HwDC8vnnn8vrr78uH3/8sZbxiSee8M1Hf/75ZylQoICWL0eOHNKvXz9ZuXKlCsSQIUM0P+/yAmV78sknZfHixToyz5w50zftmDNnjq9dyMc7GaX4TD05+9Lgu+++k6eeeko+++wzbTPyfPnl6Z9YTw0pYP4CLJ5XX31V/6cvEPhff/1V22fo0KFqudFulJ9rlAWQj7b+8MMPtY589/333/veg4lNu/Dd6tWrtS2p77Bhw+Svv/yP1WNeTv0+/fRTVWbeuWNHwhmJjL7UlWt8z3OYviUF/AjUj3qOGjVK+8iAvgrkO6G+KPV5552n9zLyB0Lx4sX1++zZs8uECRO0HMaXRn96rcdIQoaTAgLC6MN12Bjl"
    $s &= "Hj16tC8vcz8AWaCsXKtfv778+OOPKuwoZfv27aVYsWIyd+5c7cD8+fPrHBrhYoSuVKmS3leiRAkVuooVK+pnBBVFhf3NZ3wfV199tU9wvvjiC30/gpAtWzYVYt7ZqVMnzVOtWjUVJoTz5ptv1mtYC0uWLFEl4nDVAQMG6PWrrjp93PnEiRO1HczoxLtNGRCkrVu3qjnLNRxWlJO24DPp5ZcTzkWcPXu2nH/++VKzZk2t68KFC/V7RqylSxPOKEwNKfAuyAqypK5e4Kdg7kw+FIDR97LLLtPPKC1TKP6HNHgmz+BzVFSUj8ghSa7RV/QJpHfxxRfrtWuuucb3LpQWsxySZaTFPCfPDTfcoMTQsWNHlRWesWfPHu1rFPiPPwKfbAxxmrYgQeS0HQNGoJ9cY3BBvrgPXwv38L6ffvrJzqr+A/NcZAqZ5N5QreLMhgwnBYDwcL1Hjx6+a7A416677jrfNUYXrqG0BozemOtYDxs2JPwmA55e8sHW4PLLL9fPCDJAYbEmACNarly59HsEEyC4PJNrppyMYnw2fgGj6HQ+1gMwwo+ge6cPEADXDSlgoRjhYcoBIAE+Z82a1TcnRRm4hgVh0LhxY732zDPP6OfJkyfrZ5Thzz//9D2HxHcgNaQAQeM/wDncvHlzO6t618nH6AggY/oHYpoyZYp+d9NNN+l39INdJsrPZ4ia9gDUk2vUGyAzpr5YAgCfAASENUQ/8R2kiEUCxowZo9ewUgIBxceqgaBMmUwyUyYvsMogPWCeTcJ5aAOSgqyMPJkEOUA6kYqIIwWUimt0BKMIwswIi3AjQMBLCvZoQDkoD98bUsDpxGjPNVNOOhzv8i+//KKrKcYqwHrAQw+MJZMSKTCaGoExyu1V5i5duui1YEgBa4Qy"
    $s &= "MaJCikx5zHOmTk04STo1pIBlhA+EZ2LN2GDkJp8hBS+wpGgriIL7UTbz/lBIgbphiXANCw1QJiwVpk9McfgOIsXKwBGKFchnppGBgAnPvbQD8kFfmbLR52ZgMaAvaFPKwjTH5C1VqpRvOgTobxLyxVQS+cDHYPLTTmZ5MtIQcaSA4HONDkDgUGyUGgvCLF2mBymgWCxdYZHceuutWi6+Tw0peJUkECnUqlVLrwVDCoBpEyYw0zCmFeY5aSEFM2qiRExRbCRHCrQxxEc98OUw/THvD4UUsJhQcK4ZS8ELnm2eiwVH/5HwLQRSQMgAP4YX69at0yVb85xrr73W9x3tyjIl1iRWCc/FerLbFzC9YNrkBc/GoiAv9XjxxRf9vo8URBwpmCWmokWL+jnX8JozxwRpIYUPPvhA57ItWrTQz8QcADOdSIoUeL/xNtukgDPTCFYgUli2bJleS44UWIkBmLSYz4UKFZKvvvpKld48Jy2kQPsD/BvGmepFcqRg/AWM3MDUnxQKKTA6m/Lga/EC0sX3ZJ7LlMUA/04gpx6kUK9ePdm06fSviAGsQOMT8fpPIH/bOYzD0LyT2A18XQDiJL8N+pK8OClfeSX8v9EQDkQEKSBseMSZyyOwpkMJDsKMxnFHIA7KDQwp4PQJlRS49s033/gEgZEDBybzbD57ScE4mvLly6fXjDJjYXDdkALK17RpU71mSIH38BnBMh7r5EjBRNeZ56A8jGjGx0BKCyl42z0QzNJeIFIwoyPTOEZs438ghUIKwDhg8d1gvVBnVjuwEjD1mzRpot8THIR5j+LfeeedieIQAO2KA7Z169Z+QVwMHjyHqYdZPWFK1r1794DBTMbJSqIcgP5GvimjV8aYSpCPZyXX9pkZYSEFPMHMUYkI"
    $s &= "Y1WA+RjebVYQWPPF9EU4qlevrmvheM3xDyBwCAidDRhh8QATdceIzZwV0IEoB1MIlJboNBgeAcF6qFu3rjp7sDBY3TACwWiDw6pkyZLqeGK5imt4ss01BIxRByXg+cxxMdNxVKL8PPfGG29UUxdnEvdQP96F+ci8k9US6gJ5ffvtt/pulqt4N0TFfTjlsEawGACWQPny5fU+nFeMfjNmzFDB5RpzXYQPwsFKYCRCkbFgEFqIrm3btkpolJ36kMgfCCaikf4xCbLGJLZDf5lj8376rGrVqkrQ3mVElg9pKywmRl76DbMc5ejatavWkQA02or34IikjakndaPexhKCqLDACAgjf5s2bZR8kRNAGzds2FDbAFmh75MKR2a5EYcszlPkj4GHJVzeT/lMPAZ/aUNWMaZPn+57Fn3Ae+lb08/II32MXCG/yCb1R955dp8+fbSPA01nIgVhIQU7rh6YuHo6yqwNe+PsvddC2TdhnmHAc0xe/tqfvc8239nXDAhC8i5F4UwznW2Xx3zmflOeQHVhFDfk5gX3mLymXIGuAYQVLzkw1xBUrpPf69+w+8GA+7z5zOdA6/Y801tfb1kMUPyNGzf6rqPcKBgw/W5gnmfLgQ3ae8uWLYneZcD9tiXoBfUxZIN1gK8Df4KtsKbNvH1nEKiugPBnM31l2ggR82xvdGOkIiyk4ODgELlwpODg4OAHRwoODg5+cKTg4ODgB0cKDg4OfnCk4ODg4AdHCg4ODn5wpODg4OAHRwoODg5++D8bEYSTuWkTBwAAAABJRU5ErkJggg=="
    Return "data:image/png;base64," & $s
EndFunc

; ============================================================================
; CMR_FULL_FUSION_V1 - API helpers Dispatch <-> moteur CMR inclus
; ============================================================================
Func _CMR_RunFromData($sData)
    $g_bCMRRunning = True
    $g_sCMRStatus = "CMR : préparation queue..."
    $g_sCMRLastJSON = '{"status":"running","message":"preparation"}'
    _AuditLog("CMR", "RunFromData raw=" & StringLeft($sData, 500))

    _QueueClearRows()
    _ClearResults()

    Local $sTxt = StringReplace($sData, "@@BL@@", @LF & @LF)
    $sTxt = StringReplace($sTxt, ",", @LF)
    $sTxt = StringReplace($sTxt, "|", @LF)
    Local $hQ = FileOpen($g_sCMRQueueFile, 2 + 256)
    If $hQ <> -1 Then
        FileWrite($hQ, $sTxt)
        FileClose($hQ)
    EndIf

    Local $sSplitData = StringReplace($sData, "@@BL@@", Chr(30))
    Local $aGroups = StringSplit($sSplitData, Chr(30), 2)
    For $i = 0 To UBound($aGroups) - 1
        Local $sBlock = StringReplace($aGroups[$i], ",", @LF)
        $sBlock = StringReplace($sBlock, "|", @LF)
        If StringStripWS($sBlock, 3) <> "" Then _QueueAppendBlock($sBlock)
    Next

    If UBound($g_aQueueBlocks) = 0 Then
        $g_sCMRStatus = "CMR : queue vide."
        $g_sCMRLastJSON = '{"status":"error","message":"queue_vide","results":[]}'
        $g_bCMRRunning = False
        Return False
    EndIf

    $g_sCMRStatus = "CMR : génération en cours..."
    $g_sCMRLastJSON = '{"status":"running","message":"generation","results":[]}'
    Local $bOk = _RunGenerateQueue()

    $g_sCMRLastJSON = _CMR_BuildResultsJSON($bOk)
    $g_bCMRRunning = False
    If $bOk Then
        $g_sCMRStatus = "CMR : terminé. " & UBound($g_aBL) & " BL généré(s)."
    Else
        $g_sCMRStatus = "CMR : terminé avec erreurs. " & UBound($g_aBL) & " BL généré(s)."
    EndIf
    _AuditLog("CMR", $g_sCMRStatus)
    Return $bOk
EndFunc

Func _CMR_BuildResultsJSON($bOk)
    Local $s = '{"status":"' & ($bOk ? "done" : "partial") & '","running":' & ($g_bCMRRunning ? "true" : "false") & ',"message":"' & _JsonEscape($g_sCMRStatus) & '","queuePath":"' & _JsonEscape($g_sCMRQueueFile) & '","results":['
    For $i = 0 To UBound($g_aBL) - 1
        If $i > 0 Then $s &= ','
        $s &= '{"index":' & $i & ',"bl":"' & _JsonEscape($g_aBL[$i]) & '","company":"' & _JsonEscape($g_aCompany[$i]) & '","carrier":"' & _JsonEscape($g_aCarrier[$i]) & '","delivery":"' & _JsonEscape($g_aDelivery[$i]) & '","html":"' & _JsonEscape($g_aHTML[$i]) & '","pdf":"' & _JsonEscape($g_aPDF[$i]) & '"}'
    Next
    $s &= ']}'
    Return $s
EndFunc

Func _CMR_StatusJSON()
    If $g_bCMRRunning Then
        Return '{"status":"running","running":true,"message":"' & _JsonEscape($g_sCMRStatus) & '","results":[]}'
    EndIf
    If $g_sCMRLastJSON <> "" And $g_sCMRLastJSON <> "{}" Then Return $g_sCMRLastJSON
    Return '{"status":"idle","running":false,"message":"' & _JsonEscape($g_sCMRStatus) & '","results":[]}'
EndFunc

Func _CMR_MailByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aBL) Then Return False
    _CreateMailForResult($idx)
    Return True
EndFunc

Func _CMR_EdocByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aBL) Then Return False
    Return _UploadEdocForResult($idx)
EndFunc


; ============================================================================
; CMR_FULL_PATCH_V4 - batch / edit / rebuild PDF exposes au HTML complet
; ============================================================================
Func _CMR_IndexArray($sIndexes)
    Global $g_aBL
    Local $aOut[0]
    $sIndexes = StringReplace($sIndexes, ",", "|")
    Local $aRaw = StringSplit($sIndexes, "|")
    For $i = 1 To $aRaw[0]
        Local $n = Int(StringStripWS($aRaw[$i], 3))
        If $n >= 0 And $n < UBound($g_aBL) Then
            ReDim $aOut[UBound($aOut) + 1]
            $aOut[UBound($aOut) - 1] = $n
        EndIf
    Next
    Return $aOut
EndFunc

Func _CMR_MailMany($sIndexes)
    Global $g_sCMRStatus, $g_sCMRLastJSON
    Local $a = _CMR_IndexArray($sIndexes)
    Local $ok = 0
    For $i = 0 To UBound($a) - 1
        If _CreateMailForResult($a[$i]) <> "" Then $ok += 1
        Sleep(150)
    Next
    $g_sCMRStatus = "CMR : mails créés " & $ok & "/" & UBound($a)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $ok
EndFunc

Func _CMR_EdocMany($sIndexes)
    Global $g_sCMRStatus, $g_sCMRLastJSON
    Local $a = _CMR_IndexArray($sIndexes)
    Local $ok = 0
    For $i = 0 To UBound($a) - 1
        If _UploadEdocForResult($a[$i]) Then $ok += 1
        Sleep(250)
    Next
    $g_sCMRStatus = "CMR : EDOC upload " & $ok & "/" & UBound($a)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $ok
EndFunc

Func _CMR_EdocMailMany($sIndexes)
    Global $g_sCMRStatus, $g_sCMRLastJSON
    Local $a = _CMR_IndexArray($sIndexes)
    Local $mailOk = 0
    Local $edocOk = 0
    For $i = 0 To UBound($a) - 1
        If _CreateMailForResult($a[$i]) <> "" Then $mailOk += 1
        Sleep(150)
    Next
    For $i = 0 To UBound($a) - 1
        If _UploadEdocForResult($a[$i]) Then $edocOk += 1
        Sleep(250)
    Next
    $g_sCMRStatus = "CMR : mails " & $mailOk & "/" & UBound($a) & " - EDOC " & $edocOk & "/" & UBound($a)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $edocOk
EndFunc

Func _CMR_EditByIndex($idx)
    Global $g_aBL
    If $idx < 0 Or $idx >= UBound($g_aBL) Then Return False
    Return _OpenPreviewEditorByIndex($idx)
EndFunc

Func _CMR_RebuildPdfByIndex($idx)
    Global $g_aBL, $g_sCMRLastJSON
    If $idx < 0 Or $idx >= UBound($g_aBL) Then Return False
    Local $ok = _RebuildPdfByIndex($idx)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $ok
EndFunc


; ============================================================================
; CMR_STABLE_WEB_ENGINE_V2 - génération web CMR sans queue GUI fragile
; ============================================================================
Func _CMR_STABLE_ParseGroups($sData)
    Local $aRaw = StringSplit($sData, "@@BL@@", 1)
    Local $aGroups[0]

    For $i = 1 To $aRaw[0]
        Local $sBlock = StringStripWS($aRaw[$i], 3)
        $sBlock = StringReplace($sBlock, ",", @LF)
        $sBlock = StringReplace($sBlock, "|", @LF)
        $sBlock = StringReplace($sBlock, ";", @LF)
        If StringStripWS($sBlock, 3) = "" Then ContinueLoop
        Local $aNums = _QueueGroupToArray($sBlock)
        If UBound($aNums) = 0 Then ContinueLoop
        ReDim $aGroups[UBound($aGroups) + 1]
        $aGroups[UBound($aGroups) - 1] = _JoinArray($aNums, @LF)
    Next

    Return $aGroups
EndFunc

Func _CMR_STABLE_RunFromData($sData)
    $g_bCMRRunning = True
    $g_bCMRStop = False
    $g_bCMRForce = False
    $g_sCMRStatus = "CMR : préparation..."
    $g_sCMRLastJSON = '{"status":"running","running":true,"message":"preparation","results":[]}'
    _AuditLog("CMR", "StableRun raw=" & StringLeft($sData, 500))

    Local $aGroups = _CMR_STABLE_ParseGroups($sData)
    If UBound($aGroups) = 0 Then
        $g_sCMRStatus = "CMR : queue vide."
        $g_sCMRLastJSON = '{"status":"error","running":false,"message":"queue_vide","results":[]}'
        $g_bCMRRunning = False
        Return False
    EndIf

    _ClearResults()
    Local $okCount = 0
    For $i = 0 To UBound($aGroups) - 1
        If $g_bCMRStop Then ExitLoop
        Local $aNums = _QueueGroupToArray($aGroups[$i])
        If UBound($aNums) = 0 Then ContinueLoop
        $g_sCMRStatus = "CMR : génération BL " & ($i + 1) & "/" & UBound($aGroups) & " - " & _JoinArray($aNums, " / ")
        $g_sCMRLastJSON = '{"status":"running","running":true,"message":"' & _JsonEscape($g_sCMRStatus) & '","results":[]}'
        Local $ok = _GenerateOneBL($aNums, $i + 1, UBound($aGroups))
        If $ok Then $okCount += 1
    Next

    $g_bCMRRunning = False
    If $g_bCMRStop Then
        $g_sCMRStatus = "CMR : stoppé. " & $okCount & "/" & UBound($aGroups) & " BL généré(s)."
    ElseIf $okCount = UBound($aGroups) Then
        $g_sCMRStatus = "CMR : terminé. " & $okCount & " BL généré(s)."
    Else
        $g_sCMRStatus = "CMR : terminé avec erreurs. " & $okCount & "/" & UBound($aGroups) & " BL généré(s)."
    EndIf
    $g_sCMRLastJSON = _CMR_BuildResultsJSON($okCount > 0)
    _AuditLog("CMR", $g_sCMRStatus)
    Return ($okCount > 0)
EndFunc

Func _CMR_STABLE_IndexArray($sIndexes)
    Local $out[0]
    $sIndexes = StringReplace($sIndexes, ",", "|")
    Local $aRaw = StringSplit($sIndexes, "|")
    For $i = 1 To $aRaw[0]
        Local $n = Int(StringStripWS($aRaw[$i], 3))
        If $n >= 0 And $n < UBound($g_aBL) Then
            ReDim $out[UBound($out) + 1]
            $out[UBound($out) - 1] = $n
        EndIf
    Next
    Return $out
EndFunc

Func _CMR_STABLE_MailMany($sIndexes)
    Local $a = _CMR_STABLE_IndexArray($sIndexes)
    Local $ok = 0
    For $i = 0 To UBound($a) - 1
        If _CreateMailForResult($a[$i]) <> "" Then $ok += 1
        Sleep(150)
    Next
    $g_sCMRStatus = "CMR : mails créés " & $ok & "/" & UBound($a)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $ok
EndFunc

Func _CMR_STABLE_EdocMany($sIndexes)
    Local $a = _CMR_STABLE_IndexArray($sIndexes)
    Local $ok = 0
    For $i = 0 To UBound($a) - 1
        If _UploadEdocForResult($a[$i]) Then $ok += 1
        Sleep(250)
    Next
    $g_sCMRStatus = "CMR : EDOC upload " & $ok & "/" & UBound($a)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $ok
EndFunc

Func _CMR_STABLE_EdocMailMany($sIndexes)
    Local $a = _CMR_STABLE_IndexArray($sIndexes)
    Local $mailOk = 0
    Local $edocOk = 0
    For $i = 0 To UBound($a) - 1
        If _CreateMailForResult($a[$i]) <> "" Then $mailOk += 1
        Sleep(150)
    Next
    For $i = 0 To UBound($a) - 1
        If _UploadEdocForResult($a[$i]) Then $edocOk += 1
        Sleep(250)
    Next
    $g_sCMRStatus = "CMR : mails " & $mailOk & "/" & UBound($a) & " - EDOC " & $edocOk & "/" & UBound($a)
    $g_sCMRLastJSON = _CMR_BuildResultsJSON(True)
    Return $edocOk
EndFunc

; EDOC_MASTER_FUSION_V1 - EDOC Master Bot intégré au serveur Dispatch
Global $EDOC_APP_TITLE = "EDOC Master Bot V15.6"
Global $CFG_FILE = @ScriptDir & "\robot_v15_config.ini"
Global $TEMP_DIR = @TempDir & "\Temp_EDOC_Robot_V15"
Global $EDOC_WM_NOTIFY = 0x004E, $EDOC_NM_DBLCLK = -3
Global $C_BG = 0xF3F4F6, $C_CARD = 0xFFFFFF, $C_BORDER = 0xCBD5E1, $C_TEXT = 0x111827, $C_MUTED = 0x6B7280, $C_ACCENT = 0xDCFCE7, $C_BUTTON = 0xE5E7EB
Global $g_oOutlook = 0, $g_oNamespace = 0, $g_sHistorique = "|"
Global $g_aDynCtrls[30][6], $g_iDynCount = 0
Global $g_aGrouped[800][15], $g_iGroupedCount = 0
Global $g_idSelList = 0, $g_hSelList = 0

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

; ============================================================================
; EDOC_HTML_BRIDGE_V13 - Interface EDOC 100% HTML avec pont AutoIt
; ============================================================================
Func _EDOC_EnsureOutlook()
    If IsObj($g_oOutlook) And IsObj($g_oNamespace) Then Return True
    _InitConfig()
    $g_oOutlook = ObjGet("", "Outlook.Application")
    If @error Or Not IsObj($g_oOutlook) Then Return False
    $g_oNamespace = $g_oOutlook.GetNamespace("MAPI")
    Return IsObj($g_oNamespace)
EndFunc

Func _EDOC_JsonArrayFromPipe($sPipe)
    Local $s = "["
    Local $a = StringSplit($sPipe, "|")
    For $i = 1 To $a[0]
        If $a[$i] = "" Then ContinueLoop
        If $s <> "[" Then $s &= ","
        $s &= '"' & _JsonEscape($a[$i]) & '"'
    Next
    $s &= "]"
    Return $s
EndFunc

Func _EDOC_ProfilesJSON()
    Local $sList = IniRead($CFG_FILE, "System", "ProfilesList", "")
    Local $sJson = "["
    Local $aP = StringSplit($sList, "|")
    For $p = 1 To $aP[0]
        Local $prof = $aP[$p]
        If $prof = "" Then ContinueLoop
        If $sJson <> "[" Then $sJson &= ","
        Local $acts = IniRead($CFG_FILE, $prof & "_Actions", "List", "")
        $sJson &= '{"name":"' & _JsonEscape($prof) & '","actions":['
        Local $aA = StringSplit($acts, "|")
        Local $bFirst = True
        For $a = 1 To $aA[0]
            Local $act = $aA[$a]
            If $act = "" Then ContinueLoop
            Local $sec = $prof & "_" & $act
            If Not $bFirst Then $sJson &= ","
            $sJson &= '{"name":"' & _JsonEscape($act) & '","folder":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"Folder","5")) & '","keywords":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"Keywords","")) & '","sender":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"Sender","")) & '","prefix":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"Prefix","J")) & '","length":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"Length","9")) & '","docType":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"DocType","Document")) & '","lastUpload":"' & _JsonEscape(IniRead($CFG_FILE,$sec,"LastUpload_","Jamais")) & '"}'
            $bFirst = False
        Next
        $sJson &= "]}"
    Next
    $sJson &= "]"
    Return $sJson
EndFunc

Func _EDOC_WebInit()
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","reason":"outlook_not_open","stores":[],"profiles":[]}'
    Local $stores = _GetStoresList($g_oNamespace)
    Return '{"status":"ok","stores":' & _EDOC_JsonArrayFromPipe($stores) & ',"profiles":' & _EDOC_ProfilesJSON() & '}'
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

Func _EDOC_GroupedJSON()
    Local $s = "["
    For $i = 0 To $g_iGroupedCount - 1
        If $i > 0 Then $s &= ","
        Local $mail = "false"
        If $g_aGrouped[$i][3] Then $mail = "true"
        Local $pj = "false"
        If $g_aGrouped[$i][4] And $g_aGrouped[$i][10] <> "" Then $pj = "true"
        $s &= '{"index":' & $i & ',"action":"' & _JsonEscape($g_aGrouped[$i][5]) & '","date":"' & _JsonEscape($g_aGrouped[$i][8]) & '","dossiers":"' & _JsonEscape(_PipeToComma(StringTrimRight($g_aGrouped[$i][1],1))) & '","mail":' & $mail & ',"pj":' & $pj & ',"pjNames":"' & _JsonEscape(_AttNames($i)) & '","subject":"' & _JsonEscape($g_aGrouped[$i][7]) & '","sender":"' & _JsonEscape($g_aGrouped[$i][9]) & '"}'
    Next
    $s &= "]"
    Return $s
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
    Local $act = _GetJsonValue($sBody, "action")
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

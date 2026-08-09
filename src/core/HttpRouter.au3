#include-once
; ============================================================================
; HttpRouter.au3
; Routage HTTP/API.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func HttpServer_HandleClient($iSocket)
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


            Case "CMR_GENERATE"
                $sData_a = _GetJsonValue($sBody, "data")
                _CMR_RunFromData($sData_a)
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
            Case "CMR_EDIT"
                ; Ouvre la fenetre d'apercu/edition (mise en forme Expeditors
                ; deja integree au HTML genere, champs contenteditable). Le
                ; bouton "Enregistrer" de cette fenetre sauvegarde le HTML
                ; modifie et regenere le PDF sur le disque.
                $sData_a = _GetJsonValue($sBody, "index")
                _OpenPreviewEditorByIndex(Number($sData_a))
            Case "CMR_REBUILD_PDF"
                $sData_a = _GetJsonValue($sBody, "index")
                _RebuildPdfByIndex(Number($sData_a))

            Case "SP_LIST"
                ; Liste des regles transporteur (destinataires + modele mail),
                ; editables depuis l'interface web.
                _SendHttpResponse($iSocket, 200, "application/json", _SP_ListJSON())
                TCPCloseSocket($iSocket)
                Return
            Case "SP_SAVE"
                _SendHttpResponse($iSocket, 200, "application/json", _SP_SaveFromJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "SP_DELETE"
                _SendHttpResponse($iSocket, 200, "application/json", _SP_DeleteFromJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "SP_TEST_MATCH"
                _SendHttpResponse($iSocket, 200, "application/json", _SP_TestMatchJSON($sBody))
                TCPCloseSocket($iSocket)
                Return

            Case "HPE_MAPPING_LIST"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_MappingListJSON())
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_MAPPING_ADD"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_MappingAddJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_MAPPING_DELETE"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_MappingDeleteJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_MAPPING_LOOKUP"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_MappingLookupJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_SUIVI_LIST"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_SuiviListJSON())
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_SUIVI_SAVE"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_SuiviSaveJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_NOTES_LIST"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_NotesListJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_NOTE_ADD"
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_NoteAddJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "HPE_MAIL_SCAN"
                $sData_a = _GetJsonValue($sBody, "mailbox")
                _SendHttpResponse($iSocket, 200, "application/json", _HPE_MailScanJSON($sData_a))
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

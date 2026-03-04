; ╔══════════════════════════════════════════════════════════════════════════╗
; ║  DispatchMaster 2.0 - Serveur Local & Hub AutoIt                         ║
; ╚══════════════════════════════════════════════════════════════════════════╝
#include <File.au3>
#include <String.au3>
#include <Date.au3>
#include <Misc.au3>

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
Global $iFC_StepCurrent = 0

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
Global $bCOMAT_Stop  = False
Global $bCOMAT_Pause = False

Opt("TrayIconDebug", 1)
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
Global $g_sHtmlFile = @ScriptDir & "\interface.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\Interface.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\Interface_v2.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\DispatchInterface.html"
If Not FileExists($g_sHtmlFile) Then
    MsgBox(16, "Erreur", "Fichier HTML introuvable dans : " & @ScriptDir & @CRLF & "Attendu : interface.html")
    Exit
EndIf
Global $g_sHTML = FileRead($g_sHtmlFile)

ShellExecute("http://127.0.0.1:" & $g_iPort)

; Boucle principale
While 1
    Local $iClientSocket = TCPAccept($g_iMainSocket)
    If $iClientSocket <> -1 Then _HandleClient($iClientSocket)
    Sleep(20)
WEnd

; ==============================================================================
; SERVEUR HTTP
; ==============================================================================
Func _HandleClient($iSocket)
    Local $sHeader = ""
    Local $iContentLength = 0
    Local $sBody = ""

    While 1
        Local $sRecv = TCPRecv($iSocket, 2048)
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
        Sleep(10)
    WEnd

    While StringLen($sBody) < $iContentLength
        Local $sRecv = TCPRecv($iSocket, 4096)
        If @error Then ExitLoop
        If $sRecv <> "" Then $sBody &= $sRecv
        Sleep(10)
    WEnd

    Local $aLines = StringSplit($sHeader, @CRLF, 1)
    If $aLines[0] < 1 Then Return TCPCloseSocket($iSocket)
    Local $aTop = StringSplit($aLines[1], " ")
    If $aTop[0] < 2 Then Return TCPCloseSocket($iSocket)
    Local $sURL = $aTop[2]

    If $sURL = "/" Then
        Local $sHtmlModded = StringReplace($g_sHTML, "127.0.0.1:9500", "127.0.0.1:" & $g_iPort)
        $sHtmlModded = StringReplace($sHtmlModded, "127.0.0.1:8080", "127.0.0.1:" & $g_iPort)
        _SendHttpResponse($iSocket, 200, "text/html", $sHtmlModded)

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

            Case "COMAT_STOP"
                $bCOMAT_Stop = True

            Case "BATCH_CP"
                $sData_a = _GetJsonValue($sBody, "data")
                _Batch_Mails_CP($sData_a)

            Case "save-network-state"
                $sPath_a  = _GetJsonValue($sBody, "path")
                $sState_a = _GetJsonValue($sBody, "state")
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

        EndSwitch

        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok"}')

    Else
        _SendHttpResponse($iSocket, 404, "text/plain", "Not Found")
    EndIf

    TCPCloseSocket($iSocket)
EndFunc

Func _SendHttpResponse($iSocket, $iCode, $sContentType, $sData)
    Local $sStatus = "200 OK"
    If $iCode = 404 Then $sStatus = "404 Not Found"
    Local $bData    = StringToBinary($sData, 4)
    Local $iLen     = BinaryLen($bData)
    Local $sHeaders = "HTTP/1.1 " & $sStatus & @CRLF & _
                      "Content-Type: " & $sContentType & "; charset=UTF-8" & @CRLF & _
                      "Content-Length: " & $iLen & @CRLF & _
                      "Access-Control-Allow-Origin: *" & @CRLF & _
                      "Connection: close" & @CRLF & @CRLF
    TCPSend($iSocket, StringToBinary($sHeaders, 4) & $bData)
EndFunc

; ==============================================================================
; UTILITAIRES
; ==============================================================================
Func _GetJsonValue($sJson, $sKey)
    Local $aMatch = StringRegExp($sJson, '(?i)"' & $sKey & '"\s*:\s*"([^"]*)"', 3)
    If IsArray($aMatch) Then Return $aMatch[0]
    Return ""
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

Func _FC_WaitIfPaused()
    While $bFC_Pause
        _Spinner("EN PAUSE...")
        Sleep(500)
    WEnd
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
    WinActivate($hWndEdoc)
    WinWaitActive($hWndEdoc, "", 2)
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
    MsgBox(16, "Erreur", "E.TMS est fermé ou introuvable.")
    Return
EndIf
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    Send("{PGUP}")
    Sleep(150)
    Local $sCommande = $sBouton & " " & $sNumDossier
    If $sBouton = "LOG X" Then $sCommande = "LOG X"
    Local $sInst = _GetETMSInstance($hWnd)
    Local $sCtrl = "[CLASS:TEIEdit; INSTANCE:" & $sInst & "]"
    ControlFocus($hWnd, "", $sCtrl)
    ControlSetText($hWnd, "", $sCtrl, $sCommande)
    Sleep(150)
    ControlSend($hWnd, "", $sCtrl, "{F8}")
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
    Local $sLogErr = ""
    For $i = 1 To $aJobs[0]
        Local $aInfos = StringSplit($aJobs[$i], "~")
        If $aInfos[0] >= 8 Then
            If _Mail_CP($aInfos[1],$aInfos[2],$aInfos[3],$aInfos[4],$aInfos[5],$aInfos[6],$aInfos[7],$aInfos[8],$sLogErr) Then
                $iOk += 1
                Sleep(1500)
            Else
                $iErr += 1
            EndIf
        EndIf
    Next
    MsgBox(64+262144, "Bilan CP", $iOk & " mail(s)." & @CRLF & $iErr & " erreur(s)." & @CRLF & $sLogErr)
EndFunc

Func _Mail_CP($sClient,$sCmds,$sPal,$sColis,$sPoids,$sConso,$sEmailTo,$sEmailCC, ByRef $sLogErr)
    Local $sCheminBase = "F:\CDG\PRODUCT\TRANSCON\Shared\Clients\HPE\Pre-alertes\"
    Local $aCmdsArr    = StringSplit(StringRegExpReplace($sCmds, "[,;\s]*\+[,;\s]*|[,;]+", " "), " ", 2)
    Local $sCmdListe = ""
    Local $iNbCmd = 0
    Local $bFichiersOK = True
    Local $aFichiers[UBound($aCmdsArr) + 1]
    For $c = 0 To UBound($aCmdsArr) - 1
        Local $sCmd = StringStripWS($aCmdsArr[$c], 3)
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
                _Tracker_Update($iTrackIdx, 1)
                If $sCarrier = "UPS" Then
                    _Run_FileClosing_UPS($sNumJ)
                Else
                    _Run_FileClosing_Single($sNumJ, $sCarrier, $sDateG, $sHoraire, $sContact, $sDLY, $sDLYNotes)
                EndIf
                If $bFC_Stop Then
                    _Tracker_Update($iTrackIdx, 3)
                    ExitLoop 2
                EndIf
                _Tracker_Update($iTrackIdx, 2)
                $iTrackIdx += 1
                Sleep(500)
            Next
        EndIf
    Next
    _Tracker_End()
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
    Local Const $sFC_EDS      = "EXPORT_HPE_FILECLOSING_031.eds"

    Local $hWnd = _GetWindowETMS()
    If $hWnd = 0 Then Return
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    $bFC_Stop = False
    $bFC_Pause = False

    ; ── 1. LOG J ─────────────────────────────────────────────────────────────
    $iFC_StepCurrent = 1
    _Spinner("FC-UPS [" & $Num & "] 1/5 - LOG J...")
    ControlSetText($hWnd, "", $sFC_LOG, "LOG " & $Num)
    Sleep(300)
    ControlSend($hWnd, "", $sFC_LOG, "{F8}")
    Sleep(3000)
    _FC_WaitIfPaused()
    If $bFC_Stop Then Return

    ; ── 2. Toolbar EDS ───────────────────────────────────────────────────────
    $iFC_StepCurrent = 2
    _Spinner("FC-UPS [" & $Num & "] 2/5 - Lancement EDS...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    Sleep(500)
    ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
    _FC_WaitIfPaused()
    If $bFC_Stop Then Return

    ; ── 3. FileOpen (retry x3) ────────────────────────────────────────────────
    $iFC_StepCurrent = 3
    Local $bFileOK = False
    Local $iTentative = 0
    While Not $bFileOK And $iTentative < 3
        $iTentative += 1
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
            If _IsPressed("1B") Then Return
            If TimerDiff($iTimer) > 10000 Then ExitLoop
        WEnd
        If Not WinExists($sFC_FILEOPEN) Then ContinueLoop
        WinActivate($sFC_FILEOPEN)
        WinWaitActive($sFC_FILEOPEN, "", 3)
        Sleep(300)
        ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", "")
        Sleep(150)
        ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", $sFC_EDS)
        Sleep(500)
        If Not StringInStr(ControlGetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]"), "EXPORT_HPE_FILECLOSING") Then ContinueLoop
        Send("{ENTER}")
        Local $iWait = TimerInit()
        While WinExists($sFC_FILEOPEN)
            Sleep(100)
            If TimerDiff($iWait) > 5000 Then ExitLoop
        WEnd
        If Not WinExists($sFC_FILEOPEN) Then $bFileOK = True
    WEnd
    If Not $bFileOK Then
        MsgBox(16+262144, "Erreur FC-UPS", "Impossible d'ouvrir le fichier EDS." & @CRLF & "Dossier : " & $Num)
        $bFC_Stop = True
        Return
    EndIf
    _FC_WaitIfPaused()
    If $bFC_Stop Then Return

    ; ── 4. Menu Selection ────────────────────────────────────────────────────
    $iFC_StepCurrent = 4
    _WinWaitSpinner($sFC_MENU, "FC-UPS [" & $Num & "] 4/5 - Menu Selection...")
    If $bFC_Stop Then Return
    Local $hMenu = WinActivate($sFC_MENU)
    WinWaitActive($hMenu, "", 3)
    Sleep(300)
    ControlSetText($hMenu, "", "[CLASS:TEdit; INSTANCE:1]", "1")
    Sleep(300)
    ControlClick($hMenu, "", "[TEXT:OK]")
    WinWaitClose($hMenu, "", 5)
    Sleep(500)
    _FC_WaitIfPaused()
    If $bFC_Stop Then Return

    ; ── 5. Numéro J ──────────────────────────────────────────────────────────
    $iFC_StepCurrent = 5
    _WinWaitSpinner($sFC_INPUT, "FC-UPS [" & $Num & "] 5/5 - Numéro J...")
    If $bFC_Stop Then Return
    Local $hInput1 = WinActivate($sFC_INPUT)
    WinWaitActive($hInput1, "", 3)
    Sleep(300)
    ControlSetText($hInput1, "", "[CLASS:TEdit; INSTANCE:1]", $Num)
    Sleep(300)
    ControlClick($hInput1, "", "[TEXT:OK]")
    WinWaitClose($hInput1, "", 5)
    Sleep(3000) ; E.TMS charge
    _FC_WaitIfPaused()
    If $bFC_Stop Then Return

    ; ── 6. UPS = pas de Carrier ID → première popup = DEF → terminé ──────────
    $iFC_StepCurrent = 6
    _WinWaitSpinner($sFC_INPUT, "FC-UPS [" & $Num & "] DEF...")
    If $bFC_Stop Then Return
    Local $hDef = WinActivate($sFC_INPUT)
    WinWaitActive($hDef, "", 3)
    Sleep(300)
    ControlSetText($hDef, "", "[CLASS:TEdit; INSTANCE:1]", "DEF")
    Sleep(300)
    ControlClick($hDef, "", "[TEXT:OK]")
    WinWaitClose($hDef, "", 5)
    Sleep(500)

    ; ── C'EST TOUT POUR UPS ──────────────────────────────────────────────────
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
    Local Const $sFC_EDS      = "EXPORT_HPE_FILECLOSING_031.eds"

    Local $hWnd = _GetWindowETMS()
    If $hWnd = 0 Then Return
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    $bFC_Stop  = False
    $bFC_Pause = False

    If $iStartStep <= 1 Then
        $iFC_StepCurrent = 1
        _Spinner("FC [" & $Num & "] 1/7 - LOG J...")
        ControlSetText($hWnd, "", $sFC_LOG, "LOG " & $Num)
        Sleep(300)
        ControlSend($hWnd, "", $sFC_LOG, "{F8}")
        Sleep(3000)
        _FC_WaitIfPaused()
        If $bFC_Stop Then Return
    EndIf

    If $iStartStep <= 2 Then
        $iFC_StepCurrent = 2
        _Spinner("FC [" & $Num & "] 2/7 - Lancement EDS...")
        WinActivate($hWnd)
        WinWaitActive($hWnd, "", 3)
        Sleep(500)
        ControlClick($hWnd, "", $sFC_TOOLBAR, "LEFT", 1, 54, 9)
        _FC_WaitIfPaused()
        If $bFC_Stop Then Return
    EndIf

    If $iStartStep <= 3 Then
        $iFC_StepCurrent = 3
        Local $bFileOK    = False
        Local $iTentative = 0
        While Not $bFileOK And $iTentative < 3
            $iTentative += 1
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
                If _IsPressed("1B") Then Return
                If TimerDiff($iTimer) > 10000 Then ExitLoop
            WEnd
            If Not WinExists($sFC_FILEOPEN) Then ContinueLoop
            WinActivate($sFC_FILEOPEN)
            WinWaitActive($sFC_FILEOPEN, "", 3)
            Sleep(300)
            ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", "")
            Sleep(150)
            ControlSetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]", $sFC_EDS)
            Sleep(500)
            If Not StringInStr(ControlGetText($sFC_FILEOPEN, "", "[CLASS:TRzEdit; INSTANCE:1]"), "EXPORT_HPE_FILECLOSING") Then ContinueLoop
            Send("{ENTER}")
            Local $iWait = TimerInit()
            While WinExists($sFC_FILEOPEN)
                Sleep(100)
                If TimerDiff($iWait) > 5000 Then ExitLoop
            WEnd
            If Not WinExists($sFC_FILEOPEN) Then $bFileOK = True
        WEnd
        If Not $bFileOK Then
            MsgBox(16+262144, "Erreur FC", "Impossible d'ouvrir le fichier EDS." & @CRLF & "Dossier : " & $Num)
            $bFC_Stop = True
            Return
        EndIf
        _FC_WaitIfPaused()
        If $bFC_Stop Then Return
    EndIf

    If $iStartStep <= 4 Then
        $iFC_StepCurrent = 4
        _WinWaitSpinner($sFC_MENU, "FC [" & $Num & "] 4/7 - Menu Selection...")
        If $bFC_Stop Then Return
        Local $hMenu = WinActivate($sFC_MENU)
        WinWaitActive($hMenu, "", 3)
        Sleep(300)
        ControlSetText($hMenu, "", "[CLASS:TEdit; INSTANCE:1]", "1")
        Sleep(300)
        ControlClick($hMenu, "", "[TEXT:OK]")
        WinWaitClose($hMenu, "", 5)
        Sleep(500)
        _FC_WaitIfPaused()
        If $bFC_Stop Then Return
    EndIf

    If $iStartStep <= 5 Then
        $iFC_StepCurrent = 5
        _WinWaitSpinner($sFC_INPUT, "FC [" & $Num & "] 5/7 - Numero J...")
        If $bFC_Stop Then Return
        Local $hInput1 = WinActivate($sFC_INPUT)
        WinWaitActive($hInput1, "", 3)
        Sleep(300)
        ControlSetText($hInput1, "", "[CLASS:TEdit; INSTANCE:1]", $Num)
        Sleep(300)
        ControlClick($hInput1, "", "[TEXT:OK]")
        WinWaitClose($hInput1, "", 5)
        Sleep(300)
        _FC_WaitIfPaused()
        If $bFC_Stop Then Return
    EndIf

    If $iStartStep <= 6 Then
        $iFC_StepCurrent = 6
        _WinWaitSpinner($sFC_CARRIER, "FC [" & $Num & "] 6/7 - Carrier ID [" & $CarrierID & "]...")
        If $bFC_Stop Then Return
        Local $hCarrier = WinActivate($sFC_CARRIER)
        WinWaitActive($hCarrier, "", 3)
        Sleep(300)
        ControlSetText($hCarrier, "", "[CLASS:TEdit; INSTANCE:1]", $CarrierID)
        Sleep(300)
        ControlClick($hCarrier, "", "[TEXT:OK]")
        WinWaitClose($hCarrier, "", 5)
        Sleep(3000)
        _FC_WaitIfPaused()
        If $bFC_Stop Then Return
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
        Local $p
        For $p = 0 To 15
            If $aSkip[$p] Then ContinueLoop
            If $bFC_Stop Then Return
            Local $sVal       = $aVal[$p]
            Local $colNom     = Chr(67 + $p)
            Local $bColValidee = False
            Local $iTimeout   = 0
            If $p >= 14 Then $iTimeout = 3
            While Not $bColValidee
                _FC_WaitIfPaused()
                If $bFC_Stop Then Return
                _Spinner("FC [" & $Num & "] Col " & $colNom & "...")
                Local $hWinWait = WinWait($sFC_INPUT, "", $iTimeout)
                If $hWinWait = 0 And $p >= 14 Then ExitLoop 2
                Local $hWin   = WinActivate($sFC_INPUT)
                Local $sTitre = WinGetTitle($hWin)
                Sleep(150)
                If StringInStr($sTitre, "REASON") Then
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
        Next
        Sleep(500)
    EndIf

    $iFC_StepCurrent = 0
    ToolTip("")
EndFunc


; ==============================================================================
; TRACKER VISUEL
; ==============================================================================
Func _Tracker_Start($sTitle, $aList)
    $g_iTrackCount = UBound($aList)
    If $g_iTrackCount = 0 Then Return False
    ReDim $g_aTrackIDs[$g_iTrackCount]
    $g_hTracker    = GUICreate($sTitle, 280, 400, @DesktopWidth - 300, 50, -1, 0x00000008)
    GUISetBkColor(0x2D2D30, $g_hTracker)
    GUISetFont(9, 400, 0, "Segoe UI")
    $g_idTrackProg = GUICtrlCreateProgress(10, 10, 260, 15)
    $g_idTrackLbl  = GUICtrlCreateLabel("0 / " & $g_iTrackCount & " dossiers", 10, 30, 260, 20, 1)
    GUICtrlSetColor(-1, 0xFFFFFF)
    $g_idTrackLV   = GUICtrlCreateListView("Statut|Numero J", 10, 55, 260, 335)
    GUICtrlSetBkColor(-1, 0x1E1E1E)
    GUICtrlSetColor(-1, 0xCCCCCC)
    GUICtrlSendMsg($g_idTrackLV, 0x101E, 0, 80)
    GUICtrlSendMsg($g_idTrackLV, 0x101E, 1, 150)
    For $i = 0 To $g_iTrackCount - 1
        $g_aTrackIDs[$i] = GUICtrlCreateListViewItem("Attente|" & $aList[$i], $g_idTrackLV)
        GUICtrlSetColor($g_aTrackIDs[$i], 0xAAAAAA)
    Next
    GUISetState(@SW_SHOWNOACTIVATE, $g_hTracker)
    Return True
EndFunc

Func _Tracker_Update($iIndex, $iStatus)
    If $g_hTracker = 0 Then Return
    Local $sText = "Attente"
    Local $iColor = 0xAAAAAA
    Switch $iStatus
        Case 1
            $sText = "En cours"
            $iColor = 0xFFCC00
            GUICtrlSetData($g_idTrackLbl, "Traitement : " & ($iIndex+1) & " / " & $g_iTrackCount)
            GUICtrlSetData($g_idTrackProg, ($iIndex / $g_iTrackCount) * 100)
        Case 2
            $sText = "Terminé"
            $iColor = 0x00CC55
            GUICtrlSetData($g_idTrackProg, (($iIndex+1) / $g_iTrackCount) * 100)
        Case 3
            $sText = "Stop/Err"
            $iColor = 0xFF4444
    EndSwitch
    Local $sJ = GUICtrlRead($g_aTrackIDs[$iIndex])
    $sJ = StringMid($sJ, StringInStr($sJ, "|") + 1)
    GUICtrlSetData($g_aTrackIDs[$iIndex], $sText & "|" & $sJ)
    GUICtrlSetColor($g_aTrackIDs[$iIndex], $iColor)
EndFunc

Func _Tracker_End()
    If $g_hTracker Then
        GUICtrlSetData($g_idTrackProg, 100)
        GUICtrlSetData($g_idTrackLbl, "Traitement terminé !")
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
    Local $aJobs = StringSplit($sData, "|")
    Local $aValid[$aJobs[0]]
    For $i = 1 To $aJobs[0]
        Local $aInfos = StringSplit($aJobs[$i], ";")
        $aValid[$i-1] = $aInfos[1]
    Next
    _Tracker_Start("COMAT en masse", $aValid)
    For $i = 1 To $aJobs[0]
        Local $aDetails = StringSplit($aJobs[$i], ";")
        If $aDetails[0] >= 1 Then
            Local $sNumJ = $aDetails[1]
            _Tracker_Update($i-1, 1)
            _Run_COMAT_Single($sNumJ)
            If $bCOMAT_Stop Then
                _Tracker_Update($i-1, 3)
                ExitLoop
            EndIf
            _Tracker_Update($i-1, 2)
            Sleep(500)
        EndIf
    Next
    _Tracker_End()
    $bCOMAT_Stop = False
    $bCOMAT_Pause = False
EndFunc

Func _Run_COMAT_Single($Num)
    $Num = StringStripWS($Num, 8)
    If $Num = "" Then Return
    Local $hWnd = WinGetHandle("[CLASS:TfmBrowser]")
    If $hWnd = 0 Or Not WinExists($hWnd) Then
        MsgBox(16+262144, "Erreur COMAT", "Fenêtre E.TMS introuvable.")
        $bCOMAT_Stop = True
        Return
    EndIf
    $bCOMAT_Stop = False
    $bCOMAT_Pause = False
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)

    _COMAT_Spinner("COMAT [" & $Num & "] 1/5 - LOG J...")
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "")
    Sleep($COMAT_DELAY_M)
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG " & $Num)
    Sleep($COMAT_DELAY_M)
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    Sleep($COMAT_DELAY_LOAD)
    _COMAT_WaitIfPaused()
    If $bCOMAT_Stop Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 2/5 - F3...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    Send("{F3}")
    Sleep($COMAT_DELAY_L)
    _COMAT_WaitIfPaused()
    If $bCOMAT_Stop Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 3/5 - F5 x4...")
    Local $k
    For $k = 1 To 4
        Send("{F5}")
        Sleep($COMAT_DELAY_M)
    Next
    Sleep($COMAT_DELAY_L)
    _COMAT_WaitIfPaused()
    If $bCOMAT_Stop Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 4/5 - F1 + TAB + C...")
    Send("{F1}")
    Sleep($COMAT_DELAY_L)
    For $k = 1 To 6
        Send("{TAB}")
        Sleep($COMAT_DELAY_S)
    Next
    Sleep($COMAT_DELAY_M)
    Send("C")
    Sleep($COMAT_DELAY_M)
    For $k = 1 To 4
        Send("{F5}")
        If $k = 4 Then
            Sleep($COMAT_DELAY_L)
        Else
            Sleep($COMAT_DELAY_M)
        EndIf
    Next
    Sleep(800)
    _COMAT_WaitIfPaused()
    If $bCOMAT_Stop Then Return

    _COMAT_Spinner("COMAT [" & $Num & "] 5/5 - Retour LOG...")
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 3)
    Sleep($COMAT_DELAY_M)
    ControlSetText($hWnd, "", $COMAT_LOG_CTRL, "LOG")
    Sleep($COMAT_DELAY_M)
    ControlSend($hWnd, "", $COMAT_LOG_CTRL, "{F8}")
    Sleep(2000)
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
                For $j = 0 To $iTotalGui - 1
                    _Tracker_Update($j, 1)
                    _Run_COMAT_Single($aValidGui[$j])
                    If $bCOMAT_Stop Then
                        _Tracker_Update($j, 3)
                        ExitLoop
                    EndIf
                    _Tracker_Update($j, 2)
                    Sleep(500)
                Next
                _Tracker_End()
                $bCOMAT_Stop = False
                $bCOMAT_Pause = False
                MsgBox(64+262144, "Terminé", "Traitement COMAT terminé.")
                Return
        EndSwitch
    WEnd
EndFunc

Func _COMAT_Spinner($sTxt)
    ToolTip($sTxt, 0, 0, "Robot E.TMS — COMAT", 1)
EndFunc

Func _COMAT_WaitIfPaused()
    While $bCOMAT_Pause
        _COMAT_Spinner("EN PAUSE...")
        Sleep(500)
        If $bCOMAT_Stop Then Return
    WEnd
EndFunc

; ==============================================================================
; RÉSEAU PARTAGÉ
; ==============================================================================
Func _Net_SaveState($sPath, $sJSON)
    Local $hFile = FileOpen($sPath, 2)
    If $hFile = -1 Then
        ToolTip("Erreur écriture : " & $sPath, 0, 0)
        Return False
    EndIf
    FileWrite($hFile, $sJSON)
    FileClose($hFile)
    Return True
EndFunc

Func _Net_LoadState($sPath)
    If Not FileExists($sPath) Then Return "{}"
    Local $hFile = FileOpen($sPath, 0)
    If $hFile = -1 Then Return "{}"
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

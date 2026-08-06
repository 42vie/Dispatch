; ============================================================================
; Mail_Alertes.au3
; Mails alertes.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

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

; ============================================================================
; Mail_RDV.au3
; Mails RDV.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

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

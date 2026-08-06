; ============================================================================
; Mail_CP.au3
; Mails CP.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

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

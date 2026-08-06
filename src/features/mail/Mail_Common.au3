; ============================================================================
; Mail_Common.au3
; Commun mail Outlook.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

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

Func _SendMailOnlyResult($idx)
    Return (_CreateMailForResult($idx) <> "")
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

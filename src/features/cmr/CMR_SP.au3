; ============================================================================
; CMR_SP.au3
; Gestion des modeles par transporteur (Service Provider) : chargement,
; sauvegarde, regles de correspondance, application des variables au modele.
; ----------------------------------------------------------------------------
; Deplace depuis StateService.au3 (bloc CMR_FULL_FUSION_MODULE_APPENDED) pour
; que le code CMR vive dans src/features/cmr/, comme le reste du projet.
; ============================================================================


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


; ============================================================================
; {{NAME_UPPER}}_Tool.au3
; Genere par tools\NewFeature.au3 le {{DATE}}.
; Outil "{{NAME_UPPER}}" : liste d'enregistrements stockes dans
; Config\{{NAME_UPPER}}_Tool.ini (une section ini par enregistrement, prefixe
; "ITEM:" + identifiant unique). Meme structure que HPE_Tool.au3/CMR_SP.au3 --
; libre a toi d'ajouter/retirer des champs, des routes, ou de changer
; completement la logique : ce fichier n'est qu'un point de depart genere
; automatiquement, pas un contrat fige.
; ============================================================================

Global Const ${{NAME_UPPER}}_INI_PATH = @ScriptDir & "\Config\{{NAME_UPPER}}_Tool.ini"

Func _{{NAME_UPPER}}_Sec($id)
    Return "ITEM:" & StringUpper(StringStripWS($id, 3))
EndFunc

Func _{{NAME_UPPER}}_ListJSON()
    Local $sections = IniReadSectionNames(${{NAME_UPPER}}_INI_PATH)
    Local $sJson = "["
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 5) = "ITEM:" Then
                Local $sId = StringTrimLeft($sections[$i], 5)
                If $sJson <> "[" Then $sJson &= ","
                $sJson &= '{"id":"' & _JsonEscape($sId) & '"' & _
                        ',"label":"' & _JsonEscape(IniRead(${{NAME_UPPER}}_INI_PATH, $sections[$i], "LABEL", "")) & '"' & _
                        ',"value":"' & _JsonEscape(IniRead(${{NAME_UPPER}}_INI_PATH, $sections[$i], "VALUE", "")) & '"' & _
                        ',"date":"' & _JsonEscape(IniRead(${{NAME_UPPER}}_INI_PATH, $sections[$i], "DATE", "")) & '"}'
            EndIf
        Next
    EndIf
    $sJson &= "]"
    Return '{"status":"ok","items":' & $sJson & '}'
EndFunc

Func _{{NAME_UPPER}}_AddJSON($sBody)
    Local $sId = StringStripWS(_GetJsonValue($sBody, "id"), 3)
    If $sId = "" Then $sId = StringRegExpReplace(@YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC, "[^0-9]", "")
    Local $sLabel = _GetJsonValue($sBody, "label")
    Local $sValue = _GetJsonValue($sBody, "value")
    Local $sSec = _{{NAME_UPPER}}_Sec($sId)
    IniWrite(${{NAME_UPPER}}_INI_PATH, $sSec, "LABEL", $sLabel)
    IniWrite(${{NAME_UPPER}}_INI_PATH, $sSec, "VALUE", $sValue)
    IniWrite(${{NAME_UPPER}}_INI_PATH, $sSec, "DATE", _NowText())
    Return '{"status":"ok","id":"' & _JsonEscape($sId) & '"}'
EndFunc

Func _{{NAME_UPPER}}_DeleteJSON($sBody)
    Local $sId = StringStripWS(_GetJsonValue($sBody, "id"), 3)
    If $sId = "" Then Return '{"status":"error","message":"id_vide"}'
    IniDelete(${{NAME_UPPER}}_INI_PATH, _{{NAME_UPPER}}_Sec($sId))
    Return '{"status":"ok"}'
EndFunc

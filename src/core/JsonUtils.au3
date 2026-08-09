; ============================================================================
; JsonUtils.au3
; Fonctions JSON.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

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

Func _CMR_BuildResultsJSON($bOk)
    Local $s = '{"status":"' & ($bOk ? "done" : "partial") & '","running":' & ($g_bCMRRunning ? "true" : "false") & ',"message":"' & _JsonEscape($g_sCMRStatus) & '","queuePath":"' & _JsonEscape($g_sCMRQueueFile) & '","results":['
    For $i = 0 To UBound($g_aBL) - 1
        If $i > 0 Then $s &= ','
        ; Regle SP (transporteur) qui sera utilisee pour le mail de ce BL --
        ; permet a l'interface de signaler si on tombe sur le modele
        ; generique "Autre" faute de regle specifique pour ce transporteur.
        Local $sSpSec = _SP_MatchSection($g_aCarrier[$i])
        Local $sSpFallback = "false"
        If $sSpSec = "SP:AUTRE" Then $sSpFallback = "true"
        $s &= '{"index":' & $i & ',"bl":"' & _JsonEscape($g_aBL[$i]) & '","company":"' & _JsonEscape($g_aCompany[$i]) & '","carrier":"' & _JsonEscape($g_aCarrier[$i]) & '","delivery":"' & _JsonEscape($g_aDelivery[$i]) & '","html":"' & _JsonEscape($g_aHTML[$i]) & '","pdf":"' & _JsonEscape($g_aPDF[$i]) & '","spFallback":' & $sSpFallback & '}'
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

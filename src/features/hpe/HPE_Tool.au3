; ============================================================================
; HPE_Tool.au3
; Outil HPE : lien entre les references client (INC/FXC/ARN) et les dossiers
; internes (J1A/H1A), + suivi (proprietaire/statut) et notes horodatees par
; dossier. Stockage dans un fichier ini (Config\HPE_Tool.ini), meme logique
; que les regles transporteur (CMR_SP.au3) : plus rapide et plus fiable que
; l'automation Excel utilisee dans une premiere version (pas de dependance a
; Excel installe, pas de risque de fichier verrouille ou de process fantome).
; ============================================================================

Global Const $HPE_INI_PATH = @ScriptDir & "\Config\HPE_Tool.ini"

; Extrait la premiere reference INC/FXC/ARN + chiffres trouvee dans une
; chaine (ex: objet de mail), normalisee en majuscules sans espace/tiret.
; Renvoie "" si aucune reference trouvee.
Func _HPE_ExtractRef($s)
    Local $a = StringRegExp($s, "(?i)(INC|FXC|ARN)[- ]?(\d{4,})", 1)
    If @error Then Return ""
    Return StringUpper($a[0]) & $a[1]
EndFunc

Func _HPE_MapSec($ref)
    Return "MAP:" & StringUpper(StringStripWS($ref, 3))
EndFunc

Func _HPE_SuiviSec($dossier)
    Return "SUIVI:" & StringUpper(StringStripWS($dossier, 3))
EndFunc

; Encode/decode d'une note individuelle -- reutilise la meme convention que
; CMR_SP.au3 pour les retours a la ligne dans un champ ini.
Func _HPE_EncNote($s)
    $s = StringReplace($s, @CRLF, "{BR}")
    $s = StringReplace($s, @CR, "{BR}")
    $s = StringReplace($s, @LF, "{BR}")
    Return $s
EndFunc

Func _HPE_DecNote($s)
    Return StringReplace($s, "{BR}", @CRLF)
EndFunc

; ============================================================================
; MAPPING (reference HPE <-> dossier) -- une section ini par reference (donc
; forcement unique), plusieurs references peuvent partager le meme DOSSIER.
; ============================================================================

Func _HPE_MappingListJSON()
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    Local $sJson = "["
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 4) = "MAP:" Then
                Local $sRef = StringTrimLeft($sections[$i], 4)
                If $sJson <> "[" Then $sJson &= ","
                $sJson &= '{"reference":"' & _JsonEscape($sRef) & '"' & _
                        ',"type":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "TYPE", "")) & '"' & _
                        ',"dossier":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "DOSSIER", "")) & '"' & _
                        ',"dateAjout":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "DATE", "")) & '"' & _
                        ',"operateur":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "OPERATOR", "")) & '"}'
            EndIf
        Next
    EndIf
    $sJson &= "]"
    Return '{"status":"ok","mapping":' & $sJson & '}'
EndFunc

Func _HPE_MappingAddJSON($sBody)
    Local $sRef = StringUpper(StringStripWS(_GetJsonValue($sBody, "reference"), 3))
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    If $sRef = "" Or $sDossier = "" Then Return '{"status":"error","message":"champs_vides"}'
    Local $sType = ""
    Local $aType = StringRegExp($sRef, "(?i)^(INC|FXC|ARN)", 1)
    If Not @error Then $sType = StringUpper($aType[0])

    Local $sSec = _HPE_MapSec($sRef)
    IniWrite($HPE_INI_PATH, $sSec, "DOSSIER", $sDossier)
    IniWrite($HPE_INI_PATH, $sSec, "TYPE", $sType)
    IniWrite($HPE_INI_PATH, $sSec, "DATE", _NowText())
    IniWrite($HPE_INI_PATH, $sSec, "OPERATOR", @UserName)
    Return '{"status":"ok"}'
EndFunc

Func _HPE_MappingDeleteJSON($sBody)
    Local $sRef = StringUpper(StringStripWS(_GetJsonValue($sBody, "reference"), 3))
    If $sRef = "" Then Return '{"status":"error","message":"reference_vide"}'
    IniDelete($HPE_INI_PATH, _HPE_MapSec($sRef))
    Return '{"status":"ok"}'
EndFunc

; Cherche une valeur dans les references ET les dossiers du mapping, renvoie
; la paire complete si trouvee.
Func _HPE_MappingLookupJSON($sBody)
    Local $sVal = StringUpper(StringStripWS(_GetJsonValue($sBody, "value"), 3))
    If $sVal = "" Then Return '{"status":"error","message":"valeur_vide"}'

    ; Cas rapide : $sVal est directement une reference connue.
    Local $sDirect = IniRead($HPE_INI_PATH, _HPE_MapSec($sVal), "DOSSIER", "")
    If $sDirect <> "" Then
        Return '{"status":"ok","found":true,"reference":"' & _JsonEscape($sVal) & '","dossier":"' & _JsonEscape($sDirect) & '"}'
    EndIf

    ; Repli : $sVal est peut-etre un numero de dossier -- on parcourt les
    ; references pour trouver celle qui pointe dessus.
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 4) = "MAP:" Then
                Local $sDoss = StringUpper(IniRead($HPE_INI_PATH, $sections[$i], "DOSSIER", ""))
                If $sDoss = $sVal Then
                    Return '{"status":"ok","found":true,"reference":"' & _JsonEscape(StringTrimLeft($sections[$i], 4)) & '","dossier":"' & _JsonEscape($sDoss) & '"}'
                EndIf
            EndIf
        Next
    EndIf
    Return '{"status":"ok","found":false}'
EndFunc

; ============================================================================
; SUIVI (proprietaire / statut par dossier) + NOTES horodatees (empaquetees
; dans une seule cle de la meme section : plus simple qu'une section par
; note, et les notes se lisent/s'ecrivent toujours toutes ensemble).
; ============================================================================

Func _HPE_SuiviListJSON()
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    Local $sJson = "["
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 6) = "SUIVI:" Then
                Local $sDoss = StringTrimLeft($sections[$i], 6)
                If $sJson <> "[" Then $sJson &= ","
                $sJson &= '{"dossier":"' & _JsonEscape($sDoss) & '"' & _
                        ',"owner":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "OWNER", "")) & '"' & _
                        ',"status":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "STATUS", "")) & '"' & _
                        ',"maj":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "MAJ", "")) & '"}'
            EndIf
        Next
    EndIf
    $sJson &= "]"
    Return '{"status":"ok","suivi":' & $sJson & '}'
EndFunc

Func _HPE_SuiviSaveJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    If $sDossier = "" Then Return '{"status":"error","message":"dossier_vide"}'
    Local $sOwner = _GetJsonValue($sBody, "owner")
    Local $sStatus = _GetJsonValue($sBody, "status")

    Local $sSec = _HPE_SuiviSec($sDossier)
    IniWrite($HPE_INI_PATH, $sSec, "OWNER", $sOwner)
    IniWrite($HPE_INI_PATH, $sSec, "STATUS", $sStatus)
    IniWrite($HPE_INI_PATH, $sSec, "MAJ", _NowText())
    Return '{"status":"ok"}'
EndFunc

; ── Notes : cle "NOTES" = enregistrements separes par Chr(30), chaque
; enregistrement = date / auteur / note separes par Chr(31).
Func _HPE_NotesListJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    If $sDossier = "" Then Return '{"status":"error","message":"dossier_vide"}'
    Local $sPacked = IniRead($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "NOTES", "")
    Local $sJson = "["
    If $sPacked <> "" Then
        Local $aRecords = StringSplit($sPacked, Chr(30), 2)
        For $i = 0 To UBound($aRecords) - 1
            If $aRecords[$i] = "" Then ContinueLoop
            Local $aF = StringSplit($aRecords[$i], Chr(31), 2)
            If UBound($aF) < 3 Then ContinueLoop
            If $sJson <> "[" Then $sJson &= ","
            $sJson &= '{"date":"' & _JsonEscape($aF[0]) & '","auteur":"' & _JsonEscape($aF[1]) & '","note":"' & _JsonEscape(_HPE_DecNote($aF[2])) & '"}'
        Next
    EndIf
    $sJson &= "]"
    Return '{"status":"ok","notes":' & $sJson & '}'
EndFunc

Func _HPE_NoteAddJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    Local $sNote = _GetJsonValue($sBody, "note")
    If $sDossier = "" Or $sNote = "" Then Return '{"status":"error","message":"champs_vides"}'
    Local $sSec = _HPE_SuiviSec($sDossier)
    Local $sPacked = IniRead($HPE_INI_PATH, $sSec, "NOTES", "")
    Local $sRecord = _NowText() & Chr(31) & @UserName & Chr(31) & _HPE_EncNote($sNote)
    If $sPacked = "" Then
        $sPacked = $sRecord
    Else
        $sPacked &= Chr(30) & $sRecord
    EndIf
    IniWrite($HPE_INI_PATH, $sSec, "NOTES", $sPacked)
    Return '{"status":"ok"}'
EndFunc

; ============================================================================
; SCAN MAIL (Outlook) -- detecte les references INC/FXC/ARN dans l'objet des
; mails recents, les relie au dossier via le Mapping. Les references HPE
; sont dans l'objet du mail, et les reponses ("RE:", "TR:"...) portent le
; meme objet/reference -- donc on regroupe par reference : une seule alerte
; par conversation (la plus recente, les mails etant parcourus du plus
; recent au plus ancien), avec le nombre de mails du fil.
; ============================================================================

; $sMailbox : nom de la boite partagee a scanner (ex: "CDG-Transcon-HPE@expeditors.com").
; Si vide, scanne la boite par defaut de l'utilisateur (repli pratique pour
; les tests, mais en usage reel la boite HPE doit etre precisee -- c'est
; elle qui recoit les mails du client, pas la boite personnelle).
Func _HPE_MailScanJSON($sMailbox = "")
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","message":"outlook_indisponible","alerts":[]}'

    Local $oInbox = 0
    If $sMailbox = "" Then
        $oInbox = $g_oNamespace.GetDefaultFolder(6) ; olFolderInbox
    Else
        Local $sWanted = StringUpper(StringStripWS($sMailbox, 3))
        Local $sWantedNoDomain = StringRegExpReplace($sWanted, "@.*$", "")
        Local $oStoreTarget = Null
        For $oStore In $g_oNamespace.Stores
            Local $sDisp = StringUpper($oStore.DisplayName)
            If $sDisp = $sWanted Or $sDisp = $sWantedNoDomain Or StringInStr($sDisp, $sWantedNoDomain) Then
                $oStoreTarget = $oStore
                ExitLoop
            EndIf
        Next
        If $oStoreTarget = Null Then Return '{"status":"error","message":"mailbox_not_found","alerts":[]}'
        $oInbox = $oStoreTarget.GetDefaultFolder(6)
    EndIf

    Local $oItems = $oInbox.Items
    $oItems.Sort("[ReceivedTime]", True) ; plus recent d'abord
    Local $sDateLimit = _FmtDate(_DateAdd('d', -14, _NowCalc()))

    Local $oThreadData = ObjCreate("Scripting.Dictionary")  ; reference -> champs (packe)
    Local $oThreadCount = ObjCreate("Scripting.Dictionary") ; reference -> nb de mails du fil
    Local $oNoMatch = ObjCreate("Scripting.Dictionary")     ; references confirmees sans dossier lie

    Local $iScanned = 0
    For $oItem In $oItems
        If $iScanned >= 500 Then ExitLoop
        If $oItem.Class <> 43 Then ContinueLoop ; olMail uniquement
        Local $itemDate = _FmtDate($oItem.ReceivedTime)
        If $itemDate < $sDateLimit Then ExitLoop
        $iScanned += 1

        Local $sSubj = $oItem.Subject
        Local $sRef = _HPE_ExtractRef($sSubj)
        If $sRef = "" Then ContinueLoop

        If $oThreadData.Exists($sRef) Then
            ; Mail plus ancien du meme fil (ex: le message initial sous la
            ; reponse la plus recente deja enregistree) : incremente juste
            ; le compteur, ne cree pas de 2e alerte.
            $oThreadCount.Item($sRef) = $oThreadCount.Item($sRef) + 1
            ContinueLoop
        EndIf
        If $oNoMatch.Exists($sRef) Then ContinueLoop

        Local $sDossier = IniRead($HPE_INI_PATH, _HPE_MapSec($sRef), "DOSSIER", "")
        If $sDossier = "" Then
            $oNoMatch.Add($sRef, True)
            ContinueLoop
        EndIf

        Local $sOwner = IniRead($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "OWNER", "")
        Local $sStatus = IniRead($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "STATUS", "")

        Local $sReply = "false"
        If StringRegExp($sSubj, "(?i)^(RE|TR|FW|FWD)\s*:") Then $sReply = "true"

        $oThreadData.Add($sRef, $oItem.EntryID & Chr(31) & $sDossier & Chr(31) & $sSubj & Chr(31) & $oItem.SenderName & Chr(31) & $itemDate & Chr(31) & $sOwner & Chr(31) & $sStatus & Chr(31) & $sReply)
        $oThreadCount.Add($sRef, 1)
    Next

    Local $sJson = "["
    For $sKey In $oThreadData.Keys
        Local $a = StringSplit($oThreadData.Item($sKey), Chr(31), 2)
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '{"reference":"' & _JsonEscape($sKey) & '"' & _
                ',"entryId":"' & _JsonEscape($a[0]) & '"' & _
                ',"dossier":"' & _JsonEscape($a[1]) & '"' & _
                ',"subject":"' & _JsonEscape($a[2]) & '"' & _
                ',"sender":"' & _JsonEscape($a[3]) & '"' & _
                ',"date":"' & _JsonEscape($a[4]) & '"' & _
                ',"owner":"' & _JsonEscape($a[5]) & '"' & _
                ',"status":"' & _JsonEscape($a[6]) & '"' & _
                ',"reply":' & $a[7] & _
                ',"threadCount":' & $oThreadCount.Item($sKey) & '}'
    Next
    $sJson &= "]"

    ; mappingCount/suiviCount pour l'apercu Dashboard.
    Local $iMapCount = 0, $iSuiviCount = 0
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 4) = "MAP:" Then $iMapCount += 1
            If StringLeft($sections[$i], 6) = "SUIVI:" Then $iSuiviCount += 1
        Next
    EndIf
    Return '{"status":"ok","alerts":' & $sJson & ',"mappingCount":' & $iMapCount & ',"suiviCount":' & $iSuiviCount & '}'
EndFunc

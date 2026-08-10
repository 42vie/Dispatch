; ============================================================================
; HPE_Tool.au3
; Outil HPE : lien entre les references client (INC/FXC/ARN) et les dossiers
; internes (J1A/H1A), + suivi (proprietaire/statut) et notes horodatees par
; dossier. Stockage dans un fichier ini (Config\HPE_Tool.ini), meme logique
; que les regles transporteur (CMR_SP.au3) : plus rapide et plus fiable que
; l'automation Excel utilisee dans une premiere version (pas de dependance a
; Excel installe, pas de risque de fichier verrouille ou de process fantome).
; ============================================================================

Global $HPE_INI_PATH = @ScriptDir & "\Config\HPE_Tool.ini"

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

; ============================================================================
; NOMS CONNUS -- liste parametrable (section SYSTEM, cle NAMES, separes par
; "|") utilisee pour deviner automatiquement le proprietaire d'un dossier en
; cherchant un de ces noms dans l'objet/expediteur/corps (signature incluse)
; du mail le plus recent du fil. Initialisee avec quelques noms d'equipe par
; defaut au premier appel si jamais configuree. Valeur par defaut gardee en
; Local (pas Global) : elle n'est utilisee que dans cette seule fonction, pas
; besoin de l'exposer script-wide -- voir la note en tete de Globals.au3 sur
; les Global declares hors de Globals.au3 qui ne s'executent jamais.
; ============================================================================
Func _HPE_GetNames()
    Local $sPacked = IniRead($HPE_INI_PATH, "SYSTEM", "NAMES", "")
    If $sPacked = "" Then
        Local $sDefaultNames = "Jason|Nabil|Chloé|Abderrahman|Charles"
        $sPacked = $sDefaultNames
        IniWrite($HPE_INI_PATH, "SYSTEM", "NAMES", $sPacked)
    EndIf
    Return $sPacked
EndFunc

Func _HPE_NamesListJSON()
    Local $aNames = StringSplit(_HPE_GetNames(), "|", 2)
    Local $sJson = "["
    For $i = 0 To UBound($aNames) - 1
        Local $sName = StringStripWS($aNames[$i], 3)
        If $sName = "" Then ContinueLoop
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '"' & _JsonEscape($sName) & '"'
    Next
    $sJson &= "]"
    Return '{"status":"ok","names":' & $sJson & '}'
EndFunc

Func _HPE_NamesSaveJSON($sBody)
    Local $sNames = _GetJsonValue($sBody, "names")
    IniWrite($HPE_INI_PATH, "SYSTEM", "NAMES", $sNames)
    Return '{"status":"ok"}'
EndFunc

; Cherche le 1er nom connu (liste $HPE_INI_PATH/SYSTEM/NAMES) present dans
; $sText (objet + expediteur + corps du mail, signature comprise), recherche
; insensible a la casse sur mot complet (pas de faux positif sur un nom
; contenu dans un autre mot). Renvoie "" si aucun nom ne matche. Bornes de
; mot definies explicitement (lettres ASCII + Latin-1 accentue) plutot que
; \b/\w, dont le comportement Unicode par defaut de PCRE est peu fiable sur
; les caracteres accentues (ex: "Chloé").
Func _HPE_DetectOwnerName($sText)
    Local $sPad = " " & StringRegExpReplace($sText, "\s+", " ") & " "
    Local $aNames = StringSplit(_HPE_GetNames(), "|", 2)
    For $i = 0 To UBound($aNames) - 1
        Local $sName = StringStripWS($aNames[$i], 3)
        If $sName = "" Then ContinueLoop
        If StringRegExp($sPad, "(?i)[^A-Za-zÀ-ÿ]" & $sName & "[^A-Za-zÀ-ÿ]") Then Return $sName
    Next
    Return ""
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
                        ',"tracking":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "TRACKING", "")) & '"' & _
                        ',"dateAjout":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "DATE", "")) & '"' & _
                        ',"operateur":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "OPERATOR", "")) & '"}'
            EndIf
        Next
    EndIf
    $sJson &= "]"
    Return '{"status":"ok","mapping":' & $sJson & '}'
EndFunc

; Ecrit/remplace un lien reference->dossier -- centralise la logique commune
; a l'ajout manuel et a l'import en masse (evite la duplication).
Func _HPE_MapWrite($sRef, $sDossier, $sTracking)
    Local $sType = ""
    Local $aType = StringRegExp($sRef, "(?i)^(INC|FXC|ARN)", 1)
    If Not @error Then $sType = StringUpper($aType[0])
    Local $sSec = _HPE_MapSec($sRef)
    IniWrite($HPE_INI_PATH, $sSec, "DOSSIER", $sDossier)
    IniWrite($HPE_INI_PATH, $sSec, "TYPE", $sType)
    IniWrite($HPE_INI_PATH, $sSec, "TRACKING", $sTracking)
    IniWrite($HPE_INI_PATH, $sSec, "DATE", _NowText())
    IniWrite($HPE_INI_PATH, $sSec, "OPERATOR", @UserName)
EndFunc

Func _HPE_MappingAddJSON($sBody)
    Local $sRef = StringUpper(StringStripWS(_GetJsonValue($sBody, "reference"), 3))
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    Local $sTracking = StringUpper(StringStripWS(_GetJsonValue($sBody, "tracking"), 3))
    If $sRef = "" Or $sDossier = "" Then Return '{"status":"error","message":"champs_vides"}'
    _HPE_MapWrite($sRef, $sDossier, $sTracking)
    Return '{"status":"ok"}'
EndFunc

Func _HPE_MappingDeleteJSON($sBody)
    Local $sRef = StringUpper(StringStripWS(_GetJsonValue($sBody, "reference"), 3))
    If $sRef = "" Then Return '{"status":"error","message":"reference_vide"}'
    IniDelete($HPE_INI_PATH, _HPE_MapSec($sRef))
    Return '{"status":"ok"}'
EndFunc

; Import en masse depuis un tableau colle/uploade (colonnes File / SHPMNT_REF
; / SP_TRACKING_NUM -- typiquement le rapport exporte par l'outil interne
; HPE, cf. commentaire dans le bloc HPE_UI_V1 cote HTML). $sBody attend
; "rows" = une ligne par dossier, colonnes separees par ";", lignes separees
; par "|" -- deja normalise cote JS (gere l'entete xlsx/csv/coller-Excel),
; donc pas besoin ici d'un parseur JSON de tableau (absent de ce projet).
Func _HPE_MappingImportJSON($sBody)
    Local $sRows = _GetJsonValue($sBody, "rows")
    If $sRows = "" Then Return '{"status":"error","message":"aucune_ligne"}'
    Local $aRows = StringSplit($sRows, "|", 2)
    Local $iAdded = 0, $iSkipped = 0
    Local $sSkippedJson = "["
    For $i = 0 To UBound($aRows) - 1
        If $aRows[$i] = "" Then ContinueLoop
        Local $aF = StringSplit($aRows[$i], ";", 2)
        Local $sFile = "", $sShp = "", $sTrk = ""
        If UBound($aF) >= 1 Then $sFile = StringStripWS($aF[0], 3)
        If UBound($aF) >= 2 Then $sShp = StringUpper(StringStripWS($aF[1], 3))
        If UBound($aF) >= 3 Then $sTrk = StringUpper(StringStripWS($aF[2], 3))

        Local $sRef = _HPE_ExtractRef($sFile)
        If $sRef = "" Then
            ; Repli : pas de INC/FXC/ARN reconnu dans la colonne File -- on
            ; garde la valeur brute (sans le prefixe "R#" de la recherche
            ; rapide E.TMS) comme reference plutot que de rejeter la ligne.
            $sRef = StringUpper(StringStripWS(StringRegExpReplace($sFile, "(?i)^R#", ""), 3))
        EndIf

        If $sRef = "" Or $sShp = "" Then
            $iSkipped += 1
            If $sSkippedJson <> "[" Then $sSkippedJson &= ","
            $sSkippedJson &= '"' & _JsonEscape($aRows[$i]) & '"'
            ContinueLoop
        EndIf

        _HPE_MapWrite($sRef, $sShp, $sTrk)
        $iAdded += 1
    Next
    $sSkippedJson &= "]"
    Return '{"status":"ok","added":' & $iAdded & ',"skipped":' & $iSkipped & ',"skippedLines":' & $sSkippedJson & '}'
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

    ; Repli : $sVal est peut-etre un numero de dossier ou un numero de tracking --
    ; on parcourt les references pour trouver celle qui correspond.
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 4) = "MAP:" Then
                Local $sDoss = StringUpper(IniRead($HPE_INI_PATH, $sections[$i], "DOSSIER", ""))
                Local $sTrack = StringUpper(IniRead($HPE_INI_PATH, $sections[$i], "TRACKING", ""))
                If $sDoss = $sVal Or ($sTrack <> "" And $sTrack = $sVal) Then
                    Return '{"status":"ok","found":true,"reference":"' & _JsonEscape(StringTrimLeft($sections[$i], 4)) & '","dossier":"' & _JsonEscape($sDoss) & '","tracking":"' & _JsonEscape($sTrack) & '"}'
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

; Liste "|" -> tableau JSON de chaines -- utilise pour CONTRIBUTORS (tous les
; operateurs ayant deja envoye un mail sur un dossier, pas seulement le
; proprietaire principal : un dossier peut avoir plusieurs intervenants).
Func _HPE_PipeListJSON($sPacked)
    If $sPacked = "" Then Return "[]"
    Local $a = StringSplit($sPacked, "|", 2)
    Local $s = "["
    For $i = 0 To UBound($a) - 1
        If $a[$i] = "" Then ContinueLoop
        If $s <> "[" Then $s &= ","
        $s &= '"' & _JsonEscape($a[$i]) & '"'
    Next
    $s &= "]"
    Return $s
EndFunc

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
                        ',"maj":"' & _JsonEscape(IniRead($HPE_INI_PATH, $sections[$i], "MAJ", "")) & '"' & _
                        ',"contributors":' & _HPE_PipeListJSON(IniRead($HPE_INI_PATH, $sections[$i], "CONTRIBUTORS", "")) & '}'
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
; mails recents (Reception ET Envoyes), les relie au dossier via le Mapping.
; Si l'objet ne contient pas de reference INC/FXC/ARN mais contient un numero
; de tracking deja enregistre dans le Mapping, il est rattache par ce biais
; (utile pour les mails envoyes/recus qui ne reprennent pas la reference
; client mais mentionnent le tracking transporteur). Les reponses ("RE:",
; "TR:"...) portent le meme objet/reference -- donc on regroupe par
; reference : une seule alerte par conversation, avec le nombre de mails du
; fil, tous dossiers (Reception + Envoyes) confondus. Reception et Envoyes
; sont d'abord collectes dans un seul tableau puis tries ensemble par date
; (voir _HPE_CollectFolder/_ArraySort ci-dessous) -- sans ce tri combine, le
; mail "le plus recent" retenu pour un fil serait celui du dossier scanne en
; premier, meme si l'autre dossier contient en realite un mail plus recent ;
; ca fausserait a la fois la date affichee et le calcul du statut auto.
; ============================================================================

; Collecte les mails eligibles d'un dossier Outlook (Reception ou Envoyes)
; dans le tableau partage $aRows (ByRef), une ligne par mail :
; [0]=date triable [1]=reference [2]=entryId [3]=objet [4]=expediteur/destinataire
; [5]=reponse(true/false) [6]=envoye(true/false)
; [7]=expediteur reel pour un mail envoye (utilise pour l'auto-attribution du
; proprietaire/contributeurs -- vide pour un mail recu, non pertinent)
Func _HPE_CollectFolder($oFolder, $bSent, $sDateLimit, $oTrackingMap, ByRef $aRows, ByRef $iRows)
    If Not IsObj($oFolder) Then Return
    Local $oItems = $oFolder.Items
    If $bSent Then
        $oItems.Sort("[SentOn]", True)
    Else
        $oItems.Sort("[ReceivedTime]", True)
    EndIf

    Local $iScanned = 0
    For $oItem In $oItems
        If $iScanned >= 500 Then ExitLoop
        If $iRows >= UBound($aRows) Then ExitLoop ; garde-fou, ne devrait pas arriver (tableau dimensionne large)
        If $oItem.Class <> 43 Then ContinueLoop ; olMail uniquement
        Local $itemDate = ""
        If $bSent Then
            $itemDate = _FmtDate($oItem.SentOn)
        Else
            $itemDate = _FmtDate($oItem.ReceivedTime)
        EndIf
        If $itemDate < $sDateLimit Then ExitLoop
        $iScanned += 1

        Local $sSubj = $oItem.Subject
        Local $sRef = _HPE_ExtractRef($sSubj)
        If $sRef = "" And $oTrackingMap.Count > 0 Then
            ; Repli : pas de INC/FXC/ARN dans l'objet -- peut-etre un numero de
            ; tracking deja connu (mapping cree manuellement avec un tracking).
            Local $sSubjUp = StringUpper($sSubj)
            For $sTrackKey In $oTrackingMap.Keys
                If StringInStr($sSubjUp, $sTrackKey) Then
                    $sRef = $oTrackingMap.Item($sTrackKey)
                    ExitLoop
                EndIf
            Next
        EndIf
        If $sRef = "" Then ContinueLoop

        Local $sReply = "false"
        If StringRegExp($sSubj, "(?i)^(RE|TR|FW|FWD)\s*:") Then $sReply = "true"

        ; Pour un mail envoye, "l'expediteur" affiche est en realite le/les
        ; destinataire(s) -- plus utile pour retrouver le fil que son propre nom.
        Local $sFrom = ""
        Local $sSenderId = ""
        If $bSent Then
            $sFrom = "→ " & $oItem.To
            ; Expediteur reel : utile quand $oFolder est celui d'une boite
            ; partagee (ex: CDG-HPE-Transcon@...) envoyee "de la part de" par
            ; un operateur -- SentOnBehalfOfName porte alors son nom.
            ; SenderName/Sender.Name reste le repli correct pour une boite
            ; personnelle (ou SentOnBehalfOfName est vide).
            $sSenderId = $oItem.SentOnBehalfOfName
            If @error Then $sSenderId = ""
            If $sSenderId = "" Then
                Local $oSenderObj = $oItem.Sender
                If IsObj($oSenderObj) Then
                    $sSenderId = $oSenderObj.Name
                    If @error Then $sSenderId = ""
                EndIf
            EndIf
            If $sSenderId = "" Then $sSenderId = @UserName
            $sSenderId = StringReplace($sSenderId, "|", " ")
        Else
            $sFrom = $oItem.SenderName
        EndIf
        Local $sSentFlag = "false"
        If $bSent Then $sSentFlag = "true"

        $aRows[$iRows][0] = $itemDate
        $aRows[$iRows][1] = $sRef
        $aRows[$iRows][2] = $oItem.EntryID
        $aRows[$iRows][3] = $sSubj
        $aRows[$iRows][4] = $sFrom
        $aRows[$iRows][5] = $sReply
        $aRows[$iRows][6] = $sSentFlag
        $aRows[$iRows][7] = $sSenderId
        $iRows += 1
    Next
EndFunc

; Resout un nom de boite partagee (ex: "CDG-HPE-Transcon@expeditors.com") en
; ses dossiers Reception/Envoyes -- utilise par le scan global ET la
; chronologie d'un dossier. $sMailbox vide = boite par defaut de l'utilisateur.
; Renvoie [$oInbox, $oSent], ou [0, 0] si le magasin demande est introuvable
; (appelant : @error non utilise ici, on teste juste IsObj sur le retour).
Func _HPE_ResolveMailboxFolders($sMailbox)
    Local $aOut[2] = [0, 0]
    If $sMailbox = "" Then
        $aOut[0] = $g_oNamespace.GetDefaultFolder(6) ; olFolderInbox
        $aOut[1] = $g_oNamespace.GetDefaultFolder(5) ; olFolderSentMail
        Return $aOut
    EndIf
    Local $sWanted = StringUpper(StringStripWS($sMailbox, 3))
    Local $sWantedNoDomain = StringRegExpReplace($sWanted, "@.*$", "")
    For $oStore In $g_oNamespace.Stores
        Local $sDisp = StringUpper($oStore.DisplayName)
        If $sDisp = $sWanted Or $sDisp = $sWantedNoDomain Or StringInStr($sDisp, $sWantedNoDomain) Then
            $aOut[0] = $oStore.GetDefaultFolder(6)
            $aOut[1] = $oStore.GetDefaultFolder(5)
            ExitLoop
        EndIf
    Next
    Return $aOut
EndFunc

; Numeros de tracking ET de dossier (SHPMNT_REF J1A/H1A) connus -> reference
; associee, pour retrouver un fil meme quand l'objet du mail ne contient pas
; de INC/FXC/ARN -- cas frequent : le mail ne reprend que "Livraison HPE -
; J1A0050493" (le numero de dossier importe depuis l'outil interne), sans la
; reference client FXC/ARN. $sOnlyDossier restreint la carte a un seul
; dossier (chronologie) ; vide = tous les dossiers (scan global).
Func _HPE_BuildTrackingMap($sOnlyDossier = "")
    Local $oMap = ObjCreate("Scripting.Dictionary")
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    If @error Then Return $oMap
    For $i = 1 To $sections[0]
        If StringLeft($sections[$i], 4) <> "MAP:" Then ContinueLoop
        Local $sRefKey = StringTrimLeft($sections[$i], 4)
        Local $sDossier = StringUpper(StringStripWS(IniRead($HPE_INI_PATH, $sections[$i], "DOSSIER", ""), 3))
        If $sOnlyDossier <> "" And $sDossier <> $sOnlyDossier Then ContinueLoop

        Local $sTrackKey = StringUpper(StringStripWS(IniRead($HPE_INI_PATH, $sections[$i], "TRACKING", ""), 3))
        If $sTrackKey <> "" And Not $oMap.Exists($sTrackKey) Then $oMap.Add($sTrackKey, $sRefKey)
        If $sDossier <> "" And Not $oMap.Exists($sDossier) Then $oMap.Add($sDossier, $sRefKey)
    Next
    Return $oMap
EndFunc

; $sMailbox : nom de la boite partagee a scanner (ex: "CDG-HPE-Transcon@expeditors.com").
; Si vide, scanne la boite par defaut de l'utilisateur (repli pratique pour
; les tests, mais en usage reel la boite HPE doit etre precisee -- c'est
; elle qui recoit les mails du client, pas la boite personnelle).
Func _HPE_MailScanJSON($sMailbox = "")
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","message":"outlook_indisponible","alerts":[]}'

    Local $aFolders = _HPE_ResolveMailboxFolders($sMailbox)
    Local $oInbox = $aFolders[0], $oSent = $aFolders[1]
    If $sMailbox <> "" And Not IsObj($oInbox) And Not IsObj($oSent) Then Return '{"status":"error","message":"mailbox_not_found","alerts":[]}'

    Local $sDateLimit = _FmtDate(_DateAdd('d', -14, _NowCalc()))

    Local $oTrackingMap = _HPE_BuildTrackingMap()

    ; Collecte Reception + Envoyes dans un seul tableau (max 500 par dossier),
    ; puis tri global par date decroissante -- voir note en tete de section.
    Local $aRows[1010][8]
    Local $iRows = 0
    _HPE_CollectFolder($oInbox, False, $sDateLimit, $oTrackingMap, $aRows, $iRows)
    _HPE_CollectFolder($oSent, True, $sDateLimit, $oTrackingMap, $aRows, $iRows)
    If $iRows > 1 Then _ArraySort($aRows, 1, 0, $iRows - 1, 0, 0)

    Local $oThreadData = ObjCreate("Scripting.Dictionary")  ; reference -> champs (packe), dossier connu
    Local $oThreadCount = ObjCreate("Scripting.Dictionary") ; reference -> nb de mails du fil
    Local $oNoMatchData = ObjCreate("Scripting.Dictionary")  ; reference -> champs (packe), pas de dossier lie
    Local $oNoMatchCount = ObjCreate("Scripting.Dictionary") ; reference -> nb de mails du fil (non lie)

    ; Statut auto : si le mail le plus recent d'un fil est un mail RECU (pas
    ; encore repondu) et date de $iAutoStaleDays jours ou plus, on bascule le
    ; statut du dossier sur $sAutoStatusLabel -- mais seulement s'il etait vide
    ; ou deja sur ce meme statut auto, jamais si l'utilisateur a saisi autre
    ; chose (Bloque, Termine...). On efface ce statut auto des que la
    ; condition n'est plus vraie (reponse envoyee, ou plus assez ancien).
    Local $iAutoStaleDays = 2
    Local $sAutoStatusLabel = "⏳ En attente de réponse"

    ; Auto-attribution proprietaire/contributeurs -- $oOwnerAssignedThisScan
    ; garde une seule tentative d'attribution du PROPRIETAIRE principal par
    ; dossier ce scan (le 1er mail envoye rencontre est le plus recent grace
    ; au tri global, donc "le proprietaire actuel" -- mais jamais d'ecrasement
    ; si un proprietaire est deja renseigne, manuellement ou automatiquement).
    ; $oContribAddedThisScan est independant : TOUS les operateurs ayant
    ; envoye un mail sur ce dossier sont ajoutes a CONTRIBUTORS, pas
    ; seulement le proprietaire -- si Nabil ET Jason ont chacun repondu sur
    ; le meme dossier, il doit apparaitre dans le suivi Dashboard des deux.
    Local $oOwnerAssignedThisScan = ObjCreate("Scripting.Dictionary")
    Local $oContribAddedThisScan = ObjCreate("Scripting.Dictionary")

    For $iR = 0 To $iRows - 1
        Local $sRef = $aRows[$iR][1]

        If $aRows[$iR][6] = "true" And $aRows[$iR][7] <> "" Then
            Local $sDossierForOwner = IniRead($HPE_INI_PATH, _HPE_MapSec($sRef), "DOSSIER", "")
            If $sDossierForOwner <> "" Then
                Local $sSenderName = $aRows[$iR][7]
                Local $sSuiviSec = _HPE_SuiviSec($sDossierForOwner)

                If Not $oOwnerAssignedThisScan.Exists($sDossierForOwner) Then
                    $oOwnerAssignedThisScan.Add($sDossierForOwner, True)
                    If IniRead($HPE_INI_PATH, $sSuiviSec, "OWNER", "") = "" Then
                        IniWrite($HPE_INI_PATH, $sSuiviSec, "OWNER", $sSenderName)
                        IniWrite($HPE_INI_PATH, $sSuiviSec, "MAJ", _NowText())
                    EndIf
                EndIf

                Local $sContribKey = $sDossierForOwner & "|" & StringUpper($sSenderName)
                If Not $oContribAddedThisScan.Exists($sContribKey) Then
                    $oContribAddedThisScan.Add($sContribKey, True)
                    Local $sContribList = IniRead($HPE_INI_PATH, $sSuiviSec, "CONTRIBUTORS", "")
                    Local $aContribs = StringSplit($sContribList, "|", 2)
                    Local $bAlready = False
                    For $c = 0 To UBound($aContribs) - 1
                        If StringUpper(StringStripWS($aContribs[$c], 3)) = StringUpper($sSenderName) Then
                            $bAlready = True
                            ExitLoop
                        EndIf
                    Next
                    If Not $bAlready Then
                        If $sContribList = "" Then
                            $sContribList = $sSenderName
                        Else
                            $sContribList &= "|" & $sSenderName
                        EndIf
                        IniWrite($HPE_INI_PATH, $sSuiviSec, "CONTRIBUTORS", $sContribList)
                    EndIf
                EndIf
            EndIf
        EndIf

        If $oThreadData.Exists($sRef) Then
            ; Mail plus ancien du meme fil : incremente juste le compteur, ne
            ; cree pas de 2e alerte (le 1er rencontre, donc le plus recent
            ; grace au tri global, reste la representation du fil).
            $oThreadCount.Item($sRef) = $oThreadCount.Item($sRef) + 1
            ContinueLoop
        EndIf
        If $oNoMatchData.Exists($sRef) Then
            $oNoMatchCount.Item($sRef) = $oNoMatchCount.Item($sRef) + 1
            ContinueLoop
        EndIf

        Local $sEntryId = $aRows[$iR][2], $sSubj = $aRows[$iR][3], $sFrom = $aRows[$iR][4]
        Local $itemDate = $aRows[$iR][0], $sReply = $aRows[$iR][5], $sSentFlag = $aRows[$iR][6]

        Local $sDossier = IniRead($HPE_INI_PATH, _HPE_MapSec($sRef), "DOSSIER", "")
        If $sDossier = "" Then
            ; Reference HPE detectee dans l'objet mais pas encore liee a un
            ; dossier : on la garde quand meme (au lieu de la jeter) pour que
            ; l'utilisateur puisse la relier depuis l'onglet HPE -- c'est la
            ; le vrai interet du scan pour les *nouveaux* dossiers.
            $oNoMatchData.Add($sRef, $sEntryId & Chr(31) & $sSubj & Chr(31) & $sFrom & Chr(31) & $itemDate & Chr(31) & $sReply & Chr(31) & $sSentFlag)
            $oNoMatchCount.Add($sRef, 1)
            ContinueLoop
        EndIf

        Local $sOwner = IniRead($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "OWNER", "")
        Local $sStatus = IniRead($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "STATUS", "")

        ; Detection auto du proprietaire (uniquement si pas deja renseigne
        ; manuellement) : cherche un nom connu dans objet/expediteur/corps du
        ; mail representatif du fil, corps inclus (donc aussi dans une
        ; signature). Le corps n'est recupere qu'ici, pour le seul mail
        ; representatif de chaque fil -- pas pour les 500 mails bruts scannes.
        If $sOwner = "" Then
            Local $oOwnerMail = 0
            If IsObj($g_oNamespace) Then $oOwnerMail = $g_oNamespace.GetItemFromID($sEntryId)
            If IsObj($oOwnerMail) Then
                Local $sBodyText = $oOwnerMail.Body
                If @error Then $sBodyText = ""
                Local $sDetected = _HPE_DetectOwnerName($sSubj & " " & $sFrom & " " & $sBodyText)
                If $sDetected <> "" Then
                    $sOwner = $sDetected
                    IniWrite($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "OWNER", $sOwner)
                EndIf
            EndIf
        EndIf

        Local $bStale = ($sSentFlag = "false") And (_DateDiff('d', $itemDate, _NowCalc()) >= $iAutoStaleDays)
        If $bStale And ($sStatus = "" Or $sStatus = $sAutoStatusLabel) Then
            If $sStatus <> $sAutoStatusLabel Then
                IniWrite($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "STATUS", $sAutoStatusLabel)
                IniWrite($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "MAJ", _NowText())
                $sStatus = $sAutoStatusLabel
            EndIf
        ElseIf Not $bStale And $sStatus = $sAutoStatusLabel Then
            ; La situation a change (reponse envoyee, ou redevenu recent) --
            ; on efface le statut auto pour laisser l'utilisateur requalifier.
            IniWrite($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "STATUS", "")
            $sStatus = ""
        EndIf

        $oThreadData.Add($sRef, $sEntryId & Chr(31) & $sDossier & Chr(31) & $sSubj & Chr(31) & $sFrom & Chr(31) & $itemDate & Chr(31) & $sOwner & Chr(31) & $sStatus & Chr(31) & $sReply & Chr(31) & $sSentFlag)
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
                ',"sent":' & $a[8] & _
                ',"threadCount":' & $oThreadCount.Item($sKey) & '}'
    Next
    $sJson &= "]"

    Local $sUnmappedJson = "["
    For $sKey In $oNoMatchData.Keys
        Local $u = StringSplit($oNoMatchData.Item($sKey), Chr(31), 2)
        If $sUnmappedJson <> "[" Then $sUnmappedJson &= ","
        $sUnmappedJson &= '{"reference":"' & _JsonEscape($sKey) & '"' & _
                ',"entryId":"' & _JsonEscape($u[0]) & '"' & _
                ',"subject":"' & _JsonEscape($u[1]) & '"' & _
                ',"sender":"' & _JsonEscape($u[2]) & '"' & _
                ',"date":"' & _JsonEscape($u[3]) & '"' & _
                ',"reply":' & $u[4] & _
                ',"sent":' & $u[5] & _
                ',"threadCount":' & $oNoMatchCount.Item($sKey) & '}'
    Next
    $sUnmappedJson &= "]"

    ; mappingCount/suiviCount pour l'apercu Dashboard.
    Local $iMapCount = 0, $iSuiviCount = 0
    Local $sectionsCount = IniReadSectionNames($HPE_INI_PATH)
    If Not @error Then
        For $i = 1 To $sectionsCount[0]
            If StringLeft($sectionsCount[$i], 4) = "MAP:" Then $iMapCount += 1
            If StringLeft($sectionsCount[$i], 6) = "SUIVI:" Then $iSuiviCount += 1
        Next
    EndIf
    Return '{"status":"ok","alerts":' & $sJson & ',"unmapped":' & $sUnmappedJson & ',"mappingCount":' & $iMapCount & ',"suiviCount":' & $iSuiviCount & '}'
EndFunc

; Ouvre directement un mail Outlook depuis son EntryID (renvoye par
; _HPE_MailScanJSON) -- meme logique que _OpenMail() dans Mail_Common.au3.
Func _HPE_OpenMailJSON($sBody)
    Local $sEntryId = _GetJsonValue($sBody, "entryId")
    If $sEntryId = "" Then Return '{"status":"error","message":"entryId_vide"}'
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","message":"outlook_indisponible"}'

    Local $oMail = 0
    If IsObj($g_oNamespace) Then $oMail = $g_oNamespace.GetItemFromID($sEntryId)
    If Not IsObj($oMail) Then Return '{"status":"error","message":"mail_introuvable"}'
    $oMail.Display()
    Return '{"status":"ok"}'
EndFunc

; ============================================================================
; CHRONOLOGIE D'UN DOSSIER -- contrairement au scan global (qui ne garde que
; le mail le plus recent par fil, pour l'affichage en alertes), ceci renvoie
; TOUS les mails lies aux references de ce dossier, fusionnes avec ses notes
; horodatees et tries du plus recent au plus ancien. Sert a repondre a "qui a
; envoye le dernier message sur ce dossier" et a afficher l'historique complet
; des interactions (mails + notes) avec l'interlocuteur/operateur HPE.
; ============================================================================

Func _HPE_DossierTimelineJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    Local $sMailbox = _GetJsonValue($sBody, "mailbox")
    If $sDossier = "" Then Return '{"status":"error","message":"dossier_vide"}'

    Local $sJson = "["
    Local $bFirst = True

    ; Notes horodatees -- toujours disponibles, meme sans Outlook.
    Local $sPacked = IniRead($HPE_INI_PATH, _HPE_SuiviSec($sDossier), "NOTES", "")
    If $sPacked <> "" Then
        Local $aRecords = StringSplit($sPacked, Chr(30), 2)
        For $i = 0 To UBound($aRecords) - 1
            If $aRecords[$i] = "" Then ContinueLoop
            Local $aF = StringSplit($aRecords[$i], Chr(31), 2)
            If UBound($aF) < 3 Then ContinueLoop
            If Not $bFirst Then $sJson &= ","
            $bFirst = False
            $sJson &= '{"type":"note","date":"' & _JsonEscape($aF[0]) & '","who":"' & _JsonEscape($aF[1]) & '","text":"' & _JsonEscape(_HPE_DecNote($aF[2])) & '"}'
        Next
    EndIf

    If Not _EDOC_EnsureOutlook() Then
        $sJson &= "]"
        Return '{"status":"ok","events":' & $sJson & ',"mailWarning":"outlook_indisponible"}'
    EndIf

    ; References (INC/FXC/ARN) qui appartiennent a CE dossier -- un dossier
    ; peut avoir plusieurs references (voir commentaire MAPPING en tete de
    ; fichier) ; sert a filtrer les mails trouves par extraction directe
    ; (_HPE_ExtractRef), qui elle n'est pas limitee a un dossier.
    Local $oOwnRefs = ObjCreate("Scripting.Dictionary")
    Local $sections = IniReadSectionNames($HPE_INI_PATH)
    If Not @error Then
        For $i = 1 To $sections[0]
            If StringLeft($sections[$i], 4) <> "MAP:" Then ContinueLoop
            Local $sDoss = StringUpper(StringStripWS(IniRead($HPE_INI_PATH, $sections[$i], "DOSSIER", ""), 3))
            If $sDoss <> $sDossier Then ContinueLoop
            Local $sRefKey = StringTrimLeft($sections[$i], 4)
            If Not $oOwnRefs.Exists($sRefKey) Then $oOwnRefs.Add($sRefKey, True)
        Next
    EndIf

    Local $oTrackingMap = _HPE_BuildTrackingMap($sDossier)
    If Not $oTrackingMap.Exists($sDossier) Then $oTrackingMap.Add($sDossier, $sDossier)

    Local $aFolders = _HPE_ResolveMailboxFolders($sMailbox)
    Local $oInbox = $aFolders[0], $oSent = $aFolders[1]

    If IsObj($oInbox) Or IsObj($oSent) Then
        Local $sDateLimit = _FmtDate(_DateAdd('d', -60, _NowCalc())) ; fenetre plus large que le scan global (14j) : c'est un historique
        Local $aRows[1010][8]
        Local $iRows = 0
        _HPE_CollectFolder($oInbox, False, $sDateLimit, $oTrackingMap, $aRows, $iRows)
        _HPE_CollectFolder($oSent, True, $sDateLimit, $oTrackingMap, $aRows, $iRows)
        For $iR = 0 To $iRows - 1
            Local $sRef = $aRows[$iR][1]
            ; L'extraction directe (INC/FXC/ARN dans l'objet) n'est pas limitee
            ; a ce dossier -- on ecarte donc tout mail dont la reference
            ; n'appartient pas explicitement a ce dossier.
            If $sRef <> $sDossier And Not $oOwnRefs.Exists($sRef) Then ContinueLoop
            Local $sDir = "reçu"
            If $aRows[$iR][6] = "true" Then $sDir = "envoyé"
            If Not $bFirst Then $sJson &= ","
            $bFirst = False
            $sJson &= '{"type":"mail","date":"' & _JsonEscape($aRows[$iR][0]) & '","who":"' & _JsonEscape($aRows[$iR][4]) & '","text":"' & _JsonEscape($aRows[$iR][3]) & '","direction":"' & $sDir & '","entryId":"' & _JsonEscape($aRows[$iR][2]) & '"}'
        Next
    EndIf

    $sJson &= "]"
    Return '{"status":"ok","events":' & $sJson & '}'
EndFunc

; Ouvre une nouvelle fenetre de composition Outlook pre-remplie pour un
; dossier HPE (le "Mails HPE" du widget dashboard par operateur) -- reste sur
; Display() (jamais Send()) : l'operateur relit et envoie lui-meme, comme
; partout ailleurs dans Dispatch ou un mail est prepare automatiquement.
Func _HPE_ComposeMailJSON($sBody)
    Local $sDossier = StringStripWS(_GetJsonValue($sBody, "dossier"), 3)
    Local $sTo = StringStripWS(_GetJsonValue($sBody, "to"), 3)
    If Not _EDOC_EnsureOutlook() Then Return '{"status":"error","message":"outlook_indisponible"}'
    Local $oOutlook = ObjCreate("Outlook.Application")
    If Not IsObj($oOutlook) Then Return '{"status":"error","message":"outlook_indisponible"}'
    Local $oMail = $oOutlook.CreateItem(0)
    If $sTo <> "" Then $oMail.To = $sTo
    If $sDossier <> "" Then $oMail.Subject = "Dossier " & $sDossier & " - "
    $oMail.Display
    Return '{"status":"ok"}'
EndFunc

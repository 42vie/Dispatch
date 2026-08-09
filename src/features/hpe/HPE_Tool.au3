; ============================================================================
; HPE_Tool.au3
; Outil HPE : lien entre les references client (INC/FXC/ARN) et les dossiers
; internes (J1A/H1A), + suivi (proprietaire/statut) et notes horodatees par
; dossier. Stockage dans un classeur Excel (Config\HPE_Suivi.xlsx, feuilles
; Mapping / Suivi / Notes) via automation COM en late-binding (ObjCreate),
; comme _CreateMailSP le fait deja avec Outlook.Application dans ce projet --
; necessite Microsoft Excel installe sur le poste.
; ============================================================================

Global Const $HPE_XLS_PATH = @ScriptDir & "\Config\HPE_Suivi.xlsx"

; Extrait la premiere reference INC/FXC/ARN + chiffres trouvee dans une
; chaine (ex: objet de mail), normalisee en majuscules sans espace/tiret.
; Renvoie "" si aucune reference trouvee.
Func _HPE_ExtractRef($s)
    Local $a = StringRegExp($s, "(?i)(INC|FXC|ARN)[- ]?(\d{4,})", 1)
    If @error Then Return ""
    Return StringUpper($a[0]) & $a[1]
EndFunc

; Ouvre (ou cree) le classeur HPE. Renvoie un Dictionary {app, book} ou
; @error=1 en cas d'echec (Excel absent/COM indisponible).
Func _HPE_Open()
    Local $oExcel = ObjCreate("Excel.Application")
    If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
    $oExcel.Visible = False
    $oExcel.DisplayAlerts = False

    Local $oBook = 0
    If FileExists($HPE_XLS_PATH) Then
        $oBook = $oExcel.Workbooks.Open($HPE_XLS_PATH)
    Else
        DirCreate(@ScriptDir & "\Config")
        $oBook = $oExcel.Workbooks.Add()
        While $oBook.Sheets.Count < 3
            $oBook.Sheets.Add(Default, $oBook.Sheets($oBook.Sheets.Count))
        WEnd
        While $oBook.Sheets.Count > 3
            $oBook.Sheets($oBook.Sheets.Count).Delete()
        WEnd
        $oBook.Sheets(1).Name = "Mapping"
        $oBook.Sheets(1).Range("A1").Value = "Reference"
        $oBook.Sheets(1).Range("B1").Value = "Type"
        $oBook.Sheets(1).Range("C1").Value = "Dossier"
        $oBook.Sheets(1).Range("D1").Value = "DateAjout"
        $oBook.Sheets(1).Range("E1").Value = "Operateur"
        $oBook.Sheets(2).Name = "Suivi"
        $oBook.Sheets(2).Range("A1").Value = "Dossier"
        $oBook.Sheets(2).Range("B1").Value = "Proprietaire"
        $oBook.Sheets(2).Range("C1").Value = "Statut"
        $oBook.Sheets(2).Range("D1").Value = "MAJ"
        $oBook.Sheets(3).Name = "Notes"
        $oBook.Sheets(3).Range("A1").Value = "Dossier"
        $oBook.Sheets(3).Range("B1").Value = "DateHeure"
        $oBook.Sheets(3).Range("C1").Value = "Auteur"
        $oBook.Sheets(3).Range("D1").Value = "Note"
        $oBook.SaveAs($HPE_XLS_PATH, 51) ; 51 = xlOpenXMLWorkbook (.xlsx)
    EndIf

    If Not IsObj($oBook) Then
        $oExcel.Quit()
        Return SetError(2, 0, 0)
    EndIf

    Local $h = ObjCreate("Scripting.Dictionary")
    $h.Add("app", $oExcel)
    $h.Add("book", $oBook)
    Return $h
EndFunc

Func _HPE_Close($h, $bSave = True)
    If Not IsObj($h) Then Return
    If $bSave Then $h.Item("book").Save()
    $h.Item("book").Close(False)
    $h.Item("app").Quit()
EndFunc

; Premiere ligne libre (colonne A) d'une feuille -- suppose un en-tete en
; ligne 1, donc renvoie au minimum 2.
Func _HPE_NextRow($oSheet)
    Return $oSheet.Cells($oSheet.Rows.Count, 1).End(-4162).Row + 1 ; -4162 = xlUp
EndFunc

; ============================================================================
; MAPPING (reference HPE <-> dossier)
; ============================================================================

Func _HPE_MappingListJSON()
    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Mapping")
    Local $iLast = _HPE_NextRow($oSheet) - 1
    Local $sJson = "["
    For $i = 2 To $iLast
        Local $sRef = $oSheet.Cells($i, 1).Value
        If $sRef = "" Then ContinueLoop
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '{"reference":"' & _JsonEscape($sRef) & '"' & _
                ',"type":"' & _JsonEscape($oSheet.Cells($i, 2).Value) & '"' & _
                ',"dossier":"' & _JsonEscape($oSheet.Cells($i, 3).Value) & '"' & _
                ',"dateAjout":"' & _JsonEscape($oSheet.Cells($i, 4).Value) & '"' & _
                ',"operateur":"' & _JsonEscape($oSheet.Cells($i, 5).Value) & '"}'
    Next
    $sJson &= "]"
    _HPE_Close($h, False)
    Return '{"status":"ok","mapping":' & $sJson & '}'
EndFunc

Func _HPE_MappingAddJSON($sBody)
    Local $sRef = StringUpper(StringStripWS(_GetJsonValue($sBody, "reference"), 3))
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    If $sRef = "" Or $sDossier = "" Then Return '{"status":"error","message":"champs_vides"}'
    Local $sType = ""
    Local $aType = StringRegExp($sRef, "(?i)^(INC|FXC|ARN)", 1)
    If Not @error Then $sType = StringUpper($aType[0])

    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Mapping")
    Local $iLast = _HPE_NextRow($oSheet) - 1

    ; Remplace si la reference existe deja (evite les doublons), sinon ajoute
    Local $iRow = 0
    For $i = 2 To $iLast
        If StringUpper($oSheet.Cells($i, 1).Value) = $sRef Then
            $iRow = $i
            ExitLoop
        EndIf
    Next
    If $iRow = 0 Then $iRow = $iLast + 1

    $oSheet.Cells($iRow, 1).Value = $sRef
    $oSheet.Cells($iRow, 2).Value = $sType
    $oSheet.Cells($iRow, 3).Value = $sDossier
    $oSheet.Cells($iRow, 4).Value = _NowText()
    $oSheet.Cells($iRow, 5).Value = @UserName

    _HPE_Close($h, True)
    Return '{"status":"ok"}'
EndFunc

Func _HPE_MappingDeleteJSON($sBody)
    Local $sRef = StringUpper(StringStripWS(_GetJsonValue($sBody, "reference"), 3))
    If $sRef = "" Then Return '{"status":"error","message":"reference_vide"}'
    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Mapping")
    Local $iLast = _HPE_NextRow($oSheet) - 1
    For $i = 2 To $iLast
        If StringUpper($oSheet.Cells($i, 1).Value) = $sRef Then
            $oSheet.Rows($i).Delete()
            ExitLoop
        EndIf
    Next
    _HPE_Close($h, True)
    Return '{"status":"ok"}'
EndFunc

; Cherche une valeur dans les deux colonnes (reference OU dossier) et
; renvoie la paire complete si trouvee.
Func _HPE_MappingLookupJSON($sBody)
    Local $sVal = StringUpper(StringStripWS(_GetJsonValue($sBody, "value"), 3))
    If $sVal = "" Then Return '{"status":"error","message":"valeur_vide"}'
    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Mapping")
    Local $iLast = _HPE_NextRow($oSheet) - 1
    Local $sRef = "", $sDossier = ""
    For $i = 2 To $iLast
        Local $r = StringUpper($oSheet.Cells($i, 1).Value)
        Local $d = StringUpper($oSheet.Cells($i, 3).Value)
        If $r = $sVal Or $d = $sVal Then
            $sRef = $r
            $sDossier = $d
            ExitLoop
        EndIf
    Next
    _HPE_Close($h, False)
    If $sRef = "" Then Return '{"status":"ok","found":false}'
    Return '{"status":"ok","found":true,"reference":"' & _JsonEscape($sRef) & '","dossier":"' & _JsonEscape($sDossier) & '"}'
EndFunc

; ============================================================================
; SUIVI (proprietaire / statut par dossier)
; ============================================================================

Func _HPE_SuiviListJSON()
    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Suivi")
    Local $iLast = _HPE_NextRow($oSheet) - 1
    Local $sJson = "["
    For $i = 2 To $iLast
        Local $sDoss = $oSheet.Cells($i, 1).Value
        If $sDoss = "" Then ContinueLoop
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '{"dossier":"' & _JsonEscape($sDoss) & '"' & _
                ',"owner":"' & _JsonEscape($oSheet.Cells($i, 2).Value) & '"' & _
                ',"status":"' & _JsonEscape($oSheet.Cells($i, 3).Value) & '"' & _
                ',"maj":"' & _JsonEscape($oSheet.Cells($i, 4).Value) & '"}'
    Next
    $sJson &= "]"
    _HPE_Close($h, False)
    Return '{"status":"ok","suivi":' & $sJson & '}'
EndFunc

Func _HPE_SuiviSaveJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    If $sDossier = "" Then Return '{"status":"error","message":"dossier_vide"}'
    Local $sOwner = _GetJsonValue($sBody, "owner")
    Local $sStatus = _GetJsonValue($sBody, "status")

    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Suivi")
    Local $iLast = _HPE_NextRow($oSheet) - 1
    Local $iRow = 0
    For $i = 2 To $iLast
        If StringUpper($oSheet.Cells($i, 1).Value) = $sDossier Then
            $iRow = $i
            ExitLoop
        EndIf
    Next
    If $iRow = 0 Then $iRow = $iLast + 1

    $oSheet.Cells($iRow, 1).Value = $sDossier
    $oSheet.Cells($iRow, 2).Value = $sOwner
    $oSheet.Cells($iRow, 3).Value = $sStatus
    $oSheet.Cells($iRow, 4).Value = _NowText()

    _HPE_Close($h, True)
    Return '{"status":"ok"}'
EndFunc

; ============================================================================
; NOTES (log horodate par dossier)
; ============================================================================

Func _HPE_NotesListJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    If $sDossier = "" Then Return '{"status":"error","message":"dossier_vide"}'
    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Notes")
    Local $iLast = _HPE_NextRow($oSheet) - 1
    Local $sJson = "["
    For $i = 2 To $iLast
        If StringUpper($oSheet.Cells($i, 1).Value) <> $sDossier Then ContinueLoop
        If $sJson <> "[" Then $sJson &= ","
        $sJson &= '{"date":"' & _JsonEscape($oSheet.Cells($i, 2).Value) & '"' & _
                ',"auteur":"' & _JsonEscape($oSheet.Cells($i, 3).Value) & '"' & _
                ',"note":"' & _JsonEscape($oSheet.Cells($i, 4).Value) & '"}'
    Next
    $sJson &= "]"
    _HPE_Close($h, False)
    Return '{"status":"ok","notes":' & $sJson & '}'
EndFunc

Func _HPE_NoteAddJSON($sBody)
    Local $sDossier = StringUpper(StringStripWS(_GetJsonValue($sBody, "dossier"), 3))
    Local $sNote = _GetJsonValue($sBody, "note")
    If $sDossier = "" Or $sNote = "" Then Return '{"status":"error","message":"champs_vides"}'
    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible"}'
    Local $oSheet = $h.Item("book").Sheets("Notes")
    Local $iRow = _HPE_NextRow($oSheet)
    $oSheet.Cells($iRow, 1).Value = $sDossier
    $oSheet.Cells($iRow, 2).Value = _NowText()
    $oSheet.Cells($iRow, 3).Value = @UserName
    $oSheet.Cells($iRow, 4).Value = $sNote
    _HPE_Close($h, True)
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

    Local $h = _HPE_Open()
    If @error Or Not IsObj($h) Then Return '{"status":"error","message":"excel_indisponible","alerts":[]}'
    Local $oMap = $h.Item("book").Sheets("Mapping")
    Local $iMapLast = _HPE_NextRow($oMap) - 1
    Local $oSuivi = $h.Item("book").Sheets("Suivi")
    Local $iSuiviLast = _HPE_NextRow($oSuivi) - 1

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

        Local $sDossier = ""
        For $i = 2 To $iMapLast
            If StringUpper($oMap.Cells($i, 1).Value) = $sRef Then
                $sDossier = StringUpper($oMap.Cells($i, 3).Value)
                ExitLoop
            EndIf
        Next
        If $sDossier = "" Then
            $oNoMatch.Add($sRef, True)
            ContinueLoop
        EndIf

        Local $sOwner = "", $sStatus = ""
        For $j = 2 To $iSuiviLast
            If StringUpper($oSuivi.Cells($j, 1).Value) = $sDossier Then
                $sOwner = $oSuivi.Cells($j, 2).Value
                $sStatus = $oSuivi.Cells($j, 3).Value
                ExitLoop
            EndIf
        Next

        Local $sReply = "false"
        If StringRegExp($sSubj, "(?i)^(RE|TR|FW|FWD)\s*:") Then $sReply = "true"

        $oThreadData.Add($sRef, $oItem.EntryID & Chr(31) & $sDossier & Chr(31) & $sSubj & Chr(31) & $oItem.SenderName & Chr(31) & $itemDate & Chr(31) & $sOwner & Chr(31) & $sStatus & Chr(31) & $sReply)
        $oThreadCount.Add($sRef, 1)
    Next

    _HPE_Close($h, False)

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
    ; mappingCount/suiviCount : reutilise les comptes deja lus dans cette
    ; meme session Excel pour alimenter l'apercu du Dashboard sans avoir a
    ; rouvrir Excel une 2e/3e fois juste pour compter des lignes.
    Local $iMapCount = $iMapLast - 1
    If $iMapCount < 0 Then $iMapCount = 0
    Local $iSuiviCount = $iSuiviLast - 1
    If $iSuiviCount < 0 Then $iSuiviCount = 0
    Return '{"status":"ok","alerts":' & $sJson & ',"mappingCount":' & $iMapCount & ',"suiviCount":' & $iSuiviCount & '}'
EndFunc

; ============================================================================
; ContactsService.au3
; Service contacts.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _CleanContactsFiles()
    Local $aFiles[4] = [$g_sContactsFile, "", "", ""]
    ; Chercher aussi les chunks
    For $i = 0 To 9
        Local $sChkF = @ScriptDir & "\dispatch_contacts_" & $i & ".json"
        If FileExists($sChkF) And $i < 3 Then $aFiles[$i + 1] = $sChkF
    Next

    Local $iCleaned = 0
    For $f = 0 To 3
        If $aFiles[$f] = "" Then ContinueLoop
        If Not FileExists($aFiles[$f]) Then ContinueLoop

        Local $hRead = FileOpen($aFiles[$f], 256)
        If $hRead = -1 Then ContinueLoop
        Local $sContent = FileRead($hRead)
        FileClose($hRead)

        Local $sBefore = $sContent

        ; Corriger les apostrophes mal encodées
        $sContent = StringReplace($sContent, "â€™", "'")       ; UTF-8 mojibake
        $sContent = StringReplace($sContent, "&#039;", "'")     ; HTML entity
        $sContent = StringReplace($sContent, "&apos;", "'")     ; XML entity
        $sContent = StringReplace($sContent, "&#x27;", "'")     ; Hex HTML entity
        $sContent = StringReplace($sContent, "'", "'")          ; Smart quote left
        $sContent = StringReplace($sContent, "'", "'")          ; Smart quote right
        $sContent = StringReplace($sContent, "Ã©", "é")         ; UTF-8 mojibake é
        $sContent = StringReplace($sContent, "Ã¨", "è")         ; UTF-8 mojibake è
        $sContent = StringReplace($sContent, "Ãª", "ê")         ; UTF-8 mojibake ê
        $sContent = StringReplace($sContent, "Ã ", "à")         ; UTF-8 mojibake à
        $sContent = StringReplace($sContent, "Ã§", "ç")         ; UTF-8 mojibake ç
        $sContent = StringReplace($sContent, "Ã¢", "â")         ; UTF-8 mojibake â
        $sContent = StringReplace($sContent, "Ã®", "î")         ; UTF-8 mojibake î
        $sContent = StringReplace($sContent, "Ã¼", "ü")         ; UTF-8 mojibake ü
        $sContent = StringReplace($sContent, "Ã¶", "ö")         ; UTF-8 mojibake ö
        $sContent = StringReplace($sContent, "&amp;", "&")       ; Double-encoded &
        $sContent = StringReplace($sContent, "&lt;", "<")        ; HTML <
        $sContent = StringReplace($sContent, "&gt;", ">")        ; HTML >
        $sContent = StringReplace($sContent, "&quot;", '"')      ; HTML "
        ; Supprimer les espaces multiples consécutifs
        While StringInStr($sContent, "  ")
            $sContent = StringReplace($sContent, "  ", " ")
        WEnd

        If $sContent <> $sBefore Then
            Local $hWrite = FileOpen($aFiles[$f], 2 + 256)
            FileWrite($hWrite, $sContent)
            FileClose($hWrite)
            $iCleaned += 1
        EndIf
    Next

    If $iCleaned > 0 Then
        TrayTip("Dispatch — Contacts", $iCleaned & " fichier(s) contact nettoyé(s).", 5, 1)
    Else
        TrayTip("Dispatch — Contacts", "Aucun nettoyage nécessaire.", 3, 1)
    EndIf
EndFunc


; ==============================================================================
; CONTACTS TSV - Archivage des anciens JSON/chunks contacts
; ==============================================================================

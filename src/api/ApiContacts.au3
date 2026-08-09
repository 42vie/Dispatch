; ============================================================================
; ApiContacts.au3
; Routes contacts.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================
; Deplace depuis HttpRouter.au3 (blocs "/api/save-contacts",
; "/api/load-contacts") -- logique inchangee, juste regroupee ici comme
; prevu par l'architecture cible du projet. Ne pas confondre avec
; ContactsService.au3 (nettoyage/maintenance des contacts) ou
; ApiNetwork.au3 (etat reseau).

Func _ApiContacts_Save($iSocket, $sBody)
    _BackupRotate($g_sContactsFile, 10)
    Local $hFileC = FileOpen($g_sContactsFile, 2 + 256)
    If $hFileC = -1 Then
        _SendHttpResponse($iSocket, 500, "application/json", '{"status":"error","reason":"cannot_write_contacts_tsv"}')
    Else
        FileWrite($hFileC, $sBody)
        FileClose($hFileC)
        _SendHttpResponse($iSocket, 200, "application/json", '{"status":"ok","format":"tsv"}')
    EndIf
EndFunc

Func _ApiContacts_Load($iSocket)
    Local $sContacts = ""
    If FileExists($g_sContactsFile) Then
        Local $hReadC = FileOpen($g_sContactsFile, 256)
        If $hReadC <> -1 Then
            $sContacts = FileRead($hReadC)
            FileClose($hReadC)
        EndIf
    ElseIf FileExists($g_sContactsLegacyFile) Then
        Local $hOldC = FileOpen($g_sContactsLegacyFile, 256)
        If $hOldC <> -1 Then
            $sContacts = FileRead($hOldC)
            FileClose($hOldC)
        EndIf
    EndIf
    If $sContacts = "" Then $sContacts = "#DISPATCH_CONTACTS_TSV" & @LF
    _SendHttpResponse($iSocket, 200, "text/plain", $sContacts)
EndFunc

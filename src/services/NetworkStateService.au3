; ============================================================================
; NetworkStateService.au3
; Service Outlook/reseau.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _GetStoresList($oNamespace)
    Local $s = ""
    For $oStore In $oNamespace.Stores
        $s &= $oStore.DisplayName & "|"
    Next
    If StringRight($s, 1) = "|" Then $s = StringTrimRight($s, 1)
    Return $s
EndFunc

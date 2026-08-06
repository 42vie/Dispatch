; ============================================================================
; CMR_Mail.au3
; Mail CMR.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _CMR_MailByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aBL) Then Return False
    _CreateMailForResult($idx)
    Return True
EndFunc

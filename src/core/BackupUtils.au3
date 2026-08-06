; ============================================================================
; BackupUtils.au3
; Backups et rotation.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _BackupRotate($sFile, $iMax)
    If Not FileExists($sFile) Then Return
    Local $sBackDir = @ScriptDir & "\backups"
    Local $sBase = StringRegExpReplace($sFile, ".*\\", "")
    ; Rotation : supprimer le plus ancien, décaler les autres
    Local $sOldest = $sBackDir & "\" & $sBase & "." & $iMax & ".bak"
    If FileExists($sOldest) Then FileDelete($sOldest)
    For $b = $iMax - 1 To 1 Step -1
        Local $sSrc = $sBackDir & "\" & $sBase & "." & $b & ".bak"
        Local $sDst = $sBackDir & "\" & $sBase & "." & ($b + 1) & ".bak"
        If FileExists($sSrc) Then FileMove($sSrc, $sDst, 1)
    Next
    ; Copier le fichier actuel en .1.bak
    FileCopy($sFile, $sBackDir & "\" & $sBase & ".1.bak", 1)
EndFunc

; ==============================================================================
; AUDIT FC — Diagnostic détaillé pour FileClosing
; ==============================================================================

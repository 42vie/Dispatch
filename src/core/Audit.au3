; ============================================================================
; Audit.au3
; Logs et audit.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _AuditLog($sLevel, $sMsg)
    Local $sTs = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
    Local $sLine = "[" & $sTs & "] [" & $sLevel & "] " & $sMsg & @CRLF
    ConsoleWrite($sLine)
    ; Écrire dans le fichier audit
    Local $hAudit = FileOpen($g_sAuditLog, 1 + 256) ; 1=append, 256=UTF-8
    If $hAudit <> -1 Then
        FileWrite($hAudit, $sLine)
        FileClose($hAudit)
    EndIf
EndFunc

; ==============================================================================
; HEALTH CHECK SILENCIEUX — Tourne en arrière-plan toutes les minutes
; Détecte les erreurs silencieuses et les log
; ==============================================================================

; ============================================================================
; HttpResponse.au3
; Generation des reponses HTTP.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _SendHttpResponse($iSocket, $iCode, $sContentType, $sData)
    Local $sStatus = "200 OK"
    If $iCode = 400 Then $sStatus = "400 Bad Request"
    If $iCode = 404 Then $sStatus = "404 Not Found"
    Local $bData    = StringToBinary($sData, 4)
    Local $iLen     = BinaryLen($bData)
    Local $sHeaders = "HTTP/1.1 " & $sStatus & @CRLF & _
                      "Content-Type: " & $sContentType & "; charset=UTF-8" & @CRLF & _
                      "Content-Length: " & $iLen & @CRLF & _
                      "Access-Control-Allow-Origin: *" & @CRLF & _
                      "Access-Control-Allow-Headers: Content-Type" & @CRLF & _
                      "Cache-Control: no-cache" & @CRLF & _
                      "Connection: close" & @CRLF & @CRLF
    ; Envoyer par blocs pour éviter les envois partiels sur gros payloads
    Local $bAll = StringToBinary($sHeaders, 4) & $bData
    Local $iTotal = BinaryLen($bAll)
    Local $iSent = 0
    While $iSent < $iTotal
        Local $bChunk = BinaryMid($bAll, $iSent + 1, 8192)
        Local $iRes = TCPSend($iSocket, $bChunk)
        If @error Then ExitLoop
        If $iRes > 0 Then
            $iSent += $iRes
        Else
            Sleep(5)
        EndIf
    WEnd
EndFunc

; ==============================================================================
; UTILITAIRES
; ==============================================================================

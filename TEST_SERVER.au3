#NoTrayIcon
Opt("MustDeclareVars", 0)

#include <File.au3>
#include <String.au3>
#include <MsgBoxConstants.au3>

Local $sURL = "http://localhost:9500/api/ping"
Local $hOpen = InetOpen($sURL)
Local $sRead = InetRead($sURL, 1)

If @error Or StringLen($sRead) = 0 Then
    MsgBox(16, "ERREUR", "Le serveur Dispatch ne répond pas sur " & $sURL & @CRLF & @CRLF & "Veuillez d'abord lancer MainDispatch.au3 (F5 dans SciTE).")
    Exit
EndIf

Local $sResponse = BinaryToString($sRead)

If StringInStr($sResponse, "ok") Or StringInStr($sResponse, "pong") Then
    MsgBox(64, "SUCCÈ§S", "Le serveur Dispatch répond correctement !" & @CRLF & @CRLF & "Ré§§ponse: " & $sResponse & @CRLF & @CRLF & "L'interface est accessible sur http://localhost:9500")
Else
    MsgBox(48, "ATTENTION", "Le serveur répond mais la réponse semble inhabituelle :" & @CRLF & @CRLF & $sResponse)
EndIf

InetClose($hOpen)

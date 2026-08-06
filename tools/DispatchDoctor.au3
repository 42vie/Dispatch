#include-once

Opt("MustDeclareVars", 1)

Global Const $STDOUT_CHILD = 2
Global Const $STDERR_CHILD = 4
Global Const $g_sToolName = "DispatchDoctor"

_Main()

Func _Main()
    If $CmdLine[0] < 1 Then
        ConsoleWrite("Usage: DispatchDoctor.au3 <MainDispatch.au3> [--domain CMR|COMAT|ETMS]" & @CRLF)
        Exit 2
    EndIf

    Local $sScript = $CmdLine[1]
    If Not FileExists($sScript) Then
        ConsoleWrite("[ERROR] Script introuvable: " & $sScript & @CRLF)
        Exit 2
    EndIf

    Local $sDomain = "COMMON"
    If $CmdLine[0] >= 3 And $CmdLine[2] = "--domain" Then $sDomain = StringUpper($CmdLine[3])

    Local $sAutoIt = @ProgramFilesDir & "\\AutoIt3\\AutoIt3.exe"
    If Not FileExists($sAutoIt) Then $sAutoIt = @ProgramFilesDir & " (x86)\\AutoIt3\\AutoIt3.exe"
    If Not FileExists($sAutoIt) Then
        ConsoleWrite("[ERROR] AutoIt3.exe introuvable." & @CRLF)
        Exit 2
    EndIf

    Local $sCommand = '"' & $sAutoIt & '" /ErrorStdOut "' & $sScript & '"'
    Local $iPid = Run($sCommand, @WorkingDir, @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
    If $iPid = 0 Then
        ConsoleWrite("[ERROR] Impossible de lancer le compilateur AutoIt." & @CRLF)
        Exit 2
    EndIf

    Local $sOutput = ""
    While ProcessExists($iPid)
        $sOutput &= StdoutRead($iPid)
        Sleep(25)
    WEnd
    $sOutput &= StdoutRead($iPid)

    Local $sReport = _BuildReport($sScript, $sDomain, $sOutput)
    Local $sReportFile = $sScript & ".doctor.log"
    Local $hReport = FileOpen($sReportFile, 2 + 256)
    If $hReport <> -1 Then
        FileWrite($hReport, $sReport)
        FileClose($hReport)
    EndIf

    ConsoleWrite($sReport)
    If StringInStr($sOutput, "==>") > 0 Then Exit 1
    Exit 0
EndFunc

Func _BuildReport($sScript, $sDomain, $sOutput)
    Local $sReport = "DispatchDoctor - domaine: " & $sDomain & @CRLF
    $sReport &= "Script: " & $sScript & @CRLF & @CRLF

    If StringStripWS($sOutput, 3) = "" Then
        Return $sReport & "[OK] Aucune erreur retournee par AutoIt." & @CRLF
    EndIf

    Local $aErrors = StringRegExp($sOutput, '(?m)^"([^"]+)" \((\d+)\) : ==> (.+)$', 3)
    If Not IsArray($aErrors) Then
        Return $sReport & "[COMPILER OUTPUT]" & @CRLF & $sOutput & @CRLF
    EndIf

    Local $iErrorCount = UBound($aErrors) / 3
    For $i = 0 To $iErrorCount - 1
        Local $sFile = $aErrors[$i * 3]
        Local $iLine = Number($aErrors[$i * 3 + 1])
        Local $sMessage = $aErrors[$i * 3 + 2]
        $sReport &= "[ERROR] " & $sMessage & @CRLF
        $sReport &= "  Fichier : " & $sFile & @CRLF
        $sReport &= "  Ligne   : " & $iLine & @CRLF
        $sReport &= "  Domaine : " & $sDomain & @CRLF
        $sReport &= "  Code    : " & _ReadLine($sFile, $iLine) & @CRLF
        $sReport &= "  Fonction: " & _FindFunction($sFile, $iLine) & @CRLF
        $sReport &= "  Variables candidates: " & _FindVariables(_ReadLine($sFile, $iLine)) & @CRLF & @CRLF
    Next
    Return $sReport
EndFunc

Func _ReadLine($sFile, $iTarget)
    Local $hFile = FileOpen($sFile, 0)
    If $hFile = -1 Then Return "<fichier inaccessible>"
    Local $iLine = 0, $sLine = ""
    While 1
        $sLine = FileReadLine($hFile)
        If @error Then ExitLoop
        $iLine += 1
        If $iLine = $iTarget Then
            FileClose($hFile)
            Return StringStripWS($sLine, 3)
        EndIf
    WEnd
    FileClose($hFile)
    Return "<ligne introuvable>"
EndFunc

Func _FindFunction($sFile, $iTarget)
    Local $hFile = FileOpen($sFile, 0)
    If $hFile = -1 Then Return "<inconnue>"
    Local $iLine = 0, $sLine = "", $sFunction = "<globale>"
    While 1
        $sLine = FileReadLine($hFile)
        If @error Then ExitLoop
        $iLine += 1
        If StringRegExp($sLine, '(?i)^\s*Func\s+') Then $sFunction = StringRegExpReplace(StringStripWS($sLine, 3), '(?i)^Func\s+([^\(]+).*$', '$1')
        If $iLine >= $iTarget Then ExitLoop
    WEnd
    FileClose($hFile)
    Return $sFunction
EndFunc

Func _FindVariables($sLine)
    Local $aVars = StringRegExp($sLine, '\$[A-Za-z_][A-Za-z0-9_]*', 3)
    If Not IsArray($aVars) Then Return "aucune"
    Local $sResult = ""
    For $i = 0 To UBound($aVars) - 1
        If $i > 0 Then $sResult &= ", "
        $sResult &= $aVars[$i]
    Next
    Return $sResult
EndFunc
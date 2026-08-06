#include-once
Opt("MustDeclareVars", 1)

Global Const $g_sRoot = @ScriptDir
Global $g_sMain = ""
Global $g_sDomain = "COMMON"
Global $g_sMode = "full"
Global $g_sAutoIt = ""
Global $g_aFiles[0]
Global $g_aFunctions[0]
Global $g_aDiagnostics[0]
Global $g_aVisited[0]

_Main()

Func _Main()
    _ReadArguments()
    $g_sMain = _FindMainScript()
    If $g_sMain = "" Then _Fatal("MainDispatch.au3 introuvable dans " & $g_sRoot)
    DirCreate($g_sRoot & "\\reports")
    _AddFile($g_sMain)
    _ScanIncludes($g_sMain)
    _AnalyzeProject()
    If $g_sMode = "full" Or $g_sMode = "compile" Then _RunCompiler()
    _WriteReports()
    If _CountDiagnostics("ERROR") > 0 Then Exit 1
    Exit 0
EndFunc

Func _ReadArguments()
    For $i = 1 To $CmdLine[0]
        Switch $CmdLine[$i]
            Case "--domain"
                If $i < $CmdLine[0] Then $i += 1
                $g_sDomain = StringUpper($CmdLine[$i])
            Case "--mode"
                If $i < $CmdLine[0] Then $i += 1
                $g_sMode = StringLower($CmdLine[$i])
            Case "--help", "-h"
                ConsoleWrite("Usage: DispatchDoctor.au3 [MainDispatch.au3] [--domain CMR] [--mode full|compile|analyze|list]" & @CRLF)
                Exit 0
            Case Else
                If StringRight($CmdLine[$i], 4) = ".au3" Then $g_sMain = _AbsolutePath($CmdLine[$i])
        EndSwitch
    Next
    If $g_sMode <> "full" And $g_sMode <> "compile" And $g_sMode <> "analyze" And $g_sMode <> "list" Then $g_sMode = "full"
EndFunc

Func _FindMainScript()
    If $g_sMain <> "" And FileExists($g_sMain) Then Return $g_sMain
    Local $sCandidate = $g_sRoot & "\\MainDispatch.au3"
    If FileExists($sCandidate) Then Return $sCandidate
    Return ""
EndFunc

Func _AbsolutePath($sPath)
    If StringRegExp($sPath, "(?i)^[A-Z]:\\\\") Or StringLeft($sPath, 2) = "\\\\" Then Return $sPath
    Return $g_sRoot & "\\" & $sPath
EndFunc

Func _AddFile($sFile)
    If Not FileExists($sFile) Then Return False
    $sFile = _Normalize($sFile)
    For $i = 0 To UBound($g_aFiles) - 1
        If $g_aFiles[$i] = $sFile Then Return True
    Next
    ReDim $g_aFiles[UBound($g_aFiles) + 1]
    $g_aFiles[UBound($g_aFiles) - 1] = $sFile
    Return True
EndFunc

Func _ScanIncludes($sFile)
    Local $h = FileOpen($sFile, 0)
    If $h = -1 Then Return
    Local $sLine = "", $sDir = StringRegExpReplace($sFile, "\\[^\\]+$", "")
    While 1
        $sLine = FileReadLine($h)
        If @error Then ExitLoop
        Local $a = StringRegExp($sLine, '(?i)#include\s+["<]([^">]+)[">]', 1)
        If IsArray($a) Then
            Local $sInc = _Normalize($sDir & "\\" & $a[0])
            If FileExists($sInc) Then
                If _AddFile($sInc) Then _ScanIncludes($sInc)
            Else
                _AddDiagnostic("ERROR", "MISSING_INCLUDE", $sFile, 0, "Fichier inclus introuvable: " & $a[0], "Verifier le chemin relatif")
            EndIf
        EndIf
    WEnd
    FileClose($h)
EndFunc

Func _AnalyzeProject()
    If $g_sMode = "compile" Then Return
    For $i = 0 To UBound($g_aFiles) - 1
        _AnalyzeFile($g_aFiles[$i])
    Next
EndFunc

Func _AnalyzeFile($sFile)
    Local $h = FileOpen($sFile, 0)
    If $h = -1 Then Return
    Local $sLine = "", $iLine = 0, $sFunction = "<global>", $sDeclared = ""
    While 1
        $sLine = FileReadLine($h)
        If @error Then ExitLoop
        $iLine += 1
        If StringRegExp($sLine, '(?i)^\s*Func\s+') Then
            $sFunction = StringRegExpReplace(StringStripWS($sLine, 3), '(?i)^Func\s+([^\(]+).*$', '$1')
            _AddFunction($sFunction, $sFile, $iLine)
            $sDeclared &= " " & $sLine
        ElseIf StringRegExp($sLine, '(?i)\b(Local|Global|Dim|Const)\b') Then
            $sDeclared &= " " & $sLine
        EndIf
        If StringRegExp($sLine, '(?i)\bFunc\s+.*\)') Then
            $sDeclared &= " " & $sLine
        EndIf
    WEnd
    FileClose($h)
    _FindDuplicateFunctions($sFile)
EndFunc

Func _AddFunction($sName, $sFile, $iLine)
    ReDim $g_aFunctions[UBound($g_aFunctions) + 1]
    $g_aFunctions[UBound($g_aFunctions) - 1] = $sName & "|" & $sFile & "|" & $iLine
EndFunc

Func _FindDuplicateFunctions($sFile)
    Local $h = FileOpen($sFile, 0)
    If $h = -1 Then Return
    Local $sLine = "", $iLine = 0
    While 1
        $sLine = FileReadLine($h)
        If @error Then ExitLoop
        $iLine += 1
        Local $a = StringRegExp($sLine, '(?i)^\s*Func\s+([^\(]+)', 1)
        If IsArray($a) Then
            Local $sName = StringStripWS($a[0], 3), $iCount = 0, $sWhere = ""
            For $j = 0 To UBound($g_aFunctions) - 1
                Local $p = StringSplit($g_aFunctions[$j], "|", 2)
                If StringLower($p[0]) = StringLower($sName) Then
                    $iCount += 1
                    $sWhere &= $p[1] & ":" & $p[2] & " "
                EndIf
            Next
            If $iCount > 1 Then _AddDiagnostic("ERROR", "DUPLICATE_FUNCTION", $sFile, $iLine, "Fonction dupliquee: " & $sName, $sWhere)
        EndIf
    WEnd
    FileClose($h)
EndFunc

Func _RunCompiler()
    $g_sAutoIt = _FindAutoIt()
    If $g_sAutoIt = "" Then
        _AddDiagnostic("ERROR", "AUTOIT_NOT_FOUND", $g_sMain, 0, "AutoIt3.exe introuvable", "Configurer AutoIt ou l'ajouter au PATH")
        Return
    EndIf
    Local $iPid = Run('"' & $g_sAutoIt & '" /ErrorStdOut "' & $g_sMain & '"', $g_sRoot, @SW_HIDE, 2 + 4)
    If $iPid = 0 Then
        _AddDiagnostic("ERROR", "COMPILER_START", $g_sMain, 0, "Impossible de lancer AutoIt3.exe", "Verifier l'installation AutoIt")
        Return
    EndIf
    Local $sOut = ""
    While ProcessExists($iPid)
        $sOut &= StdoutRead($iPid)
        Sleep(25)
    WEnd
    $sOut &= StdoutRead($iPid)
    Local $a = StringRegExp($sOut, '(?m)^"([^"]+)" \((\d+)\) : ==> (.+)$', 3)
    If IsArray($a) Then
        For $i = 0 To UBound($a) - 1 Step 3
            Local $sMsg = $a[$i + 2], $sCode = "COMPILER_ERROR"
            If StringInStr($sMsg, "Variable used") Then $sCode = "UNDECLARED_VARIABLE"
            If StringInStr($sMsg, "Unknown function") Then $sCode = "UNKNOWN_FUNCTION"
            If StringInStr($sMsg, "Duplicate function") Then $sCode = "DUPLICATE_FUNCTION"
            _AddDiagnostic("ERROR", $sCode, $a[$i], Number($a[$i + 1]), $sMsg, _ReadLine($a[$i], Number($a[$i + 1])))
        Next
    ElseIf StringStripWS($sOut, 3) <> "" Then
        _AddDiagnostic("ERROR", "COMPILER_OUTPUT", $g_sMain, 0, StringStripWS($sOut, 3), "Consulter la sortie complete du compilateur")
    EndIf
EndFunc

Func _FindAutoIt()
    Local $a[4] = [$g_sRoot & "\\AutoIt3.exe", @ProgramFilesDir & "\\AutoIt3\\AutoIt3.exe", @ProgramFilesDir & " (x86)\\AutoIt3\\AutoIt3.exe", "AutoIt3.exe"]
    For $i = 0 To UBound($a) - 1
        If FileExists($a[$i]) Then Return $a[$i]
    Next
    Return ""
EndFunc

Func _AddDiagnostic($sSeverity, $sCode, $sFile, $iLine, $sMessage, $sSuggestion)
    ReDim $g_aDiagnostics[UBound($g_aDiagnostics) + 1]
    $g_aDiagnostics[UBound($g_aDiagnostics) - 1] = $sSeverity & "|" & $sCode & "|" & $sFile & "|" & $iLine & "|" & StringReplace($sMessage, "|", "/") & "|" & StringReplace($sSuggestion, "|", "/")
EndFunc

Func _WriteReports()
    Local $sBase = $g_sRoot & "\\reports\\DispatchDoctor-report"
    Local $sTxt = "DispatchDoctor" & @CRLF & "Projet: " & $g_sMain & @CRLF & "Domaine: " & $g_sDomain & @CRLF & "Mode: " & $g_sMode & @CRLF & @CRLF
    $sTxt &= "Fichiers analyses: " & UBound($g_aFiles) & @CRLF & "Fonctions indexees: " & UBound($g_aFunctions) & @CRLF & "Diagnostics: " & UBound($g_aDiagnostics) & @CRLF & @CRLF
    For $i = 0 To UBound($g_aDiagnostics) - 1
        $sTxt &= StringReplace($g_aDiagnostics[$i], "|", " | ") & @CRLF
    Next
    _WriteUtf8($sBase & ".txt", $sTxt)
    _WriteUtf8($sBase & ".json", _ToJson())
    _WriteUtf8($sBase & ".html", "<html><head><meta charset='utf-8'><title>DispatchDoctor</title></head><body><pre>" & _HtmlEscape($sTxt) & "</pre></body></html>")
    ConsoleWrite($sTxt)
EndFunc

Func _ToJson()
    Local $s = '{"project":"' & _JsonEscape($g_sMain) & '","domain":"' & $g_sDomain & '","files":' & UBound($g_aFiles) & ',"functions":' & UBound($g_aFunctions) & ',"diagnostics":['
    For $i = 0 To UBound($g_aDiagnostics) - 1
        Local $p = StringSplit($g_aDiagnostics[$i], "|", 2)
        If $i > 0 Then $s &= ","
        $s &= '{"severity":"' & $p[0] & '","code":"' & $p[1] & '","file":"' & _JsonEscape($p[2]) & '","line":' & $p[3] & ',"message":"' & _JsonEscape($p[4]) & '","suggestion":"' & _JsonEscape($p[5]) & '"}'
    Next
    Return $s & ']}'
EndFunc

Func _ReadLine($sFile, $iTarget)
    Local $h = FileOpen($sFile, 0)
    If $h = -1 Then Return ""
    Local $i = 0, $s = ""
    While 1
        $s = FileReadLine($h)
        If @error Then ExitLoop
        $i += 1
        If $i = $iTarget Then
            FileClose($h)
            Return StringStripWS($s, 3)
        EndIf
    WEnd
    FileClose($h)
    Return ""
EndFunc

Func _WriteUtf8($sFile, $sText)
    Local $h = FileOpen($sFile, 2 + 256)
    If $h <> -1 Then FileWrite($h, $sText) : FileClose($h)
EndFunc

Func _Normalize($sPath)
    Return StringReplace(StringRegExpReplace($sPath, "\\\\+", "\\"), "\\.\\", "\\")
EndFunc

Func _JsonEscape($s)
    $s = StringReplace($s, "\\", "\\\\")
    $s = StringReplace($s, '"', '\\"')
    $s = StringReplace($s, @CRLF, "\\n")
    Return $s
EndFunc

Func _HtmlEscape($s)
    $s = StringReplace($s, "&", "&amp;")
    $s = StringReplace($s, "<", "&lt;")
    Return StringReplace($s, ">", "&gt;")
EndFunc

Func _CountDiagnostics($sSeverity)
    Local $n = 0
    For $i = 0 To UBound($g_aDiagnostics) - 1
        If StringLeft($g_aDiagnostics[$i], StringLen($sSeverity)) = $sSeverity Then $n += 1
    Next
    Return $n
EndFunc

Func _Fatal($sMessage)
    ConsoleWrite("[FATAL] " & $sMessage & @CRLF)
    Exit 2
EndFunc
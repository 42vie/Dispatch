; ============================================================================
; FileUtils.au3
; Fonctions fichiers, CSV, encodage et nettoyage.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _CsvCacheClear($sPrefix = "")
    If Not IsObj($g_oCsvCache) Then Return False
    If $sPrefix = "" Then
        $g_oCsvCache.RemoveAll()
        Return True
    EndIf
    Local $aKeys = $g_oCsvCache.Keys
    For $i = 0 To UBound($aKeys) - 1
        If StringInStr($aKeys[$i], $sPrefix) = 1 Then $g_oCsvCache.Remove($aKeys[$i])
    Next
    Return True
EndFunc

Func _CsvLoadCached($file)
    If Not FileExists($file) Then Return SetError(1, 0, 0)
    Local $stamp = FileGetSize($file) & ":" & FileGetTime($file, 0, 1)
    Local $key = $file & "|" & $stamp
    If IsObj($g_oCsvCache) And $g_oCsvCache.Exists($key) Then Return $g_oCsvCache.Item($key)

    ; purge anciennes entrées du même fichier si le fichier a changé
    If IsObj($g_oCsvCache) Then
        Local $aKeys = $g_oCsvCache.Keys
        For $i = 0 To UBound($aKeys) - 1
            If StringInStr($aKeys[$i], $file & "|") = 1 Then $g_oCsvCache.Remove($aKeys[$i])
        Next
    EndIf

    Local $txt = FileRead($file)
    If @error Then Return SetError(2, 0, 0)
    $txt = StringReplace($txt, @CRLF, @LF)
    $txt = StringReplace($txt, @CR, @LF)
    Local $raw = StringSplit($txt, @LF, 2)
    Local $lines[0]
    For $i = 0 To UBound($raw) - 1
        If StringStripWS($raw[$i], 3) = "" Then ContinueLoop
        ReDim $lines[UBound($lines) + 1]
        $lines[UBound($lines) - 1] = $raw[$i]
    Next
    If IsObj($g_oCsvCache) Then $g_oCsvCache.Add($key, $lines)
    Return $lines
EndFunc

Func _StringToArrayBL($s)
    Local $parts = StringSplit(StringReplace($s, " / ", @LF), @LF, 2)
    Local $a[0]
    For $i = 0 To UBound($parts) - 1
        Local $j = _CleanJ($parts[$i])
        If $j = "" Then ContinueLoop
        ReDim $a[UBound($a) + 1]
        $a[UBound($a) - 1] = $j
    Next
    Return $a
EndFunc

Func _QueueGroupsFromText($txt)
    $txt = StringReplace($txt, @CRLF, @LF)
    $txt = StringReplace($txt, @CR, @LF)
    Local $raw = StringSplit($txt, @LF, 2)
    Local $groups[0]
    Local $cur = ""
    For $i = 0 To UBound($raw) - 1
        Local $line = StringStripWS($raw[$i], 3)
        If $line = "" Or $line = "---" Then
            If StringStripWS($cur, 3) <> "" Then
                ReDim $groups[UBound($groups) + 1]
                $groups[UBound($groups) - 1] = $cur
                $cur = ""
            EndIf
        Else
            If $cur <> "" Then $cur &= @LF
            $cur &= $line
        EndIf
    Next
    If StringStripWS($cur, 3) <> "" Then
        ReDim $groups[UBound($groups) + 1]
        $groups[UBound($groups) - 1] = $cur
    EndIf
    Return $groups
EndFunc

Func _QueueGroupToArray($s)
    Local $norm = StringReplace(StringReplace($s, " / ", @LF), ";", @LF)
    Local $raw = StringSplit($norm, @LF, 2)
    Local $a[0]
    For $i = 0 To UBound($raw) - 1
        Local $j = _CleanJ($raw[$i])
        If $j = "" Then ContinueLoop
        ReDim $a[UBound($a) + 1]
        $a[UBound($a) - 1] = $j
    Next
    Return $a
EndFunc

Func _QueuePreviewText()
    Local $g = _QueueGroups()
    Local $s = ""
    If UBound($g) = 0 Then Return "Queue vide. Colle des J puis clique Ajouter BL ou Import auto."
    For $i = 0 To UBound($g) - 1
        Local $a = _QueueGroupToArray($g[$i])
        $s &= "BL " & ($i + 1) & " : " & _JoinArray($a, " / ") & @CRLF
    Next
    Return $s
EndFunc

; ============================================================
; SP
; ============================================================

Func _SP_WriteRule($carrier, $to, $cc, $bcc, $subject, $pdf, $body, $signature)
    Local $sec = _SP_Sec($carrier)
    IniWrite($INI_PATH, $sec, "CARRIER", $carrier)
    IniWrite($INI_PATH, $sec, "TO", $to)
    IniWrite($INI_PATH, $sec, "CC", $cc)
    IniWrite($INI_PATH, $sec, "BCC", $bcc)
    IniWrite($INI_PATH, $sec, "SUBJECT", $subject)
    IniWrite($INI_PATH, $sec, "PDF", $pdf)
    IniWrite($INI_PATH, $sec, "BODY", _Enc($body))
    IniWrite($INI_PATH, $sec, "SIGNATURE", _Enc($signature))
EndFunc

Func _NowText()
    Return @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc

Func _LastBL_Write(ByRef $aNums)
    If Not FileExists($INPUT_PATH) Then DirCreate($INPUT_PATH)
    IniWrite($LASTBL_MARKER, "LAST", "Date", _NowText())
    IniWrite($LASTBL_MARKER, "LAST", "Operator", $g_sOperator)
    IniWrite($LASTBL_MARKER, "LAST", "WindowsUser", @UserName)
    IniWrite($LASTBL_MARKER, "LAST", "Machine", @ComputerName)
    IniWrite($LASTBL_MARKER, "LAST", "BL", _JoinArray($aNums, " / "))
    IniWrite($LASTBL_MARKER, "LAST", "InputCSV", $INPUT_CSV)
    IniWrite($LASTBL_MARKER, "LAST", "InputGEN", $INPUT_GEN)
    _LastBL_Refresh()
    Return True
EndFunc

Func _LastBL_Text()
    If Not FileExists($LASTBL_MARKER) Then Return "Dernier BL : aucune info"
    Local $d = IniRead($LASTBL_MARKER, "LAST", "Date", "")
    Local $op = IniRead($LASTBL_MARKER, "LAST", "Operator", "")
    Local $wu = IniRead($LASTBL_MARKER, "LAST", "WindowsUser", "")
    Local $pc = IniRead($LASTBL_MARKER, "LAST", "Machine", "")
    Local $bl = IniRead($LASTBL_MARKER, "LAST", "BL", "")
    If $op = "" Then $op = $wu
    If $pc <> "" Then $pc = " / " & $pc
    If $bl = "" Then Return "Dernier BL : " & $op & $pc & " - " & $d
    Return "Dernier BL : " & $bl & @CRLF & "Par " & $op & $pc & " le " & $d
EndFunc

Func _ResetFiles()
    _CsvCacheClear() ; PRO CLEAN : évite toute donnée CSV ancienne en mémoire
    If FileExists($INPUT_CSV) Then FileDelete($INPUT_CSV)
    If FileExists($INPUT_GEN) Then FileDelete($INPUT_GEN)
    Local $a[6] = ["outputDIMS.csv", "outputTDIMS.csv", "outputPACKID.csv", "outputREFS.csv", "outputEDICEC.csv", "outputGEN.csv"]
    For $i = 0 To 5
        If FileExists($OUTPUT_PATH & $a[$i]) Then FileDelete($OUTPUT_PATH & $a[$i])
    Next
EndFunc

Func _CreateInputFiles(ByRef $aNums)
    Local $first = _CleanJ($aNums[0])
    If $first = "" Then
        _Status("INPUT KO : premier J vide.")
        Return False
    EndIf

    Local $tracking = $first & $g_sOperator
    Local $csv = ""
    For $i = 0 To UBound($aNums) - 1
        Local $sJ = _CleanJ($aNums[$i])
        If $sJ <> "" Then $csv &= $sJ & "," & $tracking & @CRLF
    Next

    If StringStripWS($csv, 3) = "" Then
        _Status("INPUT KO : aucun J valide à écrire.")
        Return False
    EndIf

    If Not _WriteAnsi($INPUT_CSV, $csv) Then
        _Status("INPUT KO : impossible de créer" & @CRLF & $INPUT_CSV)
        Return False
    EndIf

    If Not _WriteAnsi($INPUT_GEN, $first & "," & $tracking & @CRLF) Then
        _Status("INPUT KO : impossible de créer" & @CRLF & $INPUT_GEN)
        Return False
    EndIf

    _LastBL_Write($aNums)
    _Status("INPUT OK :" & @CRLF & $INPUT_CSV & @CRLF & $INPUT_GEN & @CRLF & _LastBL_Text())
    Return True
EndFunc

Func _SnapCsvGet($snap, $fileName, $rowIndex, $colIndex)
    Return _CsvGet($snap & $fileName, $rowIndex, $colIndex)
EndFunc

Func _CsvGet($file, $rowIndex, $colIndex)
    Local $lines = _CsvLoadCached($file)
    If Not IsArray($lines) Then Return ""
    If $rowIndex < 0 Or $rowIndex >= UBound($lines) Then Return ""
    Local $a = _ParseCsvLine($lines[$rowIndex])
    If $colIndex < 0 Or $colIndex >= UBound($a) Then Return ""
    Return _CleanJ($a[$colIndex])
EndFunc

Func _GetLine($file, $rowIndex)
    Local $lines = _CsvLoadCached($file)
    If Not IsArray($lines) Then Return ""
    If $rowIndex < 0 Or $rowIndex >= UBound($lines) Then Return ""
    Return $lines[$rowIndex]
EndFunc

Func _ParseCsvLine($line)
    Local $delim = ","
    If StringInStr($line, ";") And Not StringInStr($line, ",") Then $delim = ";"
    If StringInStr($line, @TAB) Then $delim = @TAB
    Local $a[0], $field = "", $inQ = False
    For $i = 1 To StringLen($line)
        Local $ch = StringMid($line, $i, 1)
        If $ch = '"' Then
            If $inQ And $i < StringLen($line) And StringMid($line, $i + 1, 1) = '"' Then
                $field &= '"'
                $i += 1
            Else
                $inQ = Not $inQ
            EndIf
        ElseIf $ch = $delim And Not $inQ Then
            ReDim $a[UBound($a) + 1]
            $a[UBound($a) - 1] = $field
            $field = ""
        Else
            $field &= $ch
        EndIf
    Next
    ReDim $a[UBound($a) + 1]
    $a[UBound($a) - 1] = $field
    Return $a
EndFunc

Func _JoinUniqueColsSnap($snap, $fileName, $row, $c1, $c2, $sep)
    Local $d = ObjCreate("Scripting.Dictionary")
    Local $res = ""
    For $c = $c1 To $c2
        Local $v = _SnapCsvGet($snap, $fileName, $row, $c)
        If $v = "" Then ContinueLoop
        If Not $d.Exists($v) Then
            $d.Add($v, 1)
            If $res <> "" Then $res &= $sep
            $res &= $v
        EndIf
    Next
    Return $res
EndFunc

Func _TextToHtml($s)
    $s = _H($s)
    $s = StringReplace($s, @CRLF, "<br>")
    $s = StringReplace($s, @CR, "<br>")
    $s = StringReplace($s, @LF, "<br>")
    Return "<div style='font-family:Calibri,Arial;font-size:11pt'>" & $s & "</div>"
EndFunc

Func _ArrayNumsClipboard(ByRef $aNums)
    Local $s = ""
    For $i = 0 To UBound($aNums) - 1
        Local $j = _CleanJ($aNums[$i])
        If $j <> "" Then $s &= $j & @CRLF
    Next
    Return $s
EndFunc

Func _CleanJ($s)
    $s = String($s)
    $s = StringReplace($s, Chr(239) & Chr(187) & Chr(191), "")
    $s = StringReplace($s, "ï»¿", "")
    $s = StringReplace($s, "Ï»¿", "")
    $s = StringReplace($s, ChrW(65279), "")
    Return StringStripWS($s, 3)
EndFunc

Func _CleanNum($s)
    Return StringReplace($s, ",", ".")
EndFunc

Func _JoinArray(ByRef $a, $sep)
    Local $s = ""
    For $i = 0 To UBound($a) - 1
        If $s <> "" Then $s &= $sep
        $s &= _CleanJ($a[$i])
    Next
    Return $s
EndFunc

Func _WriteAnsi($file, $txt)
    Local $sDir = StringLeft($file, StringInStr($file, "\", 0, -1))
    If $sDir <> "" And Not FileExists($sDir) Then DirCreate($sDir)
    Local $h = FileOpen($file, 2)
    If $h = -1 Then Return False
    FileWrite($h, $txt)
    FileClose($h)
    Return FileExists($file)
EndFunc

Func _WriteUtf8($file, $txt)
    Local $sDir = StringLeft($file, StringInStr($file, "\", 0, -1))
    If $sDir <> "" And Not FileExists($sDir) Then DirCreate($sDir)
    Local $o = ObjCreate("ADODB.Stream")
    If Not IsObj($o) Then Return False
    $o.Type = 2
    $o.Charset = "utf-8"
    $o.Open
    $o.WriteText($txt)
    $o.SaveToFile($file, 2)
    $o.Close
    If Not FileExists($file) Then Return False
    If FileGetSize($file) <= 0 Then Return False
    Return True
EndFunc

Func _Safe($s)
    Local $a[12] = ["\", "/", ":", "*", "?", '"', "<", ">", "|", @CR, @LF, @TAB]
    For $i = 0 To UBound($a) - 1
        $s = StringReplace($s, $a[$i], "_")
    Next
    Return StringStripWS($s, 3)
EndFunc

Func _LoadProfiles($idCombo, $sMail)
    Local $sList = IniRead($CFG_FILE, "System", "ProfilesList", "")
    If $sList = "" Then Return
    GUICtrlSetData($idCombo, "|" & $sList, StringSplit($sList, "|")[1])
    _LoadActions(GUICtrlRead($idCombo), $sMail)
EndFunc

Func _Ext($file)
    Local $p = StringInStr($file, ".", 0, -1)
    If $p = 0 Then Return ""
    Return StringLower(StringTrimLeft($file, $p))
EndFunc

Func _SafeName($s)
    Local $bad = '<>:"/\|?*'
    For $i = 1 To StringLen($bad)
        $s = StringReplace($s, StringMid($bad, $i, 1), "_")
    Next
    Return $s
EndFunc

Func _PipeToComma($s)
    Return StringReplace($s, "|", ", ")
EndFunc

Func _PipeContains($list, $item)
    Return StringInStr("|" & $list & "|", "|" & $item & "|") > 0
EndFunc

Func _RemoveFromPipeList($list, $item)
    Local $s = "|" & $list & "|"
    $s = StringReplace($s, "|" & $item & "|", "|")
    $s = StringReplace($s, "||", "|")
    If StringLeft($s, 1) = "|" Then $s = StringTrimLeft($s, 1)
    If StringRight($s, 1) = "|" Then $s = StringTrimRight($s, 1)
    Return $s
EndFunc

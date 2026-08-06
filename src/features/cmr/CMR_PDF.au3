; ============================================================================
; CMR_PDF.au3
; Generation HTML/PDF des CMR a partir des snapshots, edition de l'apercu,
; petits helpers de mise en forme (labels, encadres) utilises par ces ecrans.
; ----------------------------------------------------------------------------
; Deplace depuis StateService.au3 (bloc CMR_FULL_FUSION_MODULE_APPENDED) pour
; que le code CMR vive dans src/features/cmr/, comme le reste du projet.
; ============================================================================


; ============================================================
; HTML / PDF FROM SNAPSHOT
; ============================================================

Func _OpenPreviewEditor($htmlFile, ByRef $aNums, $idx)
    If Not FileExists($htmlFile) Then Return False
    Local $hPrev = GUICreate("Modifier BL - " & _JoinArray($aNums, " / "), 1120, 820, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_SIZEBOX))
    GUICtrlCreateLabel("Modifie le BL puis clique Enregistrer. Enregistrer sauvegarde le HTML ET régénère le PDF utilisé pour le mail.", 10, 8, 970, 22)
    Local $bSave = GUICtrlCreateButton("Enregistrer", 10, 35, 150, 32)
    Local $bPdfOpen = GUICtrlCreateButton("Ouvrir PDF", 170, 35, 140, 32)
    Local $bHtml = GUICtrlCreateButton("Ouvrir HTML", 320, 35, 140, 32)
    Local $bClose = GUICtrlCreateButton("Fermer", 970, 35, 130, 32)
    Local $oIE = ObjCreate("Shell.Explorer.2")
    If Not IsObj($oIE) Then
        MsgBox(16+262144, $APP_TITLE, "Contrôle aperçu HTML indisponible.")
        GUIDelete($hPrev)
        Return False
    EndIf
    GUICtrlCreateObj($oIE, 10, 75, 1090, 730)
    $oIE.Silent = True
    $oIE.Navigate("file:///" & StringReplace($htmlFile, "\", "/"))
    While $oIE.Busy Or $oIE.ReadyState <> 4
        Sleep(100)
    WEnd
    GUISetState(@SW_SHOW, $hPrev)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $bClose
                ExitLoop
            Case $bSave
                Local $pdfSaved = _SavePreviewAndRebuildPdf($oIE, $aNums, $idx, True)
                If FileExists($pdfSaved) Then
                    MsgBox(64, $APP_TITLE, "Modifications enregistrées et PDF régénéré." & @CRLF & @CRLF & $pdfSaved)
                Else
                    MsgBox(16, $APP_TITLE, "Impossible d’enregistrer ou de régénérer le PDF.")
                EndIf
            Case $bHtml
                ShellExecute($htmlFile)
            Case $bPdfOpen
                If $idx >= 0 Then
                    _EnsureLatestPdfByIndex($idx)
                    If $idx < UBound($g_aPDF) And FileExists($g_aPDF[$idx]) Then ShellExecute($g_aPDF[$idx])
                EndIf
        EndSwitch
    WEnd
    GUIDelete($hPrev)
    Return True
EndFunc


Func _SaveEditedPreview(ByRef $oIE, ByRef $aNums)
    Local $html = ""
    If IsObj($oIE) And IsObj($oIE.Document) Then $html = "<!doctype html>" & $oIE.Document.documentElement.outerHTML
    If $html = "" Then Return ""
    Local $name = "Printed_Delivery_Order_" & _Safe(_JoinArray($aNums, "_")) & "_MODIFIE_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".html"
    Local $file = $HTML_PATH & $name
    If Not _WriteUtf8($file, $html) Then
        Local $fallback = @ScriptDir & "\HTML\"
        If Not FileExists($fallback) Then DirCreate($fallback)
        $file = $fallback & $name
        If Not _WriteUtf8($file, $html) Then Return ""
    EndIf
    Return $file
EndFunc


Func _SavePreviewAndRebuildPdf(ByRef $oIE, ByRef $aNums, $idx, $bOpenPdf = False)
    Local $edited = _SaveEditedPreview($oIE, $aNums)
    If Not FileExists($edited) Then Return ""
    Local $carrier = ""
    Local $delivery = ""
    If $idx >= 0 And $idx < UBound($g_aCarrier) Then
        $carrier = $g_aCarrier[$idx]
        $delivery = $g_aDelivery[$idx]
    EndIf
    Local $pdf = _HtmlToPdf($edited, $aNums, $carrier, $delivery)
    If FileExists($pdf) Then
        If $idx >= 0 And $idx < UBound($g_aPDF) Then
            $g_aHTML[$idx] = $edited
            $g_aPDF[$idx] = $pdf
        EndIf
        _Status("Modifications enregistrées + PDF régénéré :" & @CRLF & $pdf)
        If $bOpenPdf Then ShellExecute($pdf)
        Return $pdf
    EndIf
    Return ""
EndFunc


Func _EnsureLatestPdfByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return ""
    If Not FileExists($g_aHTML[$idx]) Then Return ""
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _HtmlToPdf($g_aHTML[$idx], $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
    If FileExists($pdf) Then
        $g_aPDF[$idx] = $pdf
        Return $pdf
    EndIf
    Return $g_aPDF[$idx]
EndFunc


Func _GenerateHtmlFromSnapshot(ByRef $aNums, $snap)
    Local $rows = "", $totalPcs = 0, $totalWeight = 0, $totalVol = 0
    For $r = 0 To 39
        Local $j = _SnapCsvGet($snap, "outputDIMS.csv", $r, 0)
        If $j = "" Then ContinueLoop
        Local $dim = _SnapCsvGet($snap, "outputDIMS.csv", $r, 1)
        Local $pcs = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 1)
        Local $wgt = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 2)
        Local $vol = _SnapCsvGet($snap, "outputTDIMS.csv", $r, 3)
        Local $packid = _JoinUniqueColsSnap($snap, "outputPACKID.csv", $r, 2, 12, " - ")
        Local $ref = _JoinUniqueColsSnap($snap, "outputREFS.csv", $r, 2, 12, " - ")
        $rows &= "<tr><td contenteditable='true'>" & _H($j) & "</td><td contenteditable='true'>" & _H($ref) & "</td><td contenteditable='true'>" & _H($packid) & "</td><td contenteditable='true'>" & _H($pcs) & "</td><td contenteditable='true'>" & _H($wgt) & "</td><td contenteditable='true'>" & _H($dim) & "</td></tr>"
        $totalPcs += Number(_CleanNum($pcs))
        $totalWeight += Number(_CleanNum($wgt))
        $totalVol += Number(_CleanNum($vol))
    Next
    Local $deliver = _FormatDeliverTo(_DefaultDeliverToSnap($snap))
    Local $ready = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 1)
    Local $delivDate = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 5)
    Local $delivTime = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 6)
    Local $carrier = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 7)
    Local $instr = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 8)
    If _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9) <> "" Then $instr &= " " & _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9)
    $instr = _FormatDeliveryInstructions($instr)
    Local $first = _CleanJ($aNums[0])
    Local $tracking = $first & $g_sOperator
    Local $html = _HtmlTemplate($tracking, $g_sOperator, $deliver, $ready, $delivDate, $delivTime, $carrier, $instr, $rows, _TrimNum($totalPcs), _TrimNum($totalWeight), _TrimNum($totalVol))
    Local $name = "Printed_Delivery_Order_" & _Safe(_JoinArray($aNums, "_")) & "_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".html"
    Local $file = $HTML_PATH & $name
    If Not _WriteUtf8($file, $html) Then
        Local $fallback = @ScriptDir & "\HTML\"
        If Not FileExists($fallback) Then DirCreate($fallback)
        $file = $fallback & $name
        If Not _WriteUtf8($file, $html) Then Return ""
    EndIf
    Return $file
EndFunc


Func _DefaultDeliverToSnap($snap)
    Local $s = _SnapCsvGet($snap, "outputGEN.csv", 0, 1)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 2) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 2)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 3) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 3)
    If _SnapCsvGet($snap, "outputGEN.csv", 0, 4) <> "" Then $s &= @CRLF & _SnapCsvGet($snap, "outputGEN.csv", 0, 4)
    Return $s
EndFunc


Func _SnapGetCompany($snap)
    Local $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 1)
    If $company = "" Then $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 2)
    If $company = "" Then $company = _SnapCsvGet($snap, "outputGEN.csv", 0, 3)
    $company = StringReplace($company, @CRLF, " ")
    $company = StringReplace($company, @CR, " ")
    $company = StringReplace($company, @LF, " ")
    While StringInStr($company, "  ")
        $company = StringReplace($company, "  ", " ")
    WEnd
    Return StringStripWS($company, 3)
EndFunc


Func _SnapGetDelivery($snap)
    Local $d = _SnapCsvGet($snap, "outputEDICEC.csv", 0, 8)
    If _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9) <> "" Then $d &= " " & _SnapCsvGet($snap, "outputEDICEC.csv", 0, 9)
    Return $d
EndFunc


Func _FormatDeliveryInstructions($s)
    Local $x = StringStripWS($s, 3)
    Local $rep = "$1" & @CRLF
    $x = StringRegExpReplace($x, "(?i)(entre\s+[0-9]{1,2}h(?:[0-9]{0,2})?\s+et\s+[0-9]{1,2}h(?:[0-9]{0,2})?)\s+", $rep)
    Return $x
EndFunc


Func _FormatDeliverTo($s)
    $s = StringReplace($s, " Tél", @CRLF & "Tél")
    $s = StringReplace($s, " Tel", @CRLF & "Tel")
    $s = StringReplace($s, " TEL", @CRLF & "TEL")
    $s = StringReplace($s, " Phone", @CRLF & "Phone")
    $s = StringReplace($s, " Mobile", @CRLF & "Mobile")
    $s = StringReplace($s, " +33", @CRLF & "+33")
    Return _H($s)
EndFunc


Func _HtmlTemplate($tracking, $operator, $deliver, $ready, $delivDate, $delivTime, $carrier, $instr, $rows, $totalPcs, $totalWeight, $totalVol)
    Local $date = @MDAY & "/" & @MON & "/" & @YEAR
    Local $s = "<!doctype html><html lang='fr'><head><meta charset='utf-8'><title>DELIVERY ORDER - " & _H($tracking) & "</title>" & _Css() & "</head><body>"
    $s &= "<main class='page'>"
    $s &= "<section class='header'><div class='logo'><img src='" & _LogoDataUri() & "'></div><div class='head-right'><div class='doc-title'>DELIVERY ORDER</div><div class='datebox'><span class='label'>Date :</span> " & _H($date) & "</div></div></section>"
    $s &= "<section class='topline'><div class='line'><span class='label'>Tracking# :</span><span class='value'>" & _H($tracking) & "</span></div><div class='line'><span class='label'>Operator :</span><span class='value'>" & _H($operator) & "</span></div></section>"
    $s &= "<section class='boxgrid'><div class='box'><div class='boxtitle'>DELIVER TO:</div><div class='addr' contenteditable='true'>" & $deliver & "</div></div><div class='box'><div class='boxtitle'>Pick Up From:</div><div contenteditable='true'>Expeditors International France SAS<br>La Porte des Champs<br>Batiment B - Quais 21 a 27<br>95470 Survilliers CEDEX</div></div></section>"
    $s &= "<section class='sectiongrid'><div class='info'><span class='label'>DELIVERING CARRIER:</span><br><span contenteditable='true'>" & _H($carrier) & "</span></div><div class='info'><span class='label'>Commodity:</span><br><span contenteditable='true'>IT spare parts</span></div></section>"
    $s &= "<section class='schedule'><div><span class='label'>Ready:</span><br><span contenteditable='true'>" & _H($ready) & "</span></div><div><span class='label'>Delivery date:</span><br><span contenteditable='true'>" & _H($delivDate) & "</span></div><div><span class='label'>Delivery Time:</span><br><span contenteditable='true'>" & _H($delivTime) & "</span></div></section>"
    $s &= "<section class='instructions'><span class='label'>Delivery Instructions:</span><br><div contenteditable='true'>" & _H($instr) & "</div></section>"
    $s &= "<table><thead><tr><th>Expeditors Ref</th><th>W/o &amp; S/o Ref</th><th>Pack ID</th><th>Pcs count</th><th>Weight</th><th>Dims</th></tr></thead><tbody>" & $rows
    $s &= "<tr class='total'><td>Total:</td><td></td><td></td><td>" & _H($totalPcs) & "</td><td>" & _H($totalWeight) & "</td><td>" & _H($totalVol) & "</td></tr></tbody></table>"
    $s &= "<section class='signatures'><div><div class='sigline'>X</div><div class='sublabel'>(signature)</div><div class='sigline second'>X</div><div class='sublabel'>(name)</div></div><div class='received'><div class='received-row'><span>DATE RECEIVED:</span><span class='received-line'></span></div><div class='received-row'><span>TIME RECEIVED:</span><span class='received-line'></span></div></div></section>"
    $s &= "</main></body></html>"
    Return $s
EndFunc


Func _Css()
    Return "<style>*{box-sizing:border-box}body{margin:0;background:#e9e9e9;color:#111;font-family:Calibri,Arial,sans-serif;font-size:11pt}.page{width:210mm;min-height:297mm;margin:0 auto;background:#fff;padding:12mm}.header{display:grid;grid-template-columns:62mm 1fr 58mm;gap:8mm;margin-bottom:8px}.logo img{width:58mm}.head-right{grid-column:3;padding-top:2mm}.doc-title{font-size:16pt;font-weight:700;margin:0 0 7mm 0}.topline{display:grid;grid-template-columns:1fr 1fr;gap:18mm;margin-top:6px}.line{display:flex;gap:8px;margin:4px 0}.label{font-weight:700;white-space:nowrap}.boxgrid,.sectiongrid{display:grid;grid-template-columns:1fr 1fr;gap:12mm;margin-top:10px}.box{border:1.5px solid #111;min-height:34mm;padding:7px}.boxtitle{font-weight:700;margin-bottom:6px;text-transform:uppercase}.addr{white-space:normal;word-break:break-word}.info,.schedule,.instructions{border:1px solid #111;padding:7px}.schedule{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:10px}.instructions{margin-top:10px;min-height:14mm}table{width:100%;border-collapse:collapse;margin-top:12px;table-layout:fixed}th,td{border:1px solid #111;padding:5px 4px;vertical-align:top;word-wrap:break-word;font-size:11pt}th{background:#f2f2f2;text-align:left}th:nth-child(4),th:nth-child(5),th:nth-child(6),td:nth-child(4),td:nth-child(5),td:nth-child(6){text-align:center}.total td{font-weight:700;background:#fafafa}.signatures{display:grid;grid-template-columns:1fr 1fr;gap:14mm;margin-top:18mm}.sigline{border-bottom:1px solid #111;height:15mm;font-size:18pt;padding-left:8px}.second{margin-top:8mm}.sublabel{text-align:center}.received-row{display:flex;align-items:flex-end;gap:8px;margin-bottom:13mm;font-weight:700}.received-line{border-bottom:1px solid #111;height:8mm;flex:1}[contenteditable='true']:focus{outline:1px dashed #777;background:#fffbe6}@media print{body{background:#fff}.page{margin:0;width:auto;padding:10mm}@page{size:A4 portrait;margin:8mm}}</style>"
EndFunc


Func _HtmlToPdf($sHtml, ByRef $aNums, $carrier, $delivery)
    If Not FileExists($sHtml) Then Return ""
    DirCreate($PDF_PATH)
    If $carrier = "" Then $carrier = "Autre"
    Local $rule = _SP_GetRule($carrier)
    Local $vars = _SP_BuildVariables($aNums, $delivery, $carrier)
    Local $pdfBase = _SP_ApplyTemplate($rule.Item("PDF"), $vars)
    Local $pdf = $PDF_PATH & _Safe($pdfBase) & ".pdf"
    If FileExists($pdf) Then FileDelete($pdf)
    Local $browser = _FindBrowser()
    If $browser = "" Then Return ""
    Local $url = "file:///" & StringReplace($sHtml, "\", "/")
    Local $cmd = '"' & $browser & '" --headless --disable-gpu --print-to-pdf="' & $pdf & '" "' & $url & '"'
    RunWait($cmd, @ScriptDir, @SW_HIDE)
    If FileExists($pdf) And FileGetSize($pdf) > 0 Then Return $pdf
    Return ""
EndFunc


Func _FindBrowser()
    If $g_sBrowserCache <> "" And FileExists($g_sBrowserCache) Then Return $g_sBrowserCache
    Local $a[6] = [@ProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe", @ProgramFilesDir & " (x86)\Microsoft\Edge\Application\msedge.exe", @LocalAppDataDir & "\Microsoft\Edge\Application\msedge.exe", @ProgramFilesDir & "\Google\Chrome\Application\chrome.exe", @ProgramFilesDir & " (x86)\Google\Chrome\Application\chrome.exe", @LocalAppDataDir & "\Google\Chrome\Application\chrome.exe"]
    For $i = 0 To UBound($a) - 1
        If FileExists($a[$i]) Then
            $g_sBrowserCache = $a[$i]
            Return $g_sBrowserCache
        EndIf
    Next
    Return ""
EndFunc


; ============================================================
; CSV / SNAP / MAIL / EDOC / UTILS
; ============================================================

Func _TrimNum($n)
    Local $s = StringFormat("%.3f", $n)
    While StringRight($s, 1) = "0"
        $s = StringTrimRight($s, 1)
    WEnd
    If StringRight($s, 1) = "." Then $s = StringTrimRight($s, 1)
    Return $s
EndFunc


Func _LogoDataUri()
    Local $s = ""
    $s &= "iVBORw0KGgoAAAANSUhEUgAAAQUAAABhCAYAAADfhLRpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAACbVSURBVHhe7Z0JnI3VG8epkH1nGDN2svwrZSdkCZU1a4k2lUhSISkkUvyzVkryV6mUiopKpZJISouKKMqaMGPfef6f7zOd671n7szcOzN3zOX8Pp/zmbnvPe/7nuV5fuc5z3nOuVnEwcHBwYMs9gUHB4dzG44UHBwc/OBIwcHBwQ+OFBwcHPzgSMHBwcEPjhQcHBz84EjBwcHBD44UHBwc/OBIwSHT4uTGv+TYh5/KsU8+z3zpo0/l+LJv7CKfFXCk4JApceyzpRIfe7HER1WR+BJVM1cqVlniil8kR1+daxf7rIAjBYdMh+NLlkv8RXUlLneMxOUvk7lSzpISF/sfOTxxmsiJE3bRzwo4UnDIVDg6+02JK1JR4vLGSlzBcpkr5YmR+Ao15diCRXaxzyo4UnDINDj89AyJK1lN4vKWTqyQZzrliZU9F9WVYx9/Zhf7rIMjBYczj+PH5ciU6ap4u7OXSJg22OlMWg55YtR6OfruB3bJz0o4UnA44zgybabsbdJG9rXuIvuu7po4te8hextdI3GFK0hcgbKJlTacKXeMxMdeIkf+95pd7LMWjhQczhhOHTosh4aPlX3NOogcOSqn9u5LnI4eleNffyt7G10rcfkyeFqRM1r2Nu8gx7/9wS76WQ1HCg5nBidPysG7h8juLPkTSCEJnFj9q+xt0DrBhLeVNpwpZ7RaLsRKnGtwpOCQ4Th15Ijs73KLjvy7c0XLvuYd7SyK48tXyp5azSQuV6nEShvOlLuU7GvTXU5u2W4X6ZyAIwWHjMXx"
    $s &= "43LwgeE6EsflL5skKRxbvET2VKsncbmiEyttuBL+ihwlZE/1BnJi/Qa7SOcMHCk4ZBhO7dsvBx96TOIKl08IBCpYLjEpHD8ux5Z+LfFlLpXdWQrJ7mzFT6ccJRIrcnql/GUlLl8ZnTKcWLvOW+xzDo4UHDIEJ9aul32tOieM/P8SAskmhVM7d8vh/06Vg/cPl0NDHzudHnlcDg4YqoqbSKHTmihP3li1YE7tP+BX7nMRjhQcMgQsO+7OUjCRQtqkkBzUykhvhyNThgJl5eDDY+TU4SP2K89JOFJwyBAcmfGKTgFspQyKFA4fkcNjJiSY+OkZp8ASZ65ScnDYaF0NcUiAIwWHDEFqSeHkjp1yoFdfiSdwyTPtSHMilLpoJTk84Vk3ZbDgSMEhQ5AaUji5ZZvsrdVMdmcrlqDEWArpkfLESnypanJk5mz7lQ6OFBxSg0ULF0mX1l3kxg43+qXubbrL4LsH29kVoZLCyc1bZW+t5hJXpIKO6OmWClfQlY2jc96xX+nwL8JGCpv+2iRzX50rc16eIwvmLbC/Dgnbt26X12a9Jm/OftOX/joHI80yA3b+s1NK5CghhbIUkqjsUVLywpK+VPS8otL40sb2LYpQSeFU/B45+uY8Ofreh+mb5i+UY59/Zb/OwYOwkMJniz6Thv9pKPmy5JOC5xWUMgXKyC1dbpGbO98ccrqt221yTaNrpMh5RaR49uKaCmctLL069pJ9e/fZr3YIM44fPy7vv/O+jBo6SmpWqCnlCpbzpZjcMXJV3avsWxShkoLDmUNYSOGPdX/I5Ccmy6C+g6RxjcaqxPaoYlLxbMWlQJYCAVOhrIUSyCBbcSmTv4yULVBWhS82b6zUqlhLflz1o/1qhwzEgQMHpN/N/aRkzpKOFM4i"
    $s &= "hIUUvFi/dr289dpbcsXFV0iZfGX8RhY+X3nZlfL+2+/LJx98Ih8v/NiXsDamT50uA+8YqILXoHoDKV+ovI8U6lap60ghE+Dp/z4txS4o5kjhLELYScFg9szZOuJ7SYHP7wTh8Dl16pScPHlSZr84Wy2O0nlLO1LIJEgXUmjSxs7ucAaRYaTw+qzXA5LC/Dfn21mTxIkTJ+SBvg+oEDpSyBxIEykQd3BhSTk04gk7u0MIYMA8duyYfTnViChSACu+WiGXlL5Ealas6UghEyDVpEA0YdGKcqDP/XrAikPoOHLkiCz5dImu8j391NPy0L0PyW9rfrOzhYyIIwXQs2NPubTMpfLjd44UzjRSRQoEIhUqL4cmPKMnKzmkDt8s+0bKFy4vF0VdJG2vbCsTx06Uzq06y58b/rSzhoSIJIW9e/dq7MPunbvtrxwyGCGTQr4yEl+ymhx+fpadxSFE3Hv7vfL4I49rm7es11Ju7Xqr3N/nfnngrgfsrCEhIknBIfMgJFLIWlh254mVwxOm2V8HjT1xe2TxosWy6c9N9lfnFDb/tVluaHeDrPl5jdzR4w6pXbm2dG7dWT56/yNpclkTdc6nFhFFCgcPHJQ98Xvsyw5nEEGTwrMzZXfWInJw8Ej7q5AwdfxUickbo1Gt5zqG3jNU3n3rXWnbpK1OH5rWbCrXNr5Wbup0k501JEQMKaxbs04ri3mUGSMZN23cJMMfGC6vzHjF/irsMEu23pRRCJYUTqxZJ0fnLRQ5kbhslDeYkY14lsvKXyal85VON1Kw2y2YciSFnTt2ysP3Pywvv/CynDju/5NyqXm+XTa7X4ntaVarmbbJgN4DpFX9VtKmcRuJj4v3yxcqIoYUHh3yqOTJkkeFDtMpEPA1bN+2Xf7e/neS"
    $s &= "ie937dzlu4dlzl27diV7H9/t3hXYf3Hq5CnZ8fcOnc8VzFJQTbeMAntCvv/2e/nv6P/qikztSrU19Lh5reay7Itl8sf6P0ISwtQgWFIIhMOHD2twW52L6sjzk5+3v/YDCle/an2JzhWt73rh6Re0T+y+Iv2z4x/7dj/s2L5DndRMQ+pXqy+1KtXStkO5+t3ST7764iv14ttKmBywYjHj82bJK1dfcbXKhF4/eFAWzl+oz768wuXSok4L3cdDfhv0FdHAy79crhvLLi9/uU4LSKy2Nbq0kcx8bqb89utvvn5FF54Y8YTK38A7B8qvq3+1HxsyIoIUaECsBEKeafAtm7bYWRRTxk3RsOpWDVrpfolACWbt1LKTNj745+9/pM+NfaRJjSbSumHrRPlbN2itSsYzibL0YtGCRcrQbZq00ShLlIJODzcYhYYPGq7OJdoQpURZCOoixeaJVYJCEJ956hn79pDw3YrvVECHDhgqQ+4eIls3b/X7PrWkgMLRX4Srs7lq/KjxdhYFwj9vzjw1kbEQjOw0rN5QTWW7v5AP0tuvvW0/SjHvjXnSsn5LqVikosoTIdoxeWK03fjL3hraDoV8+L6H1UINBihm0fOLSqncpaRD8w4StztO/vzjT+nRrofKBony85f2QuZs0FcQAOUyfUqijGwJ4N4SF5aQqiWqyqznwueoPbOkkL24vP164M7zYsE7C+Q/Mf/R+++88c4kAzWen/K8KmjVklWlVK5SKqTehLJccckVyqqGWA4fOqybezC9qkdX97uPTq5YtKI0r91crm97vaz6ZpVaBs9OeFauu+o6XQ4iwpK81IdOGzlkpI5EKNOw+4ap4jISMsp5U51KdeS6FtfJjGdmyPcrv5dDhw7Z1QkIdiminLQduxWx"
    $s &= "EPr07KPTFnamMkJt3rRZvdJ8F5UtSu6+5e6gHHNHDh+R1T+sllnTZ0n3a7vr6FmucDm10KqVqibReaOlQtEK0viSxvLgPQ/Kpx9+qu1g6p8SKWCVQSo4w65ueLXuZUGJKhSuoKHwXjBFfOOVN/R55DMWgkmqxFb/msS+mdt73u73vI1/bJS+N/WVsvnLah/TduzoxEJZsWyFxMXF6RLfyMEjtc/M3pxqJarJ6GGjA5rkWDorl6+UEYNHSOVilbWcyBiyhBxcVu4yicoRJeUKnC63kROsEwOsCVYMjCwh62wgnPX8LLUMIF76g3KbwZTP4cIZJQU6tl3TdqqUjz74aKLEdUZEYhJgWe5n3p4SEDoa3mygMh1R96K68kMSv/YDOdzW/TYVUjZfITRPjnxSln621M+HgWCjZBAGz/TWh/dhjja6pJFE54zWTuZ5JCNkMD0jkloWeWL0Gvf27t5bpwPJgSXYZjWbqaAhfAj1siXLVDgDARO5SlQVJRAEdc3qNXYWH3DgsscE4aadqR/k+uIzL8r/pv1Pfvr+J23Xmc/OlP639tfRnbJ7R++USGHyk5OVaBgFTf+g7B0D7H3AochGOr6nrt52Nm1NP9mJ57Ju/+XiL33PWvr5Ut21y4hr3tn16q5Jbr9f+8taNfXVgshXWmUBJd22dZtfvulTpqsC0w5eWeN/2pD3UCb63ysrlIOBwmD196uVOIzcYUHZIFBpypNTfLKUVmdicjijpMAGJ5QHc43RiMScjM8IHQnBMIIXLCmAhfMWaoeZe2nwatHV1FsbCCuWrlDBQRDIP/6x8XI0icCaQwcPqQK/OvNV3yhphIGRqEf7HjLh8QlqqmI1MG9+bvJzusHrrp53SdmCZbXe3vtQlG7XdJO/NgQWVEZx"
    $s &= "TGiEhnsQDkaSlMCIjEAiSFgPNjDP58+drxvWaF+IGtMaS2DD74F/+4Dt0y9MfSGhvh5lSIkUnpv0nE7fKhWrpP3BPUmRQvzueLXmtmzeotMX6mvair+Y2rSt5rGSl1yZehLsZqY4tAVTg6QIwYBpLe1gdudyP33qxYvPvig1ytRINDjwmWkk7frzjz/LsIHDfNMCEu//YP7pH6sdO2KsygP3MaCwzBgI1AsZpZ/OWlKgsTGx7rjhDnUcwbyTnpgkt99wu9zQ9gbp0aGHnuZjOiYUUsAku+e2e3wjsQptnhidXuzb4796geDA3JhnjEpMD4IBTh2vWcvz7+l9j53ND1gknEeAoJqRxNxfLFsxad+svezds9e+TRWK55uRDi/z/n377WyJAEEZq4SdprYjCoca+0h8o2jOaJk2Mbg4AgTTKKvWPwVSABAtIyEWCXVPihS8YNcsVpeRGe795adf7GwBwRQHS8nbR1g7wQD5MSY7gwtOXDta8N257/qRHIMC5WNK6MUvP/4iT49/WpMdnv/I/Y8kTH3yxEiL2i2SJGOA5czgyUlX4cIZJQU6+uUZL9tZ5eiRo3Jw/0EdkVFYRi0zfSC+O1js379fOjTr4EcMCFXX1l39FOqj9z5S5w0szvw8WK+zTQpKKBODIxSmIa/+71UVIi8x8Lw3X0m83Eb0mmk/FBFTPxAgHZxcnHXAPB1rTEfXQuW0ft7yUU+eq9bUv/Ne5tNYJcGAKYWZS6vCBUEKgDl8jbI19L2pIYVqJasF7YMZ3G+wz/rjb7druwX0/AcCBMC7TP0giGsaXqOWkhfImJED2gDCTmkq6AWKzn28hzbBN5YUGDBYVcJKChfOKCnwOZjVh61btmpDY8IxF0xqeTAQWF1A0I3ZTcMzKg7q"
    $s &= "N0g7l7k08eKU5cb2N8rPP/1sPyJJBCKFQPPBpIDpjtcaJTeCZ8oGaRiwBIXZaMxUFL1Diw7S67peMubhMer74Dl48HGmYQ3hAKweU12fjaCRIEfugTgAzkljfZCPNqatgwVth5IZayFYUsAPYvxE4SQF2teQD/fizxk3apydLUnQB/SFmepB3hcVv0hJzYtOV3XyyQFt0b5p+5CUllUOY5FQP8rJlCOYOoYDEUEKgIbDx8AcGe96KDiw/4D07NDTN2IYBbyvz33S9PKmUjJHSV1hCHYEMUgrKQCchESjMZc3ZcNqgcwMWAr1+iAYnXEC4kRk3utNCC3fkS6OvViVzySuIeRg4+8bddpgfCI4FhlVQwVKZvo1s5HC6IdG+/l8qOMbL79hZ0sWEKepH+9mpYRVES9sUsB5TrxEsMBX4T2AiPdAQDitP3z3Qzl2NPBqW7gQMaTAtIG5FKN+UnEKyWHjho2+EdM0PAKDsuHc+fKz097qYJEepAB++O4HnyWjz8kb62d+Erlm5sW8g7MrGa1Wfr1Svln+TdBp+dLlGgQEWOY1c2FS9VLV5YtPvvCUKjikJk4ho0jhrl53+e5D6aijd1UiGHA4UJHzi/j6hucRm+JFWkkBeb7y8it9x9p55ZP+ZuXD9gWFExFDCqzPYxKzbBhKg3vBkhpLYt4lLjp50thJdtagkF6k8Pe2v9UCMpYMSmpIAVMfa8aQGUKX1DHqoQBnIqcv80zey7F4qUFmJgWWTjVO4N86svJhgtaCBUuB17e53qewlIMjAr1IKymAzz/+XC0DdMJYsj49uaC4BtHhjMdfFG5EDCmkF+ylNASTaEavuR4s0osUmLZwarV37mpIYd++fRokZcqL4hHoktbwZVZ6fMqcyUkBc93rE0ot"
    $s &= "KUCu639bb2dLEXf2uNP3/nCRAmBVgjJTR2TBDAQqp3lLKzm0u7KdrAvzr2Kfk6RgBBjPP155nDx4kIMNaTVIL1IAEx+fqEuS+hzP9IHAKV2d+Hd1IL0shUghBcgPP4g3TiE1pECbsiRL3EAoIMYE095YceGYPnjBUjrLxPhDqpSoohGpRr5MXA/7III52zS1iFhSYLmShnlp+kv2V0li1cpVGvqLgOBYxBFnlgNZriOqMKkQ6kBIL1JA8AldZjRC6CmXmfvjAWf50TjMEE7K6V2dSA2Ie8DxxjOZb/PO5UuW29lSRLhJgWVT4gVM/UMhBVZizLSDRDkDLfcmBwLUTDuRKC9hzV6kJyl4wcao26+/XZ/ttRp4B/4gfE3hQMSSwrcrvtUOJzosGLAngBDhIlmLSMcWHXW9t3e33lIi+2mztEKRCrpmHKzCpRcp4C8haEkdn7mi1ZHojRVgTl3s/ATFg8QgtrSeK0Ho76VlExST5yL4bEsPFRlBCgNuH5AqS4FAN2+MCvJmrxykBE4Q964+sBfGHqXDRQoAWVz80WKNkDTWCokBhOjYpKJu04KIJQVGT+4PRpA57JWOwjzHoWcUihBadkAagUPhUOy5s+fajwiI9CIFdlviGUdJKINt/bC5ipBtY9XwHkarYMkrEDas36Cbw0zdeTdm6c8/hGZeZwgp9B6QKksB4Fz0OWlzlVLrIdglPt79YP8H/Xw9bAxb++tav3xpIQX6EJJJabv310u/1iVlIwO8h7galtvTGxFJCiMGjVCmhDmfGvOU/bUfMMPZysq7GDUohxfsCkRwvCMRChrMiBKIFIKNaDTAyYijk3spA/Pe33/73S8PAoOgGeEkb6i/kBVoWoQDzRv/gOLh5WbfQbDICFLgnABv/6Ac"
    $s &= "wSoD8S2+EbZAObUG1/wSeG+BDeJhIGPj5IVUWtZtaWdLEykQQEfsw7RJKYeWE2RnAth4D6HywbZDKIgoUmB5jn39dBKMScMsWbzEzubDti3bpOd1CUFLdCgmeqDDXtmsZCL7jHATQMR++ORgkwLPQIBDAbtBzW5DlN0mLQP2Y5DHJ6C5S6kwBjONQMlx1gU6FYr4D6+JDUnw3GCdrpw0ZDzzwZICo56ZuqRECgAT3rv6QJAWu1eDAcvQXrmj3ThzI25Xykt7BLeZ/qWsREcS9GXDJgWc1uaQlZSApQDx1KtWL8nDgwzY5m0IDnlmf1Aw+19CRYaQAhuQHnvoMT/hMxXjx10ICU2qcgg9HUhwkcb/eza3EA5qm3IG7NtnHwN5NG+OEjoyBgLzMmVhT/wCjc+uxeQ2p9ikoApboJyGHKcU+44FQz5GTAQJz3JyUw/OA0B4vPNKVeCWnTToCMW3lykZZZk6MdUiGnTs8LF+3wP2L2AZefuG/+tVrac7EVk/12efPP1s3sOqCNuq2V1pzPNgSQErjLm5Ttfyxkq9KvXk15+SDs754N0PdLcsK0W8h/fhF2IbPGWjLQl9Jxyc8ym8hIFcEWdgnIX0EeUcNzL5cGfe6d3oVPSCojLu0cT34JtiW7pvGvbvr5cFS6qQgp67kD1K7rvzvmSjarEOjaVA308dN9XOki4ICykQYIHCfLzgYw0rRcDs7aUmmYg61pBnPD1D8zNaktjQc1Wdq9Q7Tmd6lZZrNKb38BCUgD3vjCxGeMzIakakQNGQCBZLdHYZMYsRPsKMA+1ctEmBhNJCXIQis3Ua4WCDFzDHbWHd0CY8Hz8HdSFvSuDgDiMQ5n0II+2HAM+cNlPmvDRH244zDDGd+Q4lQOCSwntvv6eHfjKi"
    $s &= "ciiqUVYS7VepaCXdWciz8d5DXvgfKL+3zegfpkLUke2//DUOUw4poS3ML5J77yMwiKg9/BmByJRlQQJ7vBF/5n5Mb5PYkk572kfi8U4Ul4GBulEnnsUKTCAzn/waAv7vuRcoLPN370iO1crUk+3oXn+P9kmuUnowCwewQFS0BVu1A/mAuIYFwv0sP7K3x7b+kB9+sNls5aYeTDPZGBUOpDspEAGG8w9nUHTuaA1NpqNo2OQSTM5P13OegjfB0HZeGo8zF2x/AtMAIhaLZE/4pepE9+WI0g4YcNvpdWYEF+djwfMKJrqHlYnC5xeWqJxRevyX7YC0SYEOI+qSuR6jbVSuKKlTuY6O5ph6nJajBFkgVuvFKMuGJnOiU0rAouHgFIJnKKtu9f5XORB0ysx1RlUsA04g4pQnzgQM5FPwgqkWoyP+F/ZP8Bz12+SOUeLhF8DpD+3PC4upknEIDnX1bubhL3VmhYTYB0ZzhJyVH6wi9muw/Gu3M+1fOaqylpcR2S4vy3MckoOccI8pF4k+4BrE0KV1F1VGG1gSHGHGahV5uY/8DEZsKGMAQsE5YYs8RmbZXMbUzd7gNGbYGA0qQ2ZtuSEVyFpAv2fgMntPsNpsnDxxUo/NM4fv0Ib4dYixgIBxsjIwUVbyYFEyuK1asUrvDQfSnxQOH1GzE5MV4YUg0jsNvH2gjpr2QRmYpXiXGRXte0yiXExlDFBsphnJ3cN3rBdzAKcXNikwUnJiEMDyYGmRUYsTkthfAQlgXvOsUQ+OCnqLtg0UhqhGotuYH3uFkqg33sMSFlGSqTn5mpOHiA2459Z79Dns8ccfM6jvIOl/S389rRiwAkAdIR9DSFgKBPv06tjLdy4DS670CwLOFNBuX5P4rm+vvjp3trcnA6wI"
    $s &= "8iBXkBJtSfk4fIbTsFhmTQmM7pSD+1BYyK7weYW13fiLkkPuKB4ylhQ44Ia9FXYd7PrwLhJnMiZ1BsScl+doHSBR2pEyFcxaUC0x/ubPkl/LSmxNIL9QeiPdSeFcQiBS8PoFIEhMTnMiEP+zz+H4icQCHyqYinAqNY4vDhJhLZvE/1g/vI8Ar7QA09aUO6kwcAiEOArzbnwfW/7aEvSZDKkBdcdnZMoG6YRCsJAq93LwiWk3b/thlcTHx8spSdl6Sw9QHywZHNt2eUgEKVFWpjrHjidv8aUHHCmkASmRgoNDJMKRQhrgSMHhbIQjhTTAkYLD2QhHCmmAIwWHsxGOFNIARwoOZyMcKaQBjhQczkY4UkgDApFCqBuiHBwyGxwppAEEqHj3IvA/4brA3ofg4BApcKQQAji5if0E/KKP+Qk3wlLtRCg1m6/IN6T/kJDPKHBwOJNwpBAC2BVI/Dlx82y2Ig49UDIbirAc2G8RzG8+OjhkFjhSCAHsVmObLiG/Qaft/wR90o+DQ2aAIwUHBwc/OFJwcHDwQ1hIAc+7vWuNa+eKR37z5s0ycuRIWb489CPT0wOBth3boC+8+Uyfkbge6ECQSAY/qnP33XdL3759ZffuhCP5Vq1alWxbPffcc0HJLL9u3r9/f7nrrrtk165deu2FF14I+nDZzIawkML69etl1qxZUrlyZSlTpoyULl1aevfurZ0QTCN7gXBu27bNvpyp8Ntvv/mUCAFp166dZMmSRVq1amVnDTvi4uLkjjvukJiYGClbtqzs3Jnw+xE2fv/9d7ntttu0f2JjY6Vly5by1ltvyRtvvCEPPvigNGvWTIoUKSKDBiX8IG2kAjK4/vrrVf5GjBghw4cPl0mTJknnzp2lRo0aSSruJ598ou2H"
    $s &= "zCaFtWvXSocOHWTcuHE6CPDsgQMHSoMGDVTm+WGXSERYSAHQ2Llz51blyJkzp7z33nt2lqAwfvx46datm30502D+/PkSHR2tvx5twGhEve+9916/vBmF7du36/tJf/+d+LgxAwjswgsv1HyPPPKI33dYO1dccYX07NnT73qkgdGe+n399dd+1z/66COpUqWKHDgQ+DRkCIT77rsv6WPs6tSpI7Vr15Z//jl91gSDw2OPPaaE6kjBQlpJATP22WeflQsuuEBHscyIadOmSZ48eSRv3rx+x4chaFg3oVpF6YXUkAKjHMCc/uGHH/T/d955R3r16mXdFVkwBD12bOJDa7Eg9u5NfPYm1lWjRo30PuQPqyEQChcurMq/YMECv+tr1qyRSy+9VK2USESGkwLz7AceeEDNUsxVTrh5+OGHZfDgwfLoo4/6zN2XXnrJJ9jlypVTk3bRokW+53Pv888/L++//76yOiMbeO211/T55Of/efPm6bsQ+hUrVuhzeRcJE/rtt9+W+++/Xx566CE/HwDC8vnnn8vrr78uH3/8sZbxiSee8M1Hf/75ZylQoICWL0eOHNKvXz9ZuXKlCsSQIUM0P+/yAmV78sknZfHixToyz5w50zftmDNnjq9dyMc7GaX4TD05+9Lgu+++k6eeeko+++wzbTPyfPnl6Z9YTw0pYP4CLJ5XX31V/6cvEPhff/1V22fo0KFqudFulJ9rlAWQj7b+8MMPtY589/333/veg4lNu/Dd6tWrtS2p77Bhw+Svv/yP1WNeTv0+/fRTVWbeuWNHwhmJjL7UlWt8z3OYviUF/AjUj3qOGjVK+8iAvgrkO6G+KPV5552n9zLyB0Lx4sX1++zZs8uECRO0HMaXRn96rcdIQoaTAgLC6MN12Bjl"
    $s &= "Hj16tC8vcz8AWaCsXKtfv778+OOPKuwoZfv27aVYsWIyd+5c7cD8+fPrHBrhYoSuVKmS3leiRAkVuooVK+pnBBVFhf3NZ3wfV199tU9wvvjiC30/gpAtWzYVYt7ZqVMnzVOtWjUVJoTz5ptv1mtYC0uWLFEl4nDVAQMG6PWrrjp93PnEiRO1HczoxLtNGRCkrVu3qjnLNRxWlJO24DPp5ZcTzkWcPXu2nH/++VKzZk2t68KFC/V7RqylSxPOKEwNKfAuyAqypK5e4Kdg7kw+FIDR97LLLtPPKC1TKP6HNHgmz+BzVFSUj8ghSa7RV/QJpHfxxRfrtWuuucb3LpQWsxySZaTFPCfPDTfcoMTQsWNHlRWesWfPHu1rFPiPPwKfbAxxmrYgQeS0HQNGoJ9cY3BBvrgPXwv38L6ffvrJzqr+A/NcZAqZ5N5QreLMhgwnBYDwcL1Hjx6+a7A416677jrfNUYXrqG0BozemOtYDxs2JPwmA55e8sHW4PLLL9fPCDJAYbEmACNarly59HsEEyC4PJNrppyMYnw2fgGj6HQ+1gMwwo+ge6cPEADXDSlgoRjhYcoBIAE+Z82a1TcnRRm4hgVh0LhxY732zDPP6OfJkyfrZ5Thzz//9D2HxHcgNaQAQeM/wDncvHlzO6t618nH6AggY/oHYpoyZYp+d9NNN+l39INdJsrPZ4ia9gDUk2vUGyAzpr5YAgCfAASENUQ/8R2kiEUCxowZo9ewUgIBxceqgaBMmUwyUyYvsMogPWCeTcJ5aAOSgqyMPJkEOUA6kYqIIwWUimt0BKMIwswIi3AjQMBLCvZoQDkoD98bUsDpxGjPNVNOOhzv8i+//KKrKcYqwHrAQw+MJZMSKTCaGoExyu1V5i5duui1YEgBa4Qy"
    $s &= "MaJCikx5zHOmTk04STo1pIBlhA+EZ2LN2GDkJp8hBS+wpGgriIL7UTbz/lBIgbphiXANCw1QJiwVpk9McfgOIsXKwBGKFchnppGBgAnPvbQD8kFfmbLR52ZgMaAvaFPKwjTH5C1VqpRvOgTobxLyxVQS+cDHYPLTTmZ5MtIQcaSA4HONDkDgUGyUGgvCLF2mBymgWCxdYZHceuutWi6+Tw0peJUkECnUqlVLrwVDCoBpEyYw0zCmFeY5aSEFM2qiRExRbCRHCrQxxEc98OUw/THvD4UUsJhQcK4ZS8ELnm2eiwVH/5HwLQRSQMgAP4YX69at0yVb85xrr73W9x3tyjIl1iRWCc/FerLbFzC9YNrkBc/GoiAv9XjxxRf9vo8URBwpmCWmokWL+jnX8JozxwRpIYUPPvhA57ItWrTQz8QcADOdSIoUeL/xNtukgDPTCFYgUli2bJleS44UWIkBmLSYz4UKFZKvvvpKld48Jy2kQPsD/BvGmepFcqRg/AWM3MDUnxQKKTA6m/Lga/EC0sX3ZJ7LlMUA/04gpx6kUK9ePdm06fSviAGsQOMT8fpPIH/bOYzD0LyT2A18XQDiJL8N+pK8OClfeSX8v9EQDkQEKSBseMSZyyOwpkMJDsKMxnFHIA7KDQwp4PQJlRS49s033/gEgZEDBybzbD57ScE4mvLly6fXjDJjYXDdkALK17RpU71mSIH38BnBMh7r5EjBRNeZ56A8jGjGx0BKCyl42z0QzNJeIFIwoyPTOEZs438ghUIKwDhg8d1gvVBnVjuwEjD1mzRpot8THIR5j+LfeeedieIQAO2KA7Z169Z+QVwMHjyHqYdZPWFK1r1794DBTMbJSqIcgP5GvimjV8aYSpCPZyXX9pkZYSEFPMHMUYkI"
    $s &= "Y1WA+RjebVYQWPPF9EU4qlevrmvheM3xDyBwCAidDRhh8QATdceIzZwV0IEoB1MIlJboNBgeAcF6qFu3rjp7sDBY3TACwWiDw6pkyZLqeGK5imt4ss01BIxRByXg+cxxMdNxVKL8PPfGG29UUxdnEvdQP96F+ci8k9US6gJ5ffvtt/pulqt4N0TFfTjlsEawGACWQPny5fU+nFeMfjNmzFDB5RpzXYQPwsFKYCRCkbFgEFqIrm3btkpolJ36kMgfCCaikf4xCbLGJLZDf5lj8376rGrVqkrQ3mVElg9pKywmRl76DbMc5ejatavWkQA02or34IikjakndaPexhKCqLDACAgjf5s2bZR8kRNAGzds2FDbAFmh75MKR2a5EYcszlPkj4GHJVzeT/lMPAZ/aUNWMaZPn+57Fn3Ae+lb08/II32MXCG/yCb1R955dp8+fbSPA01nIgVhIQU7rh6YuHo6yqwNe+PsvddC2TdhnmHAc0xe/tqfvc8239nXDAhC8i5F4UwznW2Xx3zmflOeQHVhFDfk5gX3mLymXIGuAYQVLzkw1xBUrpPf69+w+8GA+7z5zOdA6/Y801tfb1kMUPyNGzf6rqPcKBgw/W5gnmfLgQ3ae8uWLYneZcD9tiXoBfUxZIN1gK8Df4KtsKbNvH1nEKiugPBnM31l2ggR82xvdGOkIiyk4ODgELlwpODg4OAHRwoODg5+cKTg4ODgB0cKDg4OfnCk4ODg4AdHCg4ODn5wpODg4OAHRwoODg5++D8bEYSTuWkTBwAAAABJRU5ErkJggg=="
    Return "data:image/png;base64," & $s
EndFunc


; ============================================================================
; CMR_FULL_FUSION_V1 - API helpers Dispatch <-> moteur CMR inclus
; ============================================================================

Func _ComErr()
    Return
EndFunc


Func _Box($x, $y, $w, $h)
    Local $idCard = GUICtrlCreateLabel("", $x, $y, $w, $h)
    GUICtrlSetBkColor($idCard, $C_CARD)
    GUICtrlSetState($idCard, $GUI_DISABLE)
    Local $idTop = GUICtrlCreateLabel("", $x, $y, $w, 1)
    GUICtrlSetBkColor($idTop, $C_BORDER)
    GUICtrlSetState($idTop, $GUI_DISABLE)
    Local $idBottom = GUICtrlCreateLabel("", $x, $y + $h - 1, $w, 1)
    GUICtrlSetBkColor($idBottom, $C_BORDER)
    GUICtrlSetState($idBottom, $GUI_DISABLE)
    Local $idLeft = GUICtrlCreateLabel("", $x, $y, 1, $h)
    GUICtrlSetBkColor($idLeft, $C_BORDER)
    GUICtrlSetState($idLeft, $GUI_DISABLE)
    Local $idRight = GUICtrlCreateLabel("", $x + $w - 1, $y, 1, $h)
    GUICtrlSetBkColor($idRight, $C_BORDER)
    GUICtrlSetState($idRight, $GUI_DISABLE)
    Return $idCard
EndFunc


Func _SectionTitle($sText, $x, $y, $w)
    Local $id = GUICtrlCreateLabel($sText, $x, $y, $w, 20)
    GUICtrlSetFont($id, 10, 700, 0, "Segoe UI")
    GUICtrlSetColor($id, $C_TEXT)
    GUICtrlSetBkColor($id, $C_CARD)
    Return $id
EndFunc


Func _MutedLabel($sText, $x, $y, $w, $h = 20)
    Local $id = GUICtrlCreateLabel($sText, $x, $y, $w, $h)
    GUICtrlSetColor($id, $C_MUTED)
    GUICtrlSetBkColor($id, $C_CARD)
    Return $id
EndFunc


Func _HeaderLabel($sText, $x, $y, $w)
    Local $id = GUICtrlCreateLabel($sText, $x, $y, $w, 18)
    GUICtrlSetFont($id, 8, 700, 0, "Segoe UI")
    GUICtrlSetColor($id, $C_TEXT)
    GUICtrlSetBkColor($id, $C_CARD)
    Return $id
EndFunc


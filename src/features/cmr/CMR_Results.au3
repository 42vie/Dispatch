; ============================================================================
; CMR_Results.au3
; Liste des resultats CMR (BL generes) : ajout, effacement, rafraichissement
; de la ListView, selection et actions groupees (mail/EDOC).
; ----------------------------------------------------------------------------
; Deplace depuis StateService.au3 (bloc CMR_FULL_FUSION_MODULE_APPENDED) pour
; que le code CMR vive dans src/features/cmr/, comme le reste du projet.
; ============================================================================


Func _AddResult(ByRef $aNums, $html, $pdf, $carrier, $delivery, $company, $snap)
    Local $n = UBound($g_aBL)
    Local $nHTML = UBound($g_aHTML)
    Local $nPDF = UBound($g_aPDF)
    Local $nCarrier = UBound($g_aCarrier)
    Local $nDelivery = UBound($g_aDelivery)
    Local $nCompany = UBound($g_aCompany)
    Local $nSnap = UBound($g_aSnap)
    ReDim $g_aBL[$n + 1]
    ReDim $g_aHTML[$n + 1]
    ReDim $g_aPDF[$n + 1]
    ReDim $g_aCarrier[$n + 1]
    ReDim $g_aDelivery[$n + 1]
    ReDim $g_aCompany[$n + 1]
    ReDim $g_aSnap[$n + 1]
    $g_aBL[$n] = _JoinArray($aNums, " / ")
    $g_aHTML[$n] = $html
    $g_aPDF[$n] = $pdf
    $g_aCarrier[$n] = $carrier
    $g_aDelivery[$n] = $delivery
    $g_aCompany[$n] = $company
    $g_aSnap[$n] = $snap
EndFunc


Func _ClearResults()
    Local $nBL = UBound($g_aBL)
    Local $nHTML = UBound($g_aHTML)
    Local $nPDF = UBound($g_aPDF)
    Local $nCarrier = UBound($g_aCarrier)
    Local $nDelivery = UBound($g_aDelivery)
    Local $nCompany = UBound($g_aCompany)
    Local $nSnap = UBound($g_aSnap)
    ReDim $g_aBL[0]
    ReDim $g_aHTML[0]
    ReDim $g_aPDF[0]
    ReDim $g_aCarrier[0]
    ReDim $g_aDelivery[0]
    ReDim $g_aCompany[0]
    ReDim $g_aSnap[0]
    If $idResults <> 0 Then _GUICtrlListView_DeleteAllItems($idResults)
EndFunc


Func _RefreshResultsList()
    If $idResults = 0 Then Return False
    _GUICtrlListView_DeleteAllItems($idResults)
    For $i = 0 To UBound($g_aBL) - 1
        _GUICtrlListView_AddItem($idResults, "", $i)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aBL[$i], 1)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aCompany[$i], 2)
        _GUICtrlListView_AddSubItem($idResults, $i, $g_aCarrier[$i], 3)
        _GUICtrlListView_AddSubItem($idResults, $i, _ExtractDeliveryDateTime($g_aDelivery[$i]), 4)
        _GUICtrlListView_SetItemChecked($idResults, $i, False)
    Next
    Return True
EndFunc


Func _SelectedResultIndex()
    If $idResults = 0 Then Return -1
    Local $aSel = _GUICtrlListView_GetSelectedIndices($idResults, True)
    If IsArray($aSel) And $aSel[0] > 0 Then
        Local $idx = Number($aSel[1])
        If $idx >= 0 And $idx < UBound($g_aBL) Then Return $idx
    EndIf
    MsgBox(48, $APP_TITLE, "Sélectionne une ligne dans l’onglet Résultats.")
    Return -1
EndFunc


Func _CheckedResultIndexes()
    Local $aIdx[0]
    If $idResults = 0 Then Return $aIdx
    For $i = 0 To UBound($g_aBL) - 1
        If _GUICtrlListView_GetItemChecked($idResults, $i) Then
            ReDim $aIdx[UBound($aIdx) + 1]
            $aIdx[UBound($aIdx) - 1] = $i
        EndIf
    Next
    Return $aIdx
EndFunc


Func _ActionIndexes()
    Local $aIdx = _CheckedResultIndexes()
    If UBound($aIdx) > 0 Then Return $aIdx
    Local $idx = _SelectedResultIndex()
    If $idx < 0 Then Return $aIdx
    ReDim $aIdx[1]
    $aIdx[0] = $idx
    Return $aIdx
EndFunc


Func _ResultsSummary()
    Return "BL générés : " & UBound($g_aBL)
EndFunc


Func _OpenPreviewEditorByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    _OpenPreviewEditor($g_aHTML[$idx], $a, $idx)
    Return True
EndFunc


Func _RebuildPdfByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aHTML) Then Return False
    Local $a = _StringToArrayBL($g_aBL[$idx])
    Local $pdf = _HtmlToPdf($g_aHTML[$idx], $a, $g_aCarrier[$idx], $g_aDelivery[$idx])
    If FileExists($pdf) Then
        $g_aPDF[$idx] = $pdf
        ShellExecute($pdf)
        _Status("PDF recréé depuis HTML sélectionné :" & @CRLF & $pdf)
        Return True
    EndIf
    Return False
EndFunc


Func _SendOneResult($idx)
    ; Compatibilité : mail d'abord, puis EDOC pour ce BL.
    If _CreateMailForResult($idx) = "" Then Return False
    Return _UploadEdocForResult($idx)
EndFunc


Func _ResultsSetAllChecks($b)
    If $idResults = 0 Then Return False
    Local $count = _GUICtrlListView_GetItemCount($idResults)
    For $i = 0 To $count - 1
        _GUICtrlListView_SetItemChecked($idResults, $i, $b)
    Next
    Return True
EndFunc


Func _ResultsDetailedSummary()
    If UBound($g_aBL) = 0 Then Return "Aucun résultat BL."
    Local $s = "Résultats BL : " & UBound($g_aBL) & @CRLF & @CRLF
    For $i = 0 To UBound($g_aBL) - 1
        $s &= "#" & ($i + 1) & " - " & $g_aBL[$i] & @CRLF
        $s &= "Entreprise : " & $g_aCompany[$i] & @CRLF
        $s &= "SP : " & $g_aCarrier[$i] & @CRLF
        $s &= "Livraison : " & _ExtractDeliveryDateTime($g_aDelivery[$i]) & @CRLF
        $s &= "PDF : " & $g_aPDF[$i] & @CRLF & @CRLF
    Next
    Return $s
EndFunc


Func _SendAllResults()
    If UBound($g_aPDF) = 0 Then
        MsgBox(48, $APP_TITLE, "Aucun BL généré.")
        Return False
    EndIf
    For $i = 0 To UBound($g_aPDF) - 1
        _SendOneResult($i)
    Next
    _Status("EDOC + mails lancés pour tous les BL générés.")
    Return True
EndFunc


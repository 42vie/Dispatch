; ============================================================================
; CMR_UI.au3
; Boutons avec effet de survol utilises par l'ecran CMR historique.
; ----------------------------------------------------------------------------
; Deplace depuis StateService.au3 (bloc CMR_FULL_FUSION_MODULE_APPENDED) pour
; que le code CMR vive dans src/features/cmr/, comme le reste du projet.
; ============================================================================

; ============================================================
; MODERN UX HELPERS - HOVER SOFT / NON BLOQUANT
; ============================================================

Func _CreateHoverButton($sText, $x, $y, $w, $h, $normalColor = 0xF8FAFC, $hoverColor = 0xE2E8F0, $textColor = 0x111827)
    Local $idButton = GUICtrlCreateButton($sText, $x, $y, $w, $h)
    GUICtrlSetFont($idButton, 9, 700, 0, "Segoe UI")
    GUICtrlSetBkColor($idButton, $normalColor)
    GUICtrlSetColor($idButton, $textColor)
    _RegisterHoverButton($idButton, $normalColor, $hoverColor)
    Return $idButton
EndFunc


Func _RegisterHoverButton($idButton, $normalColor, $hoverColor)
    Local $n = UBound($g_aHoverIds)
    Local $nNormal = UBound($g_aHoverNormal)
    Local $nHover = UBound($g_aHoverHover)
    Local $nIsHover = UBound($g_aHoverIsHover)
    ReDim $g_aHoverIds[$n + 1]
    ReDim $g_aHoverNormal[$n + 1]
    ReDim $g_aHoverHover[$n + 1]
    ReDim $g_aHoverIsHover[$n + 1]
    $g_aHoverIds[$n] = $idButton
    $g_aHoverNormal[$n] = $normalColor
    $g_aHoverHover[$n] = $hoverColor
    $g_aHoverIsHover[$n] = False
EndFunc


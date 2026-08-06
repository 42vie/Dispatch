; ============================================================================
; CMR_EDOC.au3
; EDOC CMR.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

Func _CMR_EdocByIndex($idx)
    If $idx < 0 Or $idx >= UBound($g_aBL) Then Return False
    Return _UploadEdocForResult($idx)
EndFunc

; EDOC_MASTER_FUSION_V1 - EDOC Master Bot intégré au serveur Dispatch
Global $EDOC_APP_TITLE = "EDOC Master Bot V15.6"
Global $CFG_FILE = @ScriptDir & "\robot_v15_config.ini"
Global $TEMP_DIR = @TempDir & "\Temp_EDOC_Robot_V15"
Global $EDOC_WM_NOTIFY = 0x004E, $EDOC_NM_DBLCLK = -3
Global $C_BG = 0xF3F4F6, $C_CARD = 0xFFFFFF, $C_BORDER = 0xCBD5E1, $C_TEXT = 0x111827, $C_MUTED = 0x6B7280, $C_ACCENT = 0xDCFCE7, $C_BUTTON = 0xE5E7EB
Global $g_oOutlook = 0, $g_oNamespace = 0, $g_sHistorique = "|"
Global $g_aDynCtrls[30][6], $g_iDynCount = 0
Global $g_aGrouped[800][15], $g_iGroupedCount = 0
Global $g_idSelList = 0, $g_hSelList = 0

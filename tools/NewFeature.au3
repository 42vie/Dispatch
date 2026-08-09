#NoTrayIcon
Opt("MustDeclareVars", 0)

; ============================================================================
; NewFeature.au3
; Generateur de squelette pour un nouvel outil Dispatch : cree le fichier
; .au3 backend (CRUD sur ini), branche les routes HTTP dans HttpRouter.au3,
; ajoute l'include dans MainDispatch.au3, et ajoute un onglet de base dans
; Interface.html. Les 4 sorties sont un POINT DE DEPART -- personnalise-les
; librement ensuite (champs, logique, mise en forme...). Rien n'est fige :
; si un fichier/route/onglet existe deja, il n'est pas touche (le
; generateur est sans danger a relancer plusieurs fois).
;
; Usage :
;   NewFeature.exe <Nom>          (ex: NewFeature.exe STOCK)
;   NewFeature.exe                (saisie interactive du nom)
;
; Les 3 templates sont dans tools\templates\*.tpl -- modifie-les
; directement si tu veux changer ce que le generateur produit par defaut,
; sans toucher a ce script.
; ============================================================================

Global $g_sRoot = _FindProjectRoot()

_NewFeature_Main()

Func _NewFeature_Main()
    Local $sName = ""
    If $CmdLine[0] >= 1 Then $sName = $CmdLine[1]
    If $sName = "" Then $sName = InputBox("Nouvel outil Dispatch", "Nom du nouvel outil (ex: STOCK, LITIGE...) :", "", "", 400, 130)
    If @error Or StringStripWS($sName, 3) = "" Then
        ConsoleWrite("Annule : aucun nom fourni." & @CRLF)
        Exit 1
    EndIf

    Local $sNameUpper = StringUpper(_CleanName($sName))
    Local $sNameLower = StringLower($sNameUpper)
    If $sNameUpper = "" Then
        MsgBox(16, "NewFeature", "Nom invalide (utilise uniquement des lettres/chiffres).")
        Exit 1
    EndIf

    Local $sTabLabel = "🧩 " & $sNameUpper

    Local $sReport = ""
    $sReport &= _GenBackend($sNameUpper, $sNameLower) & @CRLF
    $sReport &= _PatchMainInclude($sNameUpper, $sNameLower) & @CRLF
    $sReport &= _PatchRouter($sNameUpper, $sNameLower) & @CRLF
    $sReport &= _PatchInterface($sNameUpper, $sNameLower, $sTabLabel) & @CRLF

    MsgBox(64, "NewFeature", "Outil '" & $sNameUpper & "' :" & @CRLF & @CRLF & $sReport & @CRLF & _
            "Redemarre Dispatch pour voir le nouvel onglet. Verifie les fichiers avec DispatchDoctor.exe avant de committer.")
    ConsoleWrite($sReport)
EndFunc

Func _CleanName($s)
    Return StringRegExpReplace($s, "[^A-Za-z0-9]", "")
EndFunc

; Remonte depuis le dossier du script (tools\) jusqu'a trouver
; MainDispatch.au3, pour marcher quel que soit l'endroit d'ou l'exe est
; lance.
Func _FindProjectRoot()
    Local $sDir = @ScriptDir
    For $i = 1 To 4
        If FileExists($sDir & "\MainDispatch.au3") Then Return $sDir
        $sDir = _PathParent($sDir)
    Next
    Return @ScriptDir & "\.."
EndFunc

Func _PathParent($sDir)
    Local $iPos = StringInStr($sDir, "\", 0, -1)
    If $iPos <= 3 Then Return $sDir
    Return StringLeft($sDir, $iPos - 1)
EndFunc

Func _ReadTemplate($sFile)
    Local $sPath = @ScriptDir & "\templates\" & $sFile
    Local $sContent = FileRead($sPath)
    If @error Then
        MsgBox(16, "NewFeature", "Template introuvable : " & $sPath & @CRLF & _
                "Le dossier tools\templates\ doit etre a cote de NewFeature.au3.")
        Exit 1
    EndIf
    Return $sContent
EndFunc

Func _WriteFile($sPath, $sContent)
    ; 2 = ecrase/cree, 128 = UTF8 sans BOM (coherent avec le reste du projet)
    Local $hFile = FileOpen($sPath, 2 + 128)
    If $hFile = -1 Then Return SetError(1, 0, False)
    FileWrite($hFile, $sContent)
    FileClose($hFile)
    Return True
EndFunc

Func _ApplyPlaceholders($sTpl, $sNameUpper, $sNameLower, $sTabLabel = "")
    Local $s = $sTpl
    $s = StringReplace($s, "{{NAME_UPPER}}", $sNameUpper)
    $s = StringReplace($s, "{{NAME_LOWER}}", $sNameLower)
    $s = StringReplace($s, "{{TAB_LABEL}}", $sTabLabel)
    $s = StringReplace($s, "{{DATE}}", @YEAR & "-" & @MON & "-" & @MDAY)
    Return $s
EndFunc

Func _GenBackend($sNameUpper, $sNameLower)
    Local $sDir = $g_sRoot & "\src\features\" & $sNameLower
    Local $sPath = $sDir & "\" & $sNameUpper & "_Tool.au3"
    If FileExists($sPath) Then Return "- " & $sPath & " : deja existant, non touche."
    DirCreate($sDir)
    Local $sTpl = _ReadTemplate("feature_backend.au3.tpl")
    Local $sOut = _ApplyPlaceholders($sTpl, $sNameUpper, $sNameLower)
    If Not _WriteFile($sPath, $sOut) Then Return "- ECHEC ecriture : " & $sPath
    Return "- Cree : src\features\" & $sNameLower & "\" & $sNameUpper & "_Tool.au3"
EndFunc

Func _PatchMainInclude($sNameUpper, $sNameLower)
    Local $sPath = $g_sRoot & "\MainDispatch.au3"
    Local $sContent = FileRead($sPath)
    If @error Then Return "- MainDispatch.au3 introuvable a la racine attendue (" & $g_sRoot & ")."
    Local $sIncludeLine = '#include "src\features\' & $sNameLower & '\' & $sNameUpper & '_Tool.au3"'
    If StringInStr($sContent, $sIncludeLine) Then Return "- MainDispatch.au3 : include deja present."
    Local $sAnchor = "; ========== API =========="
    If Not StringInStr($sContent, $sAnchor) Then Return "- MainDispatch.au3 : ancre '" & $sAnchor & "' introuvable, include NON ajoute (a faire a la main)."
    Local $sNew = StringReplace($sContent, $sAnchor, $sIncludeLine & @CRLF & @CRLF & $sAnchor, 1)
    If Not _WriteFile($sPath, $sNew) Then Return "- ECHEC ecriture : " & $sPath
    Return "- MainDispatch.au3 : include ajoute."
EndFunc

Func _PatchRouter($sNameUpper, $sNameLower)
    Local $sPath = $g_sRoot & "\src\core\HttpRouter.au3"
    Local $sContent = FileRead($sPath)
    If @error Then Return "- HttpRouter.au3 introuvable."
    Local $sCheckAction = 'Case "' & $sNameUpper & '_LIST"'
    If StringInStr($sContent, $sCheckAction) Then Return "- HttpRouter.au3 : routes deja presentes."
    Local $sTpl = _ReadTemplate("router_case.au3.tpl")
    Local $sCaseBlock = _ApplyPlaceholders($sTpl, $sNameUpper, $sNameLower)
    Local $sAnchor = @CRLF & "        EndSwitch"
    If Not StringInStr($sContent, $sAnchor) Then Return "- HttpRouter.au3 : ancre EndSwitch introuvable, routes NON ajoutees (a faire a la main)."
    Local $sNew = StringReplace($sContent, $sAnchor, @CRLF & $sCaseBlock & $sAnchor, 1)
    If Not _WriteFile($sPath, $sNew) Then Return "- ECHEC ecriture : " & $sPath
    Return "- HttpRouter.au3 : routes " & $sNameUpper & "_LIST/ADD/DELETE ajoutees."
EndFunc

Func _PatchInterface($sNameUpper, $sNameLower, $sTabLabel)
    Local $sPath = $g_sRoot & "\Interface.html"
    Local $sContent = FileRead($sPath)
    If @error Then Return "- Interface.html introuvable."
    Local $sCheckId = 'id="' & $sNameLower & '-ui-v1-style"'
    If StringInStr($sContent, $sCheckId) Then Return "- Interface.html : onglet deja present."
    Local $sTpl = _ReadTemplate("interface_tab.html.tpl")
    Local $sBlock = _ApplyPlaceholders($sTpl, $sNameUpper, $sNameLower, $sTabLabel)
    Local $sAnchor = "</body>"
    If Not StringInStr($sContent, $sAnchor) Then Return "- Interface.html : ancre </body> introuvable, onglet NON ajoute (a faire a la main)."
    Local $sNew = StringReplace($sContent, $sAnchor, $sBlock & @CRLF & @CRLF & $sAnchor, 1)
    If Not _WriteFile($sPath, $sNew) Then Return "- ECHEC ecriture : " & $sPath
    Return "- Interface.html : onglet '" & $sTabLabel & "' ajoute."
EndFunc

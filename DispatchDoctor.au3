#NoTrayIcon
; ============================================================================
; DispatchDoctor.au3
;
; Outil permanent de diagnostic et de maintenance pour le projet Dispatch.
;
; Ce fichier est volontairement monolithique (fichier unique) et generique :
; il n'a AUCUNE connaissance codee en dur des fonctions, fichiers ou domaines
; actuels de Dispatch. Tout est decouvert dynamiquement en suivant les
; #include depuis MainDispatch.au3 (ou le script cible passe en argument).
;
; Ajouter une fonction, un fichier, un service, une route HTTP ou un domaine
; dans Dispatch NE NECESSITE JAMAIS de modifier ce fichier : DispatchDoctor
; continue de fonctionner tel quel. Seul l'ajout d'une regle de classement
; par domaine optionnelle peut justifier de toucher la table
; $g_aDomainRules (voir plus bas, region "Regles par domaine").
;
; Usage :
;   DispatchDoctor.au3
;   DispatchDoctor.au3 MainDispatch.au3
;   DispatchDoctor.au3 --domain CMR
;   DispatchDoctor.au3 --mode full|compile|analyze|list|selftest
;   DispatchDoctor.au3 MainDispatch.au3 --domain COMAT --mode full
;
; Sans argument, DispatchDoctor analyse automatiquement :
;   @ScriptDir & "\MainDispatch.au3"
;
; Par defaut, DispatchDoctor NE MODIFIE JAMAIS les fichiers du projet
; Dispatch. Il ne fait qu'analyser et ecrire des rapports dans le dossier
; "reports" a cote de lui.
;
; Note sur l'analyse heuristique des variables (section "Analyse des
; variables") : contrairement a la sortie du compilateur AutoIt (qui est
; la verite absolue), l'analyse statique de DispatchDoctor est un filet de
; securite "best effort". Elle est volontairement permissive (elle prefere
; rater un vrai probleme plutot qu'inventer un faux positif) et annonce
; toujours son niveau de confiance ("high" seulement quand confirme par le
; compilateur reel, "moyenne"/"faible" pour ses propres deductions).
; ============================================================================

Opt("MustDeclareVars", 1)
Opt("TrayIconHide", 1)

#include <StringConstants.au3>

; ============================================================================
#region Initialisation - Constantes globales
; ============================================================================

Global Const $DD_VERSION = "1.0.0"
Global Const $DD_TOOLNAME = "DispatchDoctor"
Global Const $DD_DEFAULT_MAIN = "MainDispatch.au3"

Global Const $DD_STDOUT_CHILD = 2
Global Const $DD_STDERR_CHILD = 4

Global Const $DD_SEV_ERROR = "ERROR"
Global Const $DD_SEV_WARNING = "WARNING"
Global Const $DD_SEV_INFO = "INFO"
Global Const $DD_SEV_OK = "OK"

Global Const $DD_CODE_MISSING_INCLUDE = "MISSING_INCLUDE"
Global Const $DD_CODE_CIRCULAR_INCLUDE = "CIRCULAR_INCLUDE"
Global Const $DD_CODE_DUPLICATE_INCLUDE = "DUPLICATE_INCLUDE"
Global Const $DD_CODE_ORPHAN_FILE = "ORPHAN_FILE"
Global Const $DD_CODE_DUPLICATE_FUNCTION = "DUPLICATE_FUNCTION"
Global Const $DD_CODE_SIMILAR_FUNCTION = "SIMILAR_FUNCTION_NAME"
Global Const $DD_CODE_MISSING_ENDFUNC = "MISSING_ENDFUNC"
Global Const $DD_CODE_UNBALANCED_BLOCK = "UNBALANCED_BLOCK"
Global Const $DD_CODE_UNDECLARED_VARIABLE = "UNDECLARED_VARIABLE"
Global Const $DD_CODE_UNKNOWN_FUNCTION = "UNKNOWN_FUNCTION"
Global Const $DD_CODE_BAD_PARAM_COUNT = "BAD_PARAM_COUNT"
Global Const $DD_CODE_COMPILER_ERROR = "COMPILER_ERROR"
Global Const $DD_CODE_AUTOIT_NOT_FOUND = "AUTOIT_NOT_FOUND"
Global Const $DD_CODE_REDIM_UNSAFE = "REDIM_UNSAFE_FIRST_REFERENCE"
Global Const $DD_CODE_OK = "OK"

; Colonnes table des fichiers
Global Const $FCOL_PATH = 0
Global Const $FCOL_RELPATH = 1
Global Const $FCOL_PARENT = 2
Global Const $FCOL_LINE = 3
Global Const $FCOL_DEPTH = 4
Global Const $FCOL_DOMAIN = 5
Global Const $FCOL_COLS = 6

; Colonnes table des fonctions
Global Const $UCOL_NAME = 0
Global Const $UCOL_FILE = 1
Global Const $UCOL_LINE = 2
Global Const $UCOL_ENDLINE = 3
Global Const $UCOL_PARAMS = 4
Global Const $UCOL_PARAMCOUNT = 5
Global Const $UCOL_DOMAIN = 6
Global Const $UCOL_COLS = 7

; Colonnes table des diagnostics
Global Const $DCOL_SEVERITY = 0
Global Const $DCOL_CODE = 1
Global Const $DCOL_DOMAIN = 2
Global Const $DCOL_FILE = 3
Global Const $DCOL_LINE = 4
Global Const $DCOL_FUNCTION = 5
Global Const $DCOL_SYMBOL = 6
Global Const $DCOL_MESSAGE = 7
Global Const $DCOL_SUGGESTION = 8
Global Const $DCOL_CONFIDENCE = 9
Global Const $DCOL_COLS = 10

#endregion

; ============================================================================
#region Etat global de l'execution
; ============================================================================

Global $g_sProjectRoot = ""
Global $g_sTargetScript = ""
Global $g_sMode = "full"
Global $g_sDomainFilter = ""
Global $g_sAutoItPath = ""

Global $g_aFiles[8][$FCOL_COLS]
Global $g_iFileCount = 0

Global $g_aFunctions[8][$UCOL_COLS]
Global $g_iFuncCount = 0

Global $g_aDiagnostics[8][$DCOL_COLS]
Global $g_iDiagCount = 0

Global $g_iCountError = 0
Global $g_iCountWarning = 0
Global $g_iCountInfo = 0
Global $g_iCountOk = 0

Global $g_iIncludeEdgeCount = 0
Global $g_iVariableCount = 0

Global $oDD_FileTextCache = ObjCreate("Scripting.Dictionary")
Global $oDD_VisitedFiles = ObjCreate("Scripting.Dictionary")
Global $oDD_IncludeCount = ObjCreate("Scripting.Dictionary")
Global $oDD_IncludedBy = ObjCreate("Scripting.Dictionary")
Global $oDD_Ancestors = ObjCreate("Scripting.Dictionary")
Global $oDD_GlobalVars = ObjCreate("Scripting.Dictionary")
Global $oDD_AllProjectFiles = ObjCreate("Scripting.Dictionary")

; Etat de l'interface graphique (fenetre de progression). $g_bGuiActive
; conditionne tous les appels _DD_Gui*, qui sont sans effet tant qu'elle
; est a False (permet de garder le mode "selftest" 100% console).
Global $g_bGuiActive = False
Global $g_hGui = 0
Global $g_idLblStatus = 0
Global $g_idLblStats = 0
Global $g_idEditLog = 0
Global $g_idBtnClose = 0
Global $g_idBtnOpenHtml = 0
Global $g_sGuiLogBuffer = ""
Global $g_sHtmlReportPath = ""
Global $g_idComboDomain = 0
Global $g_idBtnFilterDomain = 0
Global $g_idInputSearch = 0
Global $g_idBtnSearch = 0
Global $g_idBtnOpenFile = 0
Global $g_idListFunc = 0
Global $g_idComboTestAction = 0
Global $g_idInputParam1 = 0
Global $g_idInputParam2 = 0
Global $g_idBtnRunTest = 0
Global $g_bTcpStarted = False
Global $g_sLastHttpError = ""

; Port HTTP reel du serveur Dispatch (voir MainDispatch.au3 : $giPort = 9500).
Global Const $DD_DISPATCH_PORT = 9500
Global $g_bActionTestRunning = False
Global $g_iActionTestPid = 0

; Registre des actions "testables en reel" via l'API HTTP de Dispatch
; (POST /api/action, voir HttpRouter.au3). Ajouter une action plus tard =
; ajouter une ligne ici, rien d'autre a toucher dans le moteur de test.
; Colonnes : {libelle affiche, nom d'action HTTP, champ JSON principal,
; champ JSON secondaire ("" si aucun), avertissement specifique}.
Global $g_aTestableActions[3][5] = [ _
    ["CMR - generer un dossier (CMR_GENERATE)", "CMR_GENERATE", "data", "", _
        "Genere reellement les BL/PDF/HTML pour ce dossier (pas de mail ni d'upload EDOC automatique)."], _
    ["COMAT - traiter un dossier seul (COMAT_SOLO)", "COMAT_SOLO", "file", "", _
        "Affiche une confirmation (Oui/Non) dans la fenetre Dispatch : il faudra la valider a la main, sinon le test restera bloque jusqu'au timeout."], _
    ["ETMS - commande sur un dossier (ETMS_CMD)", "ETMS_CMD", "file", "cmd", _
        "Necessite que la fenetre E.TMS (TfmBrowser) soit deja ouverte sur ce poste, sinon l'action echoue proprement sans rien tester."] _
]

; Table de classification par domaine : {sous-chaine (chemin ou nom, en
; minuscules), domaine}. Premiere ligne qui correspond gagne. Ajouter un
; nouveau domaine plus tard = ajouter une ligne ici, rien d'autre.
Global $g_aDomainRules[14][2] = [ _
    ["\features\cmr\", "CMR"], _
    ["cmr_", "CMR"], _
    ["\features\comat\", "COMAT"], _
    ["comat_", "COMAT"], _
    ["\features\etms\", "ETMS"], _
    ["etms_", "ETMS"], _
    ["\features\mail\", "MAIL"], _
    ["mail_", "MAIL"], _
    ["\api\", "API"], _
    ["api", "API"], _
    ["http", "HTTP"], _
    ["\core\fileutils", "FILES"], _
    ["\core\backuputils", "FILES"], _
    ["file", "FILES"] _
]

; Prefixes/noms de constantes standard AutoIt et Windows API (issues des
; #include <...> systeme comme GUIConstantsEx.au3, WindowsConstants.au3,
; EditConstants.au3, ListViewConstants.au3, MsgBoxConstants.au3, ...) que
; l'outil ne lit jamais (hors perimetre du projet Dispatch). Sans cette
; liste, l'analyse heuristique des variables les signale a tort comme non
; declarees des qu'un fichier les utilise sans les redefinir localement.
; Ajouter un prefixe ici si un nouveau fichier standard AutoIt est utilise
; par le projet et genere du bruit similaire.
Global $g_aKnownConstantPrefixes[20] = ["GUI_", "WS_", "ES_", "SS_", "BS_", "CBS_", "LVS_", "LVM_", "LVIF_", _
    "LVNI_", "MB_", "TVS_", "TVIF_", "DTN_", "UDS_", "ICC_", "PBS_", "TCS_", "STM_", "GUI_RUNDEFMSG"]
Global $g_aKnownConstantExact[8] = ["idyes", "idno", "idok", "idcancel", "idabort", "idretry", "idignore", "idclose"]

#endregion

_DD_Main()

; ============================================================================
#region Initialisation - Point d'entree
; ============================================================================

Func _DD_Main()
    _DD_ParseArgs()
    _DD_ResolveProjectRoot()

    ConsoleWrite("=== " & $DD_TOOLNAME & " v" & $DD_VERSION & " ===" & @CRLF)
    ConsoleWrite("Racine projet : " & $g_sProjectRoot & @CRLF)
    ConsoleWrite("Mode          : " & $g_sMode & @CRLF)
    If $g_sDomainFilter <> "" Then ConsoleWrite("Domaine       : " & $g_sDomainFilter & @CRLF)

    If $g_sMode = "selftest" Then
        _DD_RunSelfTests()
        Exit 0
    EndIf

    ; Fenetre de progression visible pour tous les modes sauf "selftest"
    ; (qui doit rester rapide, automatisable et 100% console).
    _DD_GuiInit()
    _DD_GuiLog("Racine projet : " & $g_sProjectRoot)
    _DD_GuiLog("Mode          : " & $g_sMode)
    If $g_sDomainFilter <> "" Then _DD_GuiLog("Domaine       : " & $g_sDomainFilter)

    If Not _DD_ResolveTargetScript() Then
        ConsoleWrite("[ERROR] Script principal introuvable : " & $g_sTargetScript & @CRLF)
        _DD_GuiSetStatus("ERREUR : script principal introuvable.")
        _DD_GuiLog("[ERROR] Script principal introuvable : " & $g_sTargetScript)
        _DD_GuiWaitClose()
        Exit 2
    EndIf
    ConsoleWrite("Script cible  : " & $g_sTargetScript & @CRLF & @CRLF)
    _DD_GuiLog("Script cible  : " & $g_sTargetScript)

    Switch $g_sMode
        Case "list"
            _DD_GuiSetStatus("Analyse des includes...")
            _DD_BuildIncludeGraph($g_sTargetScript)
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Indexation des fonctions...")
            _DD_IndexFunctions()
            _DD_GuiRefreshStats()
            _DD_GuiPopulateBrowser()
            _DD_ModeList()

        Case "compile"
            _DD_GuiSetStatus("Detection d'AutoIt3.exe...")
            _DD_DetectAutoIt()
            _DD_GuiSetStatus("Compilation (AutoIt3.exe /ErrorStdOut)...")
            _DD_RunCompilerAndAnalyze()
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Generation des rapports...")
            _DD_GenerateReports()

        Case "analyze"
            _DD_GuiSetStatus("Analyse des includes...")
            _DD_BuildIncludeGraph($g_sTargetScript)
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Recherche des fichiers orphelins...")
            _DD_ScanOrphanFiles()
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Indexation des fonctions...")
            _DD_IndexFunctions()
            _DD_GuiRefreshStats()
            _DD_GuiPopulateBrowser()
            _DD_GuiSetStatus("Indexation des variables...")
            _DD_IndexGlobalVariables()
            _DD_GuiSetStatus("Analyse heuristique des variables...")
            _DD_AnalyzeAllVariables()
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Generation des rapports...")
            _DD_GenerateReports()

        Case Else ; "full"
            _DD_GuiSetStatus("Analyse des includes...")
            _DD_BuildIncludeGraph($g_sTargetScript)
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Recherche des fichiers orphelins...")
            _DD_ScanOrphanFiles()
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Indexation des fonctions...")
            _DD_IndexFunctions()
            _DD_GuiRefreshStats()
            _DD_GuiPopulateBrowser()
            _DD_GuiSetStatus("Indexation des variables...")
            _DD_IndexGlobalVariables()
            _DD_GuiSetStatus("Analyse heuristique des variables...")
            _DD_AnalyzeAllVariables()
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Detection d'AutoIt3.exe...")
            _DD_DetectAutoIt()
            _DD_GuiSetStatus("Compilation (AutoIt3.exe /ErrorStdOut)...")
            _DD_RunCompilerAndAnalyze()
            _DD_GuiRefreshStats()
            _DD_GuiSetStatus("Generation des rapports...")
            _DD_GenerateReports()
    EndSwitch

    ConsoleWrite(@CRLF & "Termine. Erreurs=" & $g_iCountError & " Avertissements=" & $g_iCountWarning & _
        " Infos=" & $g_iCountInfo & @CRLF)
    _DD_GuiLog(@CRLF & "Termine. Erreurs=" & $g_iCountError & " Avertissements=" & $g_iCountWarning & " Infos=" & $g_iCountInfo)
    _DD_GuiRefreshStats()

    ; Bloque ici tant que la fenetre est ouverte (l'utilisateur peut lire
    ; le journal, ouvrir le rapport HTML, puis fermer). Sans effet si la
    ; GUI n'a pas pu s'initialiser.
    _DD_GuiWaitClose()

    If $g_iCountError > 0 Then
        Exit 1
    Else
        Exit 0
    EndIf
EndFunc   ;==>_DD_Main

#endregion

; ============================================================================
#region Arguments
; ============================================================================

Func _DD_ParseArgs()
    Local $sFirstFree = ""
    Local $i = 1

    While $i <= $CmdLine[0]
        Local $sArg = $CmdLine[$i]

        If StringLower($sArg) = "--domain" Then
            If $i + 1 <= $CmdLine[0] Then
                $g_sDomainFilter = StringUpper($CmdLine[$i + 1])
                $i += 1
            EndIf
        ElseIf StringLower($sArg) = "--mode" Then
            If $i + 1 <= $CmdLine[0] Then
                $g_sMode = StringLower($CmdLine[$i + 1])
                $i += 1
            EndIf
        ElseIf StringLower($sArg) = "--selftest" Then
            $g_sMode = "selftest"
        ElseIf StringLeft($sArg, 2) = "--" Then
            ; Option inconnue : ignoree pour rester tolerant aux futures
            ; extensions sans casser l'outil.
        Else
            If $sFirstFree = "" Then $sFirstFree = $sArg
        EndIf

        $i += 1
    WEnd

    $g_sTargetScript = $sFirstFree
EndFunc   ;==>_DD_ParseArgs

#endregion

; ============================================================================
#region Recherche du projet
; ============================================================================

Func _DD_ResolveProjectRoot()
    $g_sProjectRoot = @ScriptDir
EndFunc   ;==>_DD_ResolveProjectRoot

Func _DD_ResolveTargetScript()
    If $g_sTargetScript = "" Then
        $g_sTargetScript = $g_sProjectRoot & "\" & $DD_DEFAULT_MAIN
    ElseIf Not _DD_IsAbsolutePath($g_sTargetScript) Then
        $g_sTargetScript = $g_sProjectRoot & "\" & $g_sTargetScript
    EndIf

    $g_sTargetScript = _DD_NormalizePath($g_sTargetScript)

    If Not FileExists($g_sTargetScript) Then Return False
    Return True
EndFunc   ;==>_DD_ResolveTargetScript

Func _DD_IsAbsolutePath($sPath)
    If StringRegExp($sPath, "(?i)^[A-Z]:\\") Then Return True
    If StringLeft($sPath, 2) = "\\" Then Return True
    Return False
EndFunc   ;==>_DD_IsAbsolutePath

#endregion

; ============================================================================
#region Detection AutoIt
; ============================================================================

Func _DD_DetectAutoIt()
    Local $sCandidate = @ScriptDir & "\AutoIt3.exe"
    If FileExists($sCandidate) Then
        $g_sAutoItPath = $sCandidate
        Return
    EndIf

    $sCandidate = @ProgramFilesDir & "\AutoIt3\AutoIt3.exe"
    If FileExists($sCandidate) Then
        $g_sAutoItPath = $sCandidate
        Return
    EndIf

    $sCandidate = @ProgramFilesDir & " (x86)\AutoIt3\AutoIt3.exe"
    If FileExists($sCandidate) Then
        $g_sAutoItPath = $sCandidate
        Return
    EndIf

    Local $sPathHit = _DD_SearchInSystemPath("AutoIt3.exe")
    If $sPathHit <> "" Then
        $g_sAutoItPath = $sPathHit
        Return
    EndIf

    $g_sAutoItPath = ""
EndFunc   ;==>_DD_DetectAutoIt

Func _DD_SearchInSystemPath($sExeName)
    Local $sPathEnv = EnvGet("PATH")
    If $sPathEnv = "" Then Return ""

    Local $aDirs = StringSplit($sPathEnv, ";", $STR_NOCOUNT)
    Local $iN = UBound($aDirs)
    Local $i = 0
    For $i = 0 To $iN - 1
        If $aDirs[$i] = "" Then ContinueLoop
        Local $sCandidate = StringRegExpReplace($aDirs[$i], "\\+$", "") & "\" & $sExeName
        If FileExists($sCandidate) Then Return $sCandidate
    Next

    Return ""
EndFunc   ;==>_DD_SearchInSystemPath

#endregion

; ============================================================================
#region Resolution des includes
; ============================================================================

; Construit l'arbre des dependances par parcours en largeur a partir du
; script racine. Chaque fichier resolu n'est "reclame" qu'une seule fois
; (dictionnaire $oDD_VisitedFiles), donc la file ne peut jamais grossir
; indefiniment : aucune boucle infinie possible, meme en cas de cycle reel.
Func _DD_BuildIncludeGraph($sRootFile)
    ConsoleWrite("[*] Analyse des includes depuis " & _DD_RelativePath($sRootFile, $g_sProjectRoot) & " ..." & @CRLF)

    Local $aQPath[64]
    Local $aQParent[64]
    Local $aQLine[64]
    Local $aQDepth[64]
    Local $iQCount = 1
    Local $iQRead = 0

    $aQPath[0] = $sRootFile
    $aQParent[0] = ""
    $aQLine[0] = 0
    $aQDepth[0] = 0

    $oDD_VisitedFiles.Add(StringLower($sRootFile), $sRootFile)
    $oDD_Ancestors.Add(StringLower($sRootFile), "")
    $oDD_IncludeCount.Add(StringLower($sRootFile), 1)
    $oDD_IncludedBy.Add(StringLower($sRootFile), "")

    While $iQRead < $iQCount
        Local $sPath = $aQPath[$iQRead]
        Local $sParent = $aQParent[$iQRead]
        Local $iInclLine = $aQLine[$iQRead]
        Local $iDepth = $aQDepth[$iQRead]
        $iQRead += 1

        Local $sKey = StringLower($sPath)
        Local $sAncestors = $oDD_Ancestors.Item($sKey)
        Local $sRel = _DD_RelativePath($sPath, $g_sProjectRoot)
        Local $sDomain = _DD_GuessDomainFromPath($sPath)
        _DD_AddFileRow($sPath, $sRel, $sParent, $iInclLine, $iDepth, $sDomain)

        If Not FileExists($sPath) Then ContinueLoop

        Local $aIncludes = _DD_ExtractIncludes($sPath)
        Local $iIncN = UBound($aIncludes, 1)
        Local $iInc = 0
        For $iInc = 0 To $iIncN - 1
            Local $bAngle = ($aIncludes[$iInc][0] = "1")
            Local $sRequested = $aIncludes[$iInc][1]
            Local $iLine2 = Number($aIncludes[$iInc][2])

            Local $sResolved = _DD_ResolveIncludePath($sRequested, $bAngle, _DD_ParentDir($sPath))

            If $sResolved = "" Then
                If Not $bAngle Then
                    Local $sDomErr = _DD_GuessDomainFromPath($sPath)
                    _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_MISSING_INCLUDE, $sDomErr, $sPath, $iLine2, "", $sRequested, _
                        "Le fichier inclus '" & $sRequested & "' est introuvable.", _
                        "Verifier le chemin relatif ou le nom du fichier (relatif a " & _DD_RelativePath($sPath, $g_sProjectRoot) & " ou a la racine du projet).", "high")
                EndIf
                ContinueLoop
            EndIf

            Local $sKey2 = StringLower($sResolved)
            Local $sChain = "|" & StringLower($sAncestors) & StringLower($sPath) & "|"

            If StringInStr($sChain, "|" & $sKey2 & "|") > 0 Then
                Local $sDomCirc = _DD_GuessDomainFromPath($sPath)
                _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_CIRCULAR_INCLUDE, $sDomCirc, $sPath, $iLine2, "", $sRequested, _
                    "Inclusion circulaire detectee : '" & _DD_RelativePath($sResolved, $g_sProjectRoot) & _
                    "' inclut (directement ou indirectement) '" & _DD_RelativePath($sPath, $g_sProjectRoot) & _
                    "' qui tente de l'inclure a nouveau.", _
                    "Ajouter #include-once en tete des deux fichiers et casser la dependance croisee (deplacer le code partage dans un troisieme fichier).", "high")
                ContinueLoop
            EndIf

            If $oDD_VisitedFiles.Exists($sKey2) Then
                $oDD_IncludeCount.Item($sKey2) = $oDD_IncludeCount.Item($sKey2) + 1
                $oDD_IncludedBy.Item($sKey2) = $oDD_IncludedBy.Item($sKey2) & ";" & $sPath & "(" & $iLine2 & ")"
                Local $sDomDup = _DD_GuessDomainFromPath($sResolved)
                _DD_AddDiagnostic($DD_SEV_INFO, $DD_CODE_DUPLICATE_INCLUDE, $sDomDup, $sPath, $iLine2, "", $sRequested, _
                    "Le fichier '" & _DD_RelativePath($sResolved, $g_sProjectRoot) & "' est inclus depuis plusieurs endroits (inclusion multiple).", _
                    "Normalement neutralise par #include-once ; verifier que ce mot-cle est present en tete de ce fichier.", "high")
                ContinueLoop
            EndIf

            $oDD_VisitedFiles.Add($sKey2, $sResolved)
            $oDD_IncludeCount.Add($sKey2, 1)
            $oDD_IncludedBy.Add($sKey2, $sPath & "(" & $iLine2 & ")")
            $oDD_Ancestors.Add($sKey2, StringLower($sAncestors) & StringLower($sPath) & "|")

            If $iQCount >= UBound($aQPath) Then
                ReDim $aQPath[UBound($aQPath) * 2]
                ReDim $aQParent[UBound($aQParent) * 2]
                ReDim $aQLine[UBound($aQLine) * 2]
                ReDim $aQDepth[UBound($aQDepth) * 2]
            EndIf
            $aQPath[$iQCount] = $sResolved
            $aQParent[$iQCount] = $sPath
            $aQLine[$iQCount] = $iLine2
            $aQDepth[$iQCount] = $iDepth + 1
            $iQCount += 1

            $g_iIncludeEdgeCount += 1
        Next
    WEnd

    ConsoleWrite("[*] " & $g_iFileCount & " fichier(s) analyse(s), " & $g_iIncludeEdgeCount & " include(s) resolu(s)." & @CRLF)
EndFunc   ;==>_DD_BuildIncludeGraph

; Extrait les directives #include "..." et #include <...> d'un fichier.
; Retourne un tableau [n][3] : {"1"/"0" (chevrons), chemin demande, ligne}.
; #include-once n'est jamais confondu avec ces directives car il n'est pas
; suivi d'un guillemet/chevron.
Func _DD_ExtractIncludes($sPath)
    Local $aLines = _DD_GetFileLines($sPath)
    Local $iN = UBound($aLines)
    Local $aResult[8][3]
    Local $iCount = 0
    Local $i = 0

    For $i = 0 To $iN - 1
        Local $aMatch = StringRegExp($aLines[$i], '(?i)^\s*#include\s+(["<])([^">]+)[">]', 1)
        If Not @error Then
            _DD_GrowTable($aResult, $iCount)
            If $aMatch[0] = "<" Then
                $aResult[$iCount][0] = "1"
            Else
                $aResult[$iCount][0] = "0"
            EndIf
            $aResult[$iCount][1] = $aMatch[1]
            $aResult[$iCount][2] = String($i + 1)
            $iCount += 1
        EndIf
    Next

    ReDim $aResult[$iCount][3]
    Return $aResult
EndFunc   ;==>_DD_ExtractIncludes

; Resout un chemin d'include en essayant, dans l'ordre :
;   1) relatif au fichier qui contient le #include (comportement AutoIt
;      standard pour #include "...") ;
;   2) relatif a la racine du projet (secours, utile aussi quand un include
;      systeme <...> est en realite fourni localement) ;
;   3) chemin absolu direct.
; Retourne "" si rien n'est trouve (les includes systeme <...> introuvables
; localement ne sont PAS des erreurs : ce sont des bibliotheques standard
; AutoIt, resolues par l'installation AutoIt elle-meme, hors du perimetre
; du projet Dispatch).
Func _DD_ResolveIncludePath($sRequested, $bAngle, $sParentDir)
    Local $sCandidate = ""

    If Not $bAngle Then
        $sCandidate = _DD_NormalizePath($sParentDir & "\" & $sRequested)
        If FileExists($sCandidate) Then Return $sCandidate
    EndIf

    $sCandidate = _DD_NormalizePath($g_sProjectRoot & "\" & $sRequested)
    If FileExists($sCandidate) Then Return $sCandidate

    If _DD_IsAbsolutePath($sRequested) And FileExists($sRequested) Then Return $sRequested

    Return ""
EndFunc   ;==>_DD_ResolveIncludePath

; Parcourt tous les .au3 reellement presents sur le disque (pas seulement
; ceux atteints depuis le script cible) et signale ceux qui ne sont inclus,
; ni directement ni indirectement, par aucun #include depuis le point
; d'entree analyse. Ce sont des candidats "code mort" ou des fichiers
; oublies dans la chaine d'inclusion.
Func _DD_ScanOrphanFiles()
    $oDD_AllProjectFiles = ObjCreate("Scripting.Dictionary")
    _DD_WalkAu3Files($g_sProjectRoot)

    Local $aKeys = $oDD_AllProjectFiles.Keys()
    Local $iOrphanCount = 0
    Local $i = 0
    Local $sSelfKey = StringLower(@ScriptFullPath)

    For $i = 0 To UBound($aKeys) - 1
        Local $sKey = $aKeys[$i]
        If $sKey = $sSelfKey Then ContinueLoop ; DispatchDoctor.au3 lui-meme n'est jamais cense etre inclus
        If Not $oDD_VisitedFiles.Exists($sKey) Then
            Local $sPath = $oDD_AllProjectFiles.Item($sKey)
            Local $sDomain = _DD_GuessDomainFromPath($sPath)
            Local $iFuncInFile = _DD_CountFuncDeclarations($sPath)
            _DD_AddDiagnostic($DD_SEV_WARNING, $DD_CODE_ORPHAN_FILE, $sDomain, $sPath, 0, "", "", _
                "Ce fichier .au3 existe sur le disque mais n'est inclus (directement ou indirectement) par aucun #include depuis " & _
                _DD_RelativePath($g_sTargetScript, $g_sProjectRoot) & ". " & $iFuncInFile & " fonction(s) qu'il contient sont donc invisibles pour Dispatch.", _
                "Verifier s'il doit etre ajoute a la chaine d'#include, ou s'il s'agit de code mort/legacy a archiver ou supprimer.", "high")
            $iOrphanCount += 1
        EndIf
    Next

    If $iOrphanCount > 0 Then
        ConsoleWrite("[*] " & $iOrphanCount & " fichier(s) .au3 orphelin(s) (non inclus) detecte(s)." & @CRLF)
    EndIf
EndFunc   ;==>_DD_ScanOrphanFiles

Func _DD_WalkAu3Files($sDir)
    Local $hSearch = FileFindFirstFile($sDir & "\*")
    If $hSearch = -1 Then Return

    Local $sName = FileFindNextFile($hSearch)
    While Not @error
        If $sName <> "." And $sName <> ".." Then
            Local $sFull = $sDir & "\" & $sName
            If StringInStr(FileGetAttrib($sFull), "D") > 0 Then
                _DD_WalkAu3Files($sFull)
            Else
                If StringRight(StringLower($sFull), 4) = ".au3" Then
                    Local $sKey = StringLower($sFull)
                    If Not $oDD_AllProjectFiles.Exists($sKey) Then $oDD_AllProjectFiles.Add($sKey, $sFull)
                EndIf
            EndIf
        EndIf
        $sName = FileFindNextFile($hSearch)
    WEnd

    FileClose($hSearch)
EndFunc   ;==>_DD_WalkAu3Files

Func _DD_CountFuncDeclarations($sPath)
    Local $aLines = _DD_GetFileLines($sPath)
    Local $iN = UBound($aLines)
    Local $iCount = 0
    Local $i = 0
    For $i = 0 To $iN - 1
        If StringRegExp($aLines[$i], "(?i)^\s*Func\s+[A-Za-z_]\w*\s*\(") Then $iCount += 1
    Next
    Return $iCount
EndFunc   ;==>_DD_CountFuncDeclarations

#endregion

; ============================================================================
#region Lecture des fichiers
; ============================================================================

; Cache le contenu texte brut de chaque fichier (cle = chemin en minuscules)
; pour eviter de relire le disque a chaque passe (includes, fonctions,
; variables). Les tableaux AutoIt transitant par un objet COM peuvent se
; comporter de facon inattendue : on ne met donc en cache que des chaines.
Func _DD_GetFileText($sPath)
    Local $sKey = StringLower($sPath)
    If $oDD_FileTextCache.Exists($sKey) Then Return $oDD_FileTextCache.Item($sKey)

    Local $sText = FileRead($sPath)
    If @error Then $sText = ""

    $oDD_FileTextCache.Add($sKey, $sText)
    Return $sText
EndFunc   ;==>_DD_GetFileText

Func _DD_GetFileLines($sPath)
    Local $sText = _DD_GetFileText($sPath)
    $sText = StringReplace($sText, @CR, "")
    Local $aLines = StringSplit($sText, @LF, $STR_NOCOUNT)
    Return $aLines
EndFunc   ;==>_DD_GetFileLines

#endregion

; ============================================================================
#region Indexation des fonctions
; ============================================================================

Func _DD_IndexFunctions()
    ConsoleWrite("[*] Indexation des fonctions..." & @CRLF)
    Local $i = 0
    For $i = 0 To $g_iFileCount - 1
        Local $sPath = $g_aFiles[$i][$FCOL_PATH]
        If FileExists($sPath) Then _DD_IndexFunctionsInFile($sPath)
    Next
    _DD_FindDuplicateAndSimilarFunctions()
    ConsoleWrite("[*] " & $g_iFuncCount & " fonction(s) indexee(s)." & @CRLF)
EndFunc   ;==>_DD_IndexFunctions

; Scanne un fichier ligne par ligne pour reperer les blocs Func...EndFunc.
; Les fonctions AutoIt ne peuvent pas etre imbriquees : rencontrer un
; nouveau "Func" alors qu'on est deja dans une fonction signale un
; EndFunc manquant sur la fonction precedente, sans jamais planter.
Func _DD_IndexFunctionsInFile($sPath)
    Local $aLines = _DD_GetFileLines($sPath)
    Local $iN = UBound($aLines)
    Local $bInFunc = False
    Local $sCurName = ""
    Local $sCurParams = ""
    Local $iCurStart = 0
    Local $i = 0

    For $i = 0 To $iN - 1
        Local $sLine = $aLines[$i]
        Local $iLineNo = $i + 1

        If Not $bInFunc Then
            Local $aMatch = StringRegExp($sLine, "(?i)^\s*Func\s+([A-Za-z_]\w*)\s*\(([^)]*)\)", 1)
            If Not @error Then
                $bInFunc = True
                $sCurName = $aMatch[0]
                $sCurParams = $aMatch[1]
                $iCurStart = $iLineNo
            EndIf
        Else
            If StringRegExp($sLine, "(?i)^\s*EndFunc\b") Then
                Local $sDomain = _DD_GuessDomainFromPath($sPath)
                If $sDomain = "COMMON" Then $sDomain = _DD_GuessDomainFromName($sCurName)
                _DD_AddFuncRow($sCurName, $sPath, $iCurStart, $iLineNo, $sCurParams, _DD_CountParams($sCurParams), $sDomain)
                $bInFunc = False
                $sCurName = ""
                $sCurParams = ""
            ElseIf StringRegExp($sLine, "(?i)^\s*Func\s+[A-Za-z_]\w*\s*\(") Then
                Local $sDomain2 = _DD_GuessDomainFromPath($sPath)
                If $sDomain2 = "COMMON" Then $sDomain2 = _DD_GuessDomainFromName($sCurName)
                _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_MISSING_ENDFUNC, $sDomain2, $sPath, $iCurStart, $sCurName, "", _
                    "La fonction '" & $sCurName & "' n'a pas de EndFunc avant la declaration suivante (ligne " & $iLineNo & ").", _
                    "Ajouter EndFunc a la fin de la fonction '" & $sCurName & "'.", "moyenne")
                Local $aMatch2 = StringRegExp($sLine, "(?i)^\s*Func\s+([A-Za-z_]\w*)\s*\(([^)]*)\)", 1)
                $sCurName = $aMatch2[0]
                $sCurParams = $aMatch2[1]
                $iCurStart = $iLineNo
            EndIf
        EndIf
    Next

    If $bInFunc Then
        Local $sDomain3 = _DD_GuessDomainFromPath($sPath)
        If $sDomain3 = "COMMON" Then $sDomain3 = _DD_GuessDomainFromName($sCurName)
        _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_MISSING_ENDFUNC, $sDomain3, $sPath, $iCurStart, $sCurName, "", _
            "La fonction '" & $sCurName & "' n'a pas de EndFunc avant la fin du fichier.", _
            "Ajouter EndFunc a la fin de la fonction '" & $sCurName & "'.", "moyenne")
    EndIf
EndFunc   ;==>_DD_IndexFunctionsInFile

Func _DD_CountParams($sParams)
    Local $s = StringStripWS($sParams, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING))
    If $s = "" Then Return 0
    Local $aParts = StringSplit($s, ",", $STR_NOCOUNT)
    Return UBound($aParts)
EndFunc   ;==>_DD_CountParams

; Detecte les fonctions definies plusieurs fois (meme nom, plusieurs
; emplacements) et les paires de noms tres proches (distance d'edition
; faible) qui trahissent souvent une faute de frappe entre deux
; implementations censees etre identiques (ex. HttpRouter vs HttpServer).
Func _DD_FindDuplicateAndSimilarFunctions()
    Local $oSeen = ObjCreate("Scripting.Dictionary")
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        Local $sKey = StringLower($g_aFunctions[$i][$UCOL_NAME])
        If $oSeen.Exists($sKey) Then
            $oSeen.Item($sKey) = $oSeen.Item($sKey) & "," & $i
        Else
            $oSeen.Add($sKey, String($i))
        EndIf
    Next

    Local $aKeys = $oSeen.Keys()
    Local $iK = 0
    For $iK = 0 To UBound($aKeys) - 1
        Local $sIdxList = $oSeen.Item($aKeys[$iK])
        If StringInStr($sIdxList, ",") > 0 Then
            Local $aIdx = StringSplit($sIdxList, ",", $STR_NOCOUNT)
            Local $sLocations = ""
            Local $j = 0
            For $j = 0 To UBound($aIdx) - 1
                Local $iRow = Number($aIdx[$j])
                If $sLocations <> "" Then $sLocations &= " | "
                $sLocations &= _DD_RelativePath($g_aFunctions[$iRow][$UCOL_FILE], $g_sProjectRoot) & ":" & $g_aFunctions[$iRow][$UCOL_LINE]
            Next
            Local $iFirstRow = Number($aIdx[0])
            _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_DUPLICATE_FUNCTION, $g_aFunctions[$iFirstRow][$UCOL_DOMAIN], _
                $g_aFunctions[$iFirstRow][$UCOL_FILE], $g_aFunctions[$iFirstRow][$UCOL_LINE], $g_aFunctions[$iFirstRow][$UCOL_NAME], _
                $g_aFunctions[$iFirstRow][$UCOL_NAME], _
                "La fonction '" & $g_aFunctions[$iFirstRow][$UCOL_NAME] & "' est definie plusieurs fois : " & $sLocations, _
                "Garder une seule implementation active et retirer l'inclusion ou la definition redondante.", "high")
        EndIf
    Next

    For $i = 0 To $g_iFuncCount - 1
        Local $j2 = 0
        For $j2 = $i + 1 To $g_iFuncCount - 1
            If StringLower($g_aFunctions[$i][$UCOL_NAME]) = StringLower($g_aFunctions[$j2][$UCOL_NAME]) Then ContinueLoop

            Local $iDist = _DD_Levenshtein($g_aFunctions[$i][$UCOL_NAME], $g_aFunctions[$j2][$UCOL_NAME])
            Local $iMaxLen = StringLen($g_aFunctions[$i][$UCOL_NAME])
            If StringLen($g_aFunctions[$j2][$UCOL_NAME]) > $iMaxLen Then $iMaxLen = StringLen($g_aFunctions[$j2][$UCOL_NAME])

            If $iDist > 0 And $iDist <= 2 And $iMaxLen >= 6 Then
                _DD_AddDiagnostic($DD_SEV_WARNING, $DD_CODE_SIMILAR_FUNCTION, $g_aFunctions[$i][$UCOL_DOMAIN], _
                    $g_aFunctions[$i][$UCOL_FILE], $g_aFunctions[$i][$UCOL_LINE], $g_aFunctions[$i][$UCOL_NAME], _
                    $g_aFunctions[$i][$UCOL_NAME], _
                    "Nom de fonction tres proche de '" & $g_aFunctions[$j2][$UCOL_NAME] & "' (" & _
                    _DD_RelativePath($g_aFunctions[$j2][$UCOL_FILE], $g_sProjectRoot) & ":" & $g_aFunctions[$j2][$UCOL_LINE] & _
                    "). Distance d'edition = " & $iDist & ".", _
                    "Verifier qu'il ne s'agit pas de deux implementations censees etre identiques (faute de frappe), ou renommer pour clarifier l'intention.", "moyenne")
            EndIf
        Next
    Next
EndFunc   ;==>_DD_FindDuplicateAndSimilarFunctions

Func _DD_FindEnclosingFunction($sFile, $iLine)
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        If StringLower($g_aFunctions[$i][$UCOL_FILE]) = StringLower($sFile) Then
            If $iLine >= $g_aFunctions[$i][$UCOL_LINE] And $iLine <= $g_aFunctions[$i][$UCOL_ENDLINE] Then Return $i
        EndIf
    Next
    Return -1
EndFunc   ;==>_DD_FindEnclosingFunction

Func _DD_FunctionExists($sName)
    Local $sKey = StringLower($sName)
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        If StringLower($g_aFunctions[$i][$UCOL_NAME]) = $sKey Then Return True
    Next
    Return False
EndFunc   ;==>_DD_FunctionExists

; Cherche, parmi toutes les fonctions indexees, le nom le plus proche
; (distance d'edition) de $sName. Retourne "" si rien d'assez proche.
Func _DD_ClosestFunctionName($sName)
    Local $sBest = ""
    Local $iBestDist = 999
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        Local $iDist = _DD_Levenshtein(StringLower($sName), StringLower($g_aFunctions[$i][$UCOL_NAME]))
        If $iDist > 0 And $iDist < $iBestDist Then
            $iBestDist = $iDist
            $sBest = $g_aFunctions[$i][$UCOL_NAME]
        EndIf
    Next
    If $sBest <> "" And $iBestDist <= 3 Then Return $sBest
    Return ""
EndFunc   ;==>_DD_ClosestFunctionName

#endregion

; ============================================================================
#region Analyse des variables
; ============================================================================

; Construit la liste des noms de variables "connus" a l'echelle du projet :
; toute variable declaree (Global/Local/Dim/Const/Static) ou affectee
; (meme sans mot-cle, ce qui est une declaration Global implicite valide en
; AutoIt des qu'on est hors Opt("MustDeclareVars",1)) n'importe ou dans le
; graphe d'inclusion. Volontairement permissif : mieux vaut rater un vrai
; probleme que d'inventer un faux positif (voir remarque en tete de
; fichier). La verite definitive reste la sortie du compilateur AutoIt.
Func _DD_IndexGlobalVariables()
    ConsoleWrite("[*] Indexation des variables (declarations + affectations)..." & @CRLF)

    Local $i = 0
    For $i = 0 To $g_iFileCount - 1
        Local $sPath = $g_aFiles[$i][$FCOL_PATH]
        If Not FileExists($sPath) Then ContinueLoop

        Local $aLines = _DD_GetFileLines($sPath)
        Local $iN = UBound($aLines)
        Local $j = 0
        For $j = 0 To $iN - 1
            Local $aNames = _DD_ExtractDeclaredNames($aLines[$j])
            Local $k = 0
            For $k = 0 To UBound($aNames) - 1
                Local $sKey = StringLower($aNames[$k])
                If $sKey <> "" And Not $oDD_GlobalVars.Exists($sKey) Then
                    $oDD_GlobalVars.Add($sKey, $sPath & "|" & String($j + 1))
                    $g_iVariableCount += 1
                EndIf
            Next
        Next
    Next

    For $i = 0 To $g_iFuncCount - 1
        Local $aParams = StringSplit($g_aFunctions[$i][$UCOL_PARAMS], ",", $STR_NOCOUNT)
        Local $p = 0
        For $p = 0 To UBound($aParams) - 1
            Local $aPName = StringRegExp($aParams[$p], "\$(\w+)", 1)
            If Not @error Then
                Local $sKey2 = StringLower($aPName[0])
                If Not $oDD_GlobalVars.Exists($sKey2) Then $oDD_GlobalVars.Add($sKey2, "param")
            EndIf
        Next
    Next

    ConsoleWrite("[*] " & $g_iVariableCount & " variable(s) distincte(s) recensee(s)." & @CRLF)
EndFunc   ;==>_DD_IndexGlobalVariables

; Extrait les noms de variables "declares" par une ligne de code :
;   - avec mot-cle Global/Local/Dim/Const/Static : tous les $xxx presents
;     apres le mot-cle (cible ET valeur d'initialisation eventuelle, par
;     construction volontairement large pour eviter les faux positifs) ;
;   - sans mot-cle : uniquement si $xxx est en tout debut de ligne suivi
;     d'un operateur d'affectation (affectation nue = Global implicite en
;     AutoIt quand MustDeclareVars est desactive).
Func _DD_ExtractDeclaredNames($sLine)
    Local $aNames[4]
    Local $iCount = 0
    Local $sTrim = StringStripWS($sLine, $STR_STRIPLEADING)

    If StringLeft($sTrim, 1) = ";" Then
        Local $aEmpty[0]
        Return $aEmpty
    EndIf

    If StringRegExp($sTrim, "(?i)^(Global|Local|Dim|Const|Static)\b") Then
        Local $sAfterKw = StringRegExpReplace($sTrim, "(?i)^(Global|Local|Dim|Const|Static)\s+", "")
        Local $aAll = StringRegExp($sAfterKw, "\$(\w+)", 3)
        If Not @error Then
            Local $j = 0
            For $j = 0 To UBound($aAll) - 1
                If $iCount >= UBound($aNames) Then ReDim $aNames[UBound($aNames) * 2]
                $aNames[$iCount] = $aAll[$j]
                $iCount += 1
            Next
        EndIf
    Else
        Local $aBare = StringRegExp($sTrim, "(?i)^\$(\w+)\s*(?:\[[^\]]*\])?\s*(?:\+=|-=|\*=|/=|&=|=(?!=))", 1)
        If Not @error Then
            If $iCount >= UBound($aNames) Then ReDim $aNames[UBound($aNames) * 2]
            $aNames[$iCount] = $aBare[0]
            $iCount += 1
        EndIf
    EndIf

    ReDim $aNames[$iCount]
    Return $aNames
EndFunc   ;==>_DD_ExtractDeclaredNames

; Vrai si $sKeyLower (nom de variable en minuscules, sans le $) ressemble
; a une constante standard AutoIt/Windows API plutot qu'a une variable du
; projet Dispatch (qui n'utilise jamais ces prefixes dans sa propre
; convention de nommage). Voir $g_aKnownConstantPrefixes plus haut.
Func _DD_IsLikelyBuiltinConstant($sKeyLower)
    Local $i = 0
    For $i = 0 To UBound($g_aKnownConstantExact) - 1
        If $sKeyLower = $g_aKnownConstantExact[$i] Then Return True
    Next

    Local $sUpper = StringUpper($sKeyLower)
    For $i = 0 To UBound($g_aKnownConstantPrefixes) - 1
        If StringLeft($sUpper, StringLen($g_aKnownConstantPrefixes[$i])) = $g_aKnownConstantPrefixes[$i] Then Return True
    Next

    Return False
EndFunc   ;==>_DD_IsLikelyBuiltinConstant

Func _DD_AnalyzeAllVariables()
    ConsoleWrite("[*] Analyse heuristique des variables (best-effort)..." & @CRLF)
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        _DD_AnalyzeFunctionVariables($i)
        _DD_CheckReDimSafety($i)
    Next
EndFunc   ;==>_DD_AnalyzeAllVariables

; Piege AutoIt reel et deterministe (pas une simple heuristique) : si la
; toute premiere reference a une variable dans une fonction est un ReDim,
; AutoIt ne la relie PAS automatiquement a une variable Global du meme nom
; declaree ailleurs, meme sans MustDeclareVars. Le ReDim leve alors
; "Variable used without being declared" au runtime -- uniquement quand
; cette fonction precise est appelee, ce que /ErrorStdOut ne declenche pas
; forcement pendant les quelques secondes ou l'on regarde tourner le
; script. C'est exactement le bug remonte sur _QueueClearRows (generation
; CMR) : la variable etait bien "Global $g_aQueueBlocks[0]" ailleurs, mais
; jamais relue avant le ReDim dans cette fonction precise. Confiance
; "high" car ce n'est pas une supposition : c'est un comportement connu et
; systematique du langage.
Func _DD_CheckReDimSafety($iFuncRow)
    Local $sName = $g_aFunctions[$iFuncRow][$UCOL_NAME]
    Local $sFile = $g_aFunctions[$iFuncRow][$UCOL_FILE]
    Local $iStart = $g_aFunctions[$iFuncRow][$UCOL_LINE]
    Local $iEnd = $g_aFunctions[$iFuncRow][$UCOL_ENDLINE]
    Local $sDomain = $g_aFunctions[$iFuncRow][$UCOL_DOMAIN]

    Local $oSeen = ObjCreate("Scripting.Dictionary")
    Local $oFlagged = ObjCreate("Scripting.Dictionary")

    Local $aParams = StringSplit($g_aFunctions[$iFuncRow][$UCOL_PARAMS], ",", $STR_NOCOUNT)
    Local $p = 0
    For $p = 0 To UBound($aParams) - 1
        Local $aPName = StringRegExp($aParams[$p], "\$(\w+)", 1)
        If Not @error Then
            Local $k = StringLower($aPName[0])
            If Not $oSeen.Exists($k) Then $oSeen.Add($k, True)
        EndIf
    Next

    Local $aLines = _DD_GetFileLines($sFile)
    Local $iLine = 0

    For $iLine = $iStart To $iEnd - 1
        If $iLine - 1 < 0 Or $iLine - 1 >= UBound($aLines) Then ContinueLoop
        Local $sLine = $aLines[$iLine - 1]

        Local $sTrim = StringStripWS($sLine, $STR_STRIPLEADING)
        If StringLeft($sTrim, 1) = ";" Then ContinueLoop

        Local $aReDim = StringRegExp($sLine, "(?i)^\s*ReDim\s+\$(\w+)", 1)
        If Not @error Then
            Local $sVar = $aReDim[0]
            Local $sKey = StringLower($sVar)
            If Not $oSeen.Exists($sKey) And Not $oFlagged.Exists($sKey) Then
                _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_REDIM_UNSAFE, $sDomain, $sFile, $iLine, $sName, "$" & $sVar, _
                    "ReDim sur '$" & $sVar & "' est la toute premiere reference a cette variable dans '" & $sName & _
                    "'. Sous AutoIt, ReDim ne se lie pas automatiquement a une variable existante dans ce cas et peut lever 'Variable used without being declared' au runtime des que cette fonction est appelee -- meme si '$" & _
                    $sVar & "' est bien declaree ailleurs.", _
                    "Ajouter 'Global $" & $sVar & "' (ou 'Local $" & $sVar & "' si elle doit rester locale) au tout debut de '" & $sName & "', avant ce ReDim.", "high")
                $oFlagged.Add($sKey, True)
            EndIf
        EndIf

        Local $aUsages = StringRegExp($sLine, "\$(\w+)", 3)
        If Not @error Then
            Local $u = 0
            For $u = 0 To UBound($aUsages) - 1
                Local $k2 = StringLower($aUsages[$u])
                If Not $oSeen.Exists($k2) Then $oSeen.Add($k2, True)
            Next
        EndIf
    Next
EndFunc   ;==>_DD_CheckReDimSafety

Func _DD_AnalyzeFunctionVariables($iFuncRow)
    Local $sName = $g_aFunctions[$iFuncRow][$UCOL_NAME]
    Local $sFile = $g_aFunctions[$iFuncRow][$UCOL_FILE]
    Local $iStart = $g_aFunctions[$iFuncRow][$UCOL_LINE]
    Local $iEnd = $g_aFunctions[$iFuncRow][$UCOL_ENDLINE]
    Local $sDomain = $g_aFunctions[$iFuncRow][$UCOL_DOMAIN]

    Local $oLocalKnown = ObjCreate("Scripting.Dictionary")

    Local $aParams = StringSplit($g_aFunctions[$iFuncRow][$UCOL_PARAMS], ",", $STR_NOCOUNT)
    Local $p = 0
    For $p = 0 To UBound($aParams) - 1
        Local $aPName = StringRegExp($aParams[$p], "\$(\w+)", 1)
        If Not @error Then
            Local $k = StringLower($aPName[0])
            If Not $oLocalKnown.Exists($k) Then $oLocalKnown.Add($k, True)
        EndIf
    Next

    Local $aLines = _DD_GetFileLines($sFile)
    Local $iLine = 0

    For $iLine = $iStart To $iEnd - 1
        If $iLine - 1 < 0 Or $iLine - 1 >= UBound($aLines) Then ContinueLoop
        Local $sLine = $aLines[$iLine - 1]

        Local $aDecl = _DD_ExtractDeclaredNames($sLine)
        Local $d = 0
        For $d = 0 To UBound($aDecl) - 1
            Local $k2 = StringLower($aDecl[$d])
            If $k2 <> "" And Not $oLocalKnown.Exists($k2) Then $oLocalKnown.Add($k2, True)
        Next

        Local $aForVar = StringRegExp($sLine, "(?i)^\s*For\s+\$(\w+)\s+(?:=|In\b)", 1)
        If Not @error Then
            Local $k3 = StringLower($aForVar[0])
            If Not $oLocalKnown.Exists($k3) Then $oLocalKnown.Add($k3, True)
        EndIf

        ; Idiome AutoIt courant : IsDeclared("nom") garde l'usage d'une
        ; variable qui peut ne pas exister selon le chemin d'execution
        ; (ex. _API_JobStatus dans StateService.au3). Ce n'est pas un bug :
        ; on considere le nom comme connu des qu'il apparait dans un tel
        ; garde-fou, meme sans declaration Local/Global classique.
        Local $aGuards = StringRegExp($sLine, '(?i)IsDeclared\s*\(\s*"(\w+)"', 3)
        If Not @error Then
            Local $g = 0
            For $g = 0 To UBound($aGuards) - 1
                Local $k4 = StringLower($aGuards[$g])
                If Not $oLocalKnown.Exists($k4) Then $oLocalKnown.Add($k4, True)
            Next
        EndIf
    Next

    For $iLine = $iStart To $iEnd - 1
        If $iLine - 1 < 0 Or $iLine - 1 >= UBound($aLines) Then ContinueLoop
        Local $sLine = $aLines[$iLine - 1]

        Local $sTrim = StringStripWS($sLine, $STR_STRIPLEADING)
        If StringLeft($sTrim, 1) = ";" Then ContinueLoop

        Local $aUsages = StringRegExp($sLine, "\$(\w+)", 3)
        If @error Then ContinueLoop

        Local $u = 0
        For $u = 0 To UBound($aUsages) - 1
            Local $sVar = $aUsages[$u]
            Local $sKeyU = StringLower($sVar)

            If $oLocalKnown.Exists($sKeyU) Then ContinueLoop
            If $oDD_GlobalVars.Exists($sKeyU) Then ContinueLoop
            If _DD_IsLikelyBuiltinConstant($sKeyU) Then ContinueLoop

            Local $sSuggestion = _DD_SuggestCloseName($sKeyU, $oLocalKnown)

            _DD_AddDiagnostic($DD_SEV_WARNING, $DD_CODE_UNDECLARED_VARIABLE, $sDomain, $sFile, $iLine, $sName, "$" & $sVar, _
                "Variable '$" & $sVar & "' utilisee dans '" & $sName & "' sans declaration Local/Global/Dim visible (analyse heuristique, a confirmer par le compilateur).", _
                $sSuggestion, "moyenne")

            ; on ne signale qu'une fois la meme variable par fonction
            $oLocalKnown.Add($sKeyU, True)
        Next
    Next
EndFunc   ;==>_DD_AnalyzeFunctionVariables

Func _DD_SuggestCloseName($sTarget, $oLocalKnown)
    Local $sBest = ""
    Local $iBestDist = 999

    Local $aLocalKeys = $oLocalKnown.Keys()
    Local $i = 0
    For $i = 0 To UBound($aLocalKeys) - 1
        Local $iDist = _DD_Levenshtein($sTarget, $aLocalKeys[$i])
        If $iDist > 0 And $iDist < $iBestDist Then
            $iBestDist = $iDist
            $sBest = $aLocalKeys[$i]
        EndIf
    Next

    If $sBest <> "" And $iBestDist <= 2 Then
        Return "Peut-etre une faute de frappe : variable proche declaree '$" & $sBest & _
            "' dans la meme fonction. Sinon, ajouter Local $" & $sTarget & " = ... ou Global $" & $sTarget & " = ..."
    EndIf

    Return "Verifier Globals.au3 (ou le fichier pertinent) ou ajouter une declaration : Local $" & $sTarget & _
        " = ... (Global si la variable doit etre partagee entre fonctions)."
EndFunc   ;==>_DD_SuggestCloseName

#endregion

; ============================================================================
#region Analyse des erreurs du compilateur
; ============================================================================

; Lance AutoIt3.exe /ErrorStdOut sur le script cible, capture stdout/stderr
; sans jamais ouvrir de fenetre visible, puis analyse la sortie. C'est la
; seule source de verite absolue : tout ce qu'elle rapporte est marque
; confiance "high".
Func _DD_RunCompilerAndAnalyze()
    If $g_sAutoItPath = "" Then
        _DD_AddDiagnostic($DD_SEV_WARNING, $DD_CODE_AUTOIT_NOT_FOUND, "COMMON", $g_sTargetScript, 0, "", "", _
            "AutoIt3.exe introuvable (recherche : dossier de l'outil, Program Files\AutoIt3, Program Files (x86)\AutoIt3, PATH systeme).", _
            "Installer AutoIt ou placer AutoIt3.exe a cote de " & $DD_TOOLNAME & ".au3.", "high")
        Return
    EndIf

    ConsoleWrite("[*] Compilation avec " & $g_sAutoItPath & " ..." & @CRLF)

    Local $sCmd = '"' & $g_sAutoItPath & '" /ErrorStdOut "' & $g_sTargetScript & '"'
    Local $iPid = Run($sCmd, $g_sProjectRoot, @SW_HIDE, BitOR($DD_STDOUT_CHILD, $DD_STDERR_CHILD))
    If @error Or $iPid = 0 Then
        _DD_AddDiagnostic($DD_SEV_ERROR, $DD_CODE_COMPILER_ERROR, "COMMON", $g_sTargetScript, 0, "", "", _
            "Impossible de lancer AutoIt3.exe.", "Verifier les droits d'execution et le chemin detecte : " & $g_sAutoItPath, "high")
        Return
    EndIf

    Local $sOutput = ""
    Local $hTimer = TimerInit()
    While ProcessExists($iPid)
        $sOutput &= StdoutRead($iPid)
        $sOutput &= StderrRead($iPid)
        If TimerDiff($hTimer) > 60000 Then
            ProcessClose($iPid)
            ExitLoop
        EndIf
        Sleep(25)
    WEnd
    $sOutput &= StdoutRead($iPid)
    $sOutput &= StderrRead($iPid)

    _DD_ParseCompilerOutput($sOutput)

    If StringStripWS($sOutput, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING)) = "" Then
        _DD_AddDiagnostic($DD_SEV_OK, $DD_CODE_OK, "COMMON", $g_sTargetScript, 0, "", "", _
            "Aucune erreur de compilation detectee par AutoIt3.exe /ErrorStdOut.", "", "high")
    EndIf
EndFunc   ;==>_DD_RunCompilerAndAnalyze

; Reconnait le format d'erreur AutoIt :
;   "chemin\fichier.au3" (95) : ==> Message.:
;   ligne de code source
;   ~~~~~~~~~~~~^ (pointeur eventuel)
; et enrichit chaque erreur avec la fonction englobante, le domaine, et une
; tentative d'identification precise du symbole en cause (sans jamais
; supposer que c'est la variable de boucle $i par defaut).
Func _DD_ParseCompilerOutput($sOutput)
    If StringStripWS($sOutput, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING)) = "" Then Return

    Local $sNorm = StringReplace($sOutput, @CR, "")
    Local $aLines = StringSplit($sNorm, @LF, $STR_NOCOUNT)
    Local $iN = UBound($aLines)
    Local Const $sErrPattern = '(?i)^"([^"]+)"\s*\((\d+)(?:,\s*\d+)?\)\s*:\s*(?:==>\s*)?(.*)$'

    Local $i = 0
    For $i = 0 To $iN - 1
        Local $aMatch = StringRegExp($aLines[$i], $sErrPattern, 1)
        If @error Then ContinueLoop

        Local $sFile = $aMatch[0]
        Local $iLine = Number($aMatch[1])
        Local $sMessage = StringStripWS($aMatch[2], $STR_STRIPTRAILING)
        If $sMessage = "" Then $sMessage = "Erreur de compilation AutoIt."

        Local $sSourceLine = ""
        If $i + 1 < $iN And Not StringRegExp($aLines[$i + 1], $sErrPattern) Then
            $sSourceLine = $aLines[$i + 1]
        EndIf

        Local $sLower = StringLower($sMessage)
        Local $sCode = $DD_CODE_COMPILER_ERROR
        If StringInStr($sLower, "variable used without being declared") Then
            $sCode = $DD_CODE_UNDECLARED_VARIABLE
        ElseIf StringInStr($sLower, "unknown function name") Then
            $sCode = $DD_CODE_UNKNOWN_FUNCTION
        ElseIf StringInStr($sLower, "duplicate function name") Then
            $sCode = $DD_CODE_DUPLICATE_FUNCTION
        ElseIf StringInStr($sLower, "incorrect number of parameters") Then
            $sCode = $DD_CODE_BAD_PARAM_COUNT
        ElseIf StringInStr($sLower, "endfunc") Then
            $sCode = $DD_CODE_MISSING_ENDFUNC
        ElseIf StringInStr($sLower, "endif") Or StringInStr($sLower, "unable to parse") Or StringInStr($sLower, "unterminated") Then
            $sCode = $DD_CODE_UNBALANCED_BLOCK
        EndIf

        Local $iFuncRow = _DD_FindEnclosingFunction($sFile, $iLine)
        Local $sFuncName = ""
        Local $sDomain = _DD_GuessDomainFromPath($sFile)
        If $iFuncRow >= 0 Then
            $sFuncName = $g_aFunctions[$iFuncRow][$UCOL_NAME]
            $sDomain = $g_aFunctions[$iFuncRow][$UCOL_DOMAIN]
        EndIf

        Local $sSymbol = ""
        Local $sSuggestion = ""
        Local $sConfidence = "high"

        If $sCode = $DD_CODE_UNDECLARED_VARIABLE And $sSourceLine <> "" Then
            Local $aVars = StringRegExp($sSourceLine, "\$(\w+)", 3)
            If Not @error Then
                If UBound($aVars) = 1 Then
                    $sSymbol = "$" & $aVars[0]
                    $sSuggestion = "Ajouter une declaration : Local $" & $aVars[0] & " = ... (ou Global si la variable doit etre partagee)."
                    $sConfidence = "high"
                Else
                    Local $sAll = ""
                    Local $v = 0
                    For $v = 0 To UBound($aVars) - 1
                        If $sAll <> "" Then $sAll &= ", "
                        $sAll &= "$" & $aVars[$v]
                    Next
                    $sSymbol = $sAll
                    $sSuggestion = "Plusieurs variables presentes sur cette ligne (" & $sAll & _
                        "). Analyser laquelle n'a pas de declaration Local/Global/Dim visible dans '" & $sFuncName & _
                        "' (ne pas supposer automatiquement qu'il s'agit de la variable de boucle)."
                    $sConfidence = "faible"
                EndIf
            EndIf

        ElseIf $sCode = $DD_CODE_DUPLICATE_FUNCTION Then
            ; Va chercher, dans l'index deja construit, TOUTES les
            ; definitions du nom concerne pour dire precisement ou est
            ; le doublon dans le code (pas juste "il y a un doublon").
            Local $sDupName = ""
            Local $aDupMatch = StringRegExp($sSourceLine, "(?i)Func\s+([A-Za-z_]\w*)", 1)
            If Not @error Then
                $sDupName = $aDupMatch[0]
            Else
                Local $aLineAtErr = _DD_GetFileLines($sFile)
                If $iLine - 1 >= 0 And $iLine - 1 < UBound($aLineAtErr) Then
                    Local $aDupMatch2 = StringRegExp($aLineAtErr[$iLine - 1], "(?i)Func\s+([A-Za-z_]\w*)", 1)
                    If Not @error Then $sDupName = $aDupMatch2[0]
                EndIf
            EndIf

            If $sDupName <> "" Then
                $sSymbol = $sDupName
                Local $sLocations = ""
                Local $f = 0
                For $f = 0 To $g_iFuncCount - 1
                    If StringLower($g_aFunctions[$f][$UCOL_NAME]) = StringLower($sDupName) Then
                        If $sLocations <> "" Then $sLocations &= " | "
                        $sLocations &= _DD_RelativePath($g_aFunctions[$f][$UCOL_FILE], $g_sProjectRoot) & ":" & $g_aFunctions[$f][$UCOL_LINE]
                    EndIf
                Next
                If $sLocations <> "" Then
                    $sSuggestion = "Definitions trouvees dans le code : " & $sLocations & _
                        ". Garder une seule implementation active et retirer l'inclusion ou la definition redondante."
                    $sConfidence = "high"
                Else
                    $sSuggestion = "Rechercher 'Func " & $sDupName & "(' dans le projet pour localiser les definitions en conflit."
                    $sConfidence = "moyenne"
                EndIf
            EndIf

        ElseIf $sCode = $DD_CODE_UNKNOWN_FUNCTION And $sSourceLine <> "" Then
            ; Identifie l'identifiant appele qui n'existe dans aucune
            ; fonction indexee, puis propose le nom connu le plus proche
            ; (faute de frappe probable).
            Local $aCalls = StringRegExp($sSourceLine, "([A-Za-z_]\w*)\s*\(", 3)
            Local $sCandidates = ""
            If Not @error Then
                Local $c = 0
                For $c = 0 To UBound($aCalls) - 1
                    Local $sCallName = StringRegExpReplace($aCalls[$c], "\($", "")
                    If Not _DD_FunctionExists($sCallName) Then
                        If $sCandidates <> "" Then $sCandidates &= ", "
                        $sCandidates &= $sCallName
                    EndIf
                Next
            EndIf

            If $sCandidates <> "" And StringInStr($sCandidates, ",") = 0 Then
                $sSymbol = $sCandidates
                Local $sClosest = _DD_ClosestFunctionName($sCandidates)
                If $sClosest <> "" Then
                    $sSuggestion = "Fonction '" & $sCandidates & "' introuvable. Faute de frappe probable, fonction proche existante : '" & $sClosest & "'."
                Else
                    $sSuggestion = "Fonction '" & $sCandidates & "' introuvable dans le projet. Verifier qu'elle est bien definie et que son fichier est inclus."
                EndIf
                $sConfidence = "high"
            ElseIf $sCandidates <> "" Then
                $sSymbol = $sCandidates
                $sSuggestion = "Plusieurs appels sur cette ligne ne correspondent a aucune fonction indexee (" & $sCandidates & _
                    "). Verifier chacun individuellement (l'un d'eux peut etre une fonction standard AutoIt non indexee ici)."
                $sConfidence = "faible"
            EndIf
        EndIf

        Local $sFullMessage = $sMessage
        If $sSourceLine <> "" Then $sFullMessage &= " | Ligne source : " & StringStripWS($sSourceLine, $STR_STRIPLEADING)

        _DD_AddDiagnostic($DD_SEV_ERROR, $sCode, $sDomain, $sFile, $iLine, $sFuncName, $sSymbol, $sFullMessage, $sSuggestion, $sConfidence)
    Next
EndFunc   ;==>_DD_ParseCompilerOutput

#endregion

; ============================================================================
#region Regles par domaine
; ============================================================================

; La table de donnees $g_aDomainRules est declaree plus haut (region "Etat
; global") car AutoIt execute les instructions de premier niveau dans
; l'ordre du fichier : elle doit exister avant l'appel a _DD_Main(). Les
; fonctions ci-dessous sont le seul point d'usage de cette table ; ajouter
; un nouveau domaine ne demande jamais de toucher a autre chose qu'a cette
; table.
Func _DD_GuessDomainFromPath($sPath)
    Local $sLower = StringLower($sPath)
    Local $iN = UBound($g_aDomainRules, 1)
    Local $i = 0
    For $i = 0 To $iN - 1
        If StringInStr($sLower, $g_aDomainRules[$i][0]) > 0 Then Return $g_aDomainRules[$i][1]
    Next
    Return "COMMON"
EndFunc   ;==>_DD_GuessDomainFromPath

Func _DD_GuessDomainFromName($sName)
    Local $sLower = StringLower($sName)
    Local $iN = UBound($g_aDomainRules, 1)
    Local $i = 0
    For $i = 0 To $iN - 1
        If StringInStr($g_aDomainRules[$i][0], "\") = 0 Then
            If StringInStr($sLower, $g_aDomainRules[$i][0]) > 0 Then Return $g_aDomainRules[$i][1]
        EndIf
    Next
    Return "COMMON"
EndFunc   ;==>_DD_GuessDomainFromName

#endregion

; ============================================================================
#region Orchestration des rapports
; ============================================================================

Func _DD_GenerateReports()
    Local $sReportsDir = @ScriptDir & "\reports"
    If Not FileExists($sReportsDir) Then DirCreate($sReportsDir)

    _DD_WriteTextReport($sReportsDir & "\DispatchDoctor-report.txt")
    _DD_WriteJsonReport($sReportsDir & "\DispatchDoctor-report.json")
    _DD_WriteHtmlReport($sReportsDir & "\DispatchDoctor-report.html")
    $g_sHtmlReportPath = $sReportsDir & "\DispatchDoctor-report.html"

    ConsoleWrite("[*] Rapports generes dans " & $sReportsDir & @CRLF)
    _DD_GuiLog("[*] Rapports generes dans " & $sReportsDir)
EndFunc   ;==>_DD_GenerateReports

Func _DD_DiagnosticPassesFilter($iRow)
    If $g_sDomainFilter = "" Then Return True
    Return (StringUpper($g_aDiagnostics[$iRow][$DCOL_DOMAIN]) = $g_sDomainFilter)
EndFunc   ;==>_DD_DiagnosticPassesFilter

Func _DD_ModeList()
    ConsoleWrite(@CRLF & "--- Fichiers detectes (" & $g_iFileCount & ") ---" & @CRLF)
    Local $i = 0
    For $i = 0 To $g_iFileCount - 1
        ConsoleWrite($g_aFiles[$i][$FCOL_RELPATH] & "  [" & $g_aFiles[$i][$FCOL_DOMAIN] & "]" & @CRLF)
    Next

    ConsoleWrite(@CRLF & "--- Fonctions detectees (" & $g_iFuncCount & ") ---" & @CRLF)
    For $i = 0 To $g_iFuncCount - 1
        If $g_sDomainFilter <> "" And StringUpper($g_aFunctions[$i][$UCOL_DOMAIN]) <> $g_sDomainFilter Then ContinueLoop
        ConsoleWrite($g_aFunctions[$i][$UCOL_NAME] & "(" & $g_aFunctions[$i][$UCOL_PARAMS] & ")  " & _
            _DD_RelativePath($g_aFunctions[$i][$UCOL_FILE], $g_sProjectRoot) & ":" & $g_aFunctions[$i][$UCOL_LINE] & _
            "  [" & $g_aFunctions[$i][$UCOL_DOMAIN] & "]" & @CRLF)
    Next

    Local $oDomains = ObjCreate("Scripting.Dictionary")
    For $i = 0 To $g_iFileCount - 1
        Local $k = $g_aFiles[$i][$FCOL_DOMAIN]
        If Not $oDomains.Exists($k) Then $oDomains.Add($k, 1)
    Next
    ConsoleWrite(@CRLF & "--- Domaines detectes ---" & @CRLF)
    Local $aDomKeys = $oDomains.Keys()
    For $i = 0 To UBound($aDomKeys) - 1
        ConsoleWrite("- " & $aDomKeys[$i] & @CRLF)
    Next
EndFunc   ;==>_DD_ModeList

#endregion

; ============================================================================
#region Rapports TXT
; ============================================================================

Func _DD_WriteTextReport($sPath)
    Local $hFile = FileOpen($sPath, 2)
    If $hFile = -1 Then Return

    Local $sDomLine = "Tous"
    If $g_sDomainFilter <> "" Then $sDomLine = $g_sDomainFilter
    Local $sAutoItLine = "non detecte"
    If $g_sAutoItPath <> "" Then $sAutoItLine = $g_sAutoItPath

    FileWriteLine($hFile, "========================================================")
    FileWriteLine($hFile, $DD_TOOLNAME & " v" & $DD_VERSION & " - Rapport de diagnostic")
    FileWriteLine($hFile, "========================================================")
    FileWriteLine($hFile, "Date              : " & _DD_NowString())
    FileWriteLine($hFile, "Projet            : " & $g_sProjectRoot)
    FileWriteLine($hFile, "Fichier principal : " & $g_sTargetScript)
    FileWriteLine($hFile, "Mode              : " & $g_sMode)
    FileWriteLine($hFile, "Domaine           : " & $sDomLine)
    FileWriteLine($hFile, "AutoIt detecte    : " & $sAutoItLine)
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "Fichiers analyses : " & $g_iFileCount)
    FileWriteLine($hFile, "Includes resolus  : " & $g_iIncludeEdgeCount)
    FileWriteLine($hFile, "Fonctions         : " & $g_iFuncCount)
    FileWriteLine($hFile, "Variables         : " & $g_iVariableCount)
    FileWriteLine($hFile, "Erreurs           : " & $g_iCountError)
    FileWriteLine($hFile, "Avertissements    : " & $g_iCountWarning)
    FileWriteLine($hFile, "Infos             : " & $g_iCountInfo)
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "--- Arbre des fichiers inclus ---")

    Local $i = 0
    For $i = 0 To $g_iFileCount - 1
        Local $sIndent = ""
        Local $d = 0
        For $d = 1 To Number($g_aFiles[$i][$FCOL_DEPTH])
            $sIndent &= "  "
        Next
        Local $sMark = "->"
        If Number($g_aFiles[$i][$FCOL_DEPTH]) = 0 Then $sMark = ""
        FileWriteLine($hFile, $sIndent & $sMark & " " & $g_aFiles[$i][$FCOL_RELPATH] & "  [" & $g_aFiles[$i][$FCOL_DOMAIN] & "]")
    Next

    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "--- Diagnostics ---")

    For $i = 0 To $g_iDiagCount - 1
        If Not _DD_DiagnosticPassesFilter($i) Then ContinueLoop
        FileWriteLine($hFile, "")
        FileWriteLine($hFile, "[" & $g_aDiagnostics[$i][$DCOL_SEVERITY] & "] " & $g_aDiagnostics[$i][$DCOL_CODE])
        If $g_aDiagnostics[$i][$DCOL_SYMBOL] <> "" Then FileWriteLine($hFile, "Symbole      : " & $g_aDiagnostics[$i][$DCOL_SYMBOL])
        FileWriteLine($hFile, "Fichier      : " & _DD_RelativePath($g_aDiagnostics[$i][$DCOL_FILE], $g_sProjectRoot))
        If Number($g_aDiagnostics[$i][$DCOL_LINE]) > 0 Then FileWriteLine($hFile, "Ligne        : " & $g_aDiagnostics[$i][$DCOL_LINE])
        If $g_aDiagnostics[$i][$DCOL_FUNCTION] <> "" Then FileWriteLine($hFile, "Fonction     : " & $g_aDiagnostics[$i][$DCOL_FUNCTION])
        FileWriteLine($hFile, "Domaine      : " & $g_aDiagnostics[$i][$DCOL_DOMAIN])
        FileWriteLine($hFile, "Message      : " & $g_aDiagnostics[$i][$DCOL_MESSAGE])
        If $g_aDiagnostics[$i][$DCOL_SUGGESTION] <> "" Then FileWriteLine($hFile, "Suggestion   : " & $g_aDiagnostics[$i][$DCOL_SUGGESTION])
        FileWriteLine($hFile, "Confiance    : " & $g_aDiagnostics[$i][$DCOL_CONFIDENCE])
    Next

    FileClose($hFile)
EndFunc   ;==>_DD_WriteTextReport

#endregion

; ============================================================================
#region Rapports JSON
; ============================================================================

Func _DD_JsonEscape($s)
    Local $r = $s
    $r = StringReplace($r, "\", "\\")
    $r = StringReplace($r, '"', '\"')
    $r = StringReplace($r, @CRLF, "\n")
    $r = StringReplace($r, @LF, "\n")
    $r = StringReplace($r, @CR, "\n")
    $r = StringReplace($r, @TAB, "\t")
    Return $r
EndFunc   ;==>_DD_JsonEscape

Func _DD_WriteJsonReport($sPath)
    Local $hFile = FileOpen($sPath, 2)
    If $hFile = -1 Then Return

    Local $sDomLine = "ALL"
    If $g_sDomainFilter <> "" Then $sDomLine = $g_sDomainFilter

    FileWriteLine($hFile, "{")
    FileWriteLine($hFile, '  "tool": "' & $DD_TOOLNAME & '",')
    FileWriteLine($hFile, '  "version": "' & $DD_VERSION & '",')
    FileWriteLine($hFile, '  "date": "' & _DD_JsonEscape(_DD_NowString()) & '",')
    FileWriteLine($hFile, '  "project": "' & _DD_JsonEscape($g_sProjectRoot) & '",')
    FileWriteLine($hFile, '  "mainFile": "' & _DD_JsonEscape($g_sTargetScript) & '",')
    FileWriteLine($hFile, '  "mode": "' & _DD_JsonEscape($g_sMode) & '",')
    FileWriteLine($hFile, '  "domain": "' & _DD_JsonEscape($sDomLine) & '",')
    FileWriteLine($hFile, '  "autoItPath": "' & _DD_JsonEscape($g_sAutoItPath) & '",')
    FileWriteLine($hFile, '  "stats": {')
    FileWriteLine($hFile, '    "files": ' & $g_iFileCount & ",")
    FileWriteLine($hFile, '    "includes": ' & $g_iIncludeEdgeCount & ",")
    FileWriteLine($hFile, '    "functions": ' & $g_iFuncCount & ",")
    FileWriteLine($hFile, '    "variables": ' & $g_iVariableCount & ",")
    FileWriteLine($hFile, '    "errors": ' & $g_iCountError & ",")
    FileWriteLine($hFile, '    "warnings": ' & $g_iCountWarning & ",")
    FileWriteLine($hFile, '    "infos": ' & $g_iCountInfo)
    FileWriteLine($hFile, "  },")
    FileWriteLine($hFile, '  "diagnostics": [')

    Local $bFirst = True
    Local $i = 0
    For $i = 0 To $g_iDiagCount - 1
        If Not _DD_DiagnosticPassesFilter($i) Then ContinueLoop
        If Not $bFirst Then FileWriteLine($hFile, ",")
        $bFirst = False

        Local $sLine = "    {"
        $sLine &= '"severity": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_SEVERITY]) & '", '
        $sLine &= '"code": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_CODE]) & '", '
        $sLine &= '"domain": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_DOMAIN]) & '", '
        $sLine &= '"file": "' & _DD_JsonEscape(_DD_RelativePath($g_aDiagnostics[$i][$DCOL_FILE], $g_sProjectRoot)) & '", '
        $sLine &= '"line": ' & Number($g_aDiagnostics[$i][$DCOL_LINE]) & ", "
        $sLine &= '"function": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_FUNCTION]) & '", '
        $sLine &= '"symbol": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_SYMBOL]) & '", '
        $sLine &= '"message": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_MESSAGE]) & '", '
        $sLine &= '"suggestion": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_SUGGESTION]) & '", '
        $sLine &= '"confidence": "' & _DD_JsonEscape($g_aDiagnostics[$i][$DCOL_CONFIDENCE]) & '"'
        $sLine &= "}"
        FileWrite($hFile, $sLine)
    Next

    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "  ]")
    FileWriteLine($hFile, "}")

    FileClose($hFile)
EndFunc   ;==>_DD_WriteJsonReport

#endregion

; ============================================================================
#region Rapports HTML
; ============================================================================

Func _DD_HtmlEscape($s)
    Local $r = $s
    $r = StringReplace($r, "&", "&amp;")
    $r = StringReplace($r, "<", "&lt;")
    $r = StringReplace($r, ">", "&gt;")
    $r = StringReplace($r, '"', "&quot;")
    Return $r
EndFunc   ;==>_DD_HtmlEscape

Func _DD_WriteHtmlReport($sPath)
    Local $hFile = FileOpen($sPath, 2)
    If $hFile = -1 Then Return

    Local $sDomLine = "Tous"
    If $g_sDomainFilter <> "" Then $sDomLine = $g_sDomainFilter
    Local $sAutoItLine = "non detecte"
    If $g_sAutoItPath <> "" Then $sAutoItLine = $g_sAutoItPath

    FileWriteLine($hFile, "<!DOCTYPE html><html lang='fr'><head><meta charset='utf-8'>")
    FileWriteLine($hFile, "<title>" & $DD_TOOLNAME & " - Rapport</title>")
    FileWriteLine($hFile, "<style>")
    FileWriteLine($hFile, "body{font-family:Segoe UI,Arial,sans-serif;background:#1e1e1e;color:#ddd;margin:20px}")
    FileWriteLine($hFile, "h1{color:#4da6ff}")
    FileWriteLine($hFile, "table{border-collapse:collapse;width:100%;margin-top:10px}")
    FileWriteLine($hFile, "td,th{border:1px solid #444;padding:6px 8px;text-align:left;vertical-align:top;font-size:13px}")
    FileWriteLine($hFile, "th{background:#2d2d2d}")
    FileWriteLine($hFile, ".ERROR{background:#4a1515;color:#ff8080}")
    FileWriteLine($hFile, ".WARNING{background:#4a3a15;color:#ffcc66}")
    FileWriteLine($hFile, ".INFO{background:#153a4a;color:#66ccff}")
    FileWriteLine($hFile, ".OK{background:#154a20;color:#66ff99}")
    FileWriteLine($hFile, ".stats span{display:inline-block;margin-right:20px;padding:6px 12px;background:#2d2d2d;border-radius:4px}")
    FileWriteLine($hFile, "</style></head><body>")
    FileWriteLine($hFile, "<h1>" & $DD_TOOLNAME & " v" & $DD_VERSION & " - Rapport de diagnostic</h1>")
    FileWriteLine($hFile, "<p>Date : " & _DD_HtmlEscape(_DD_NowString()) & "<br>")
    FileWriteLine($hFile, "Projet : " & _DD_HtmlEscape($g_sProjectRoot) & "<br>")
    FileWriteLine($hFile, "Fichier principal : " & _DD_HtmlEscape($g_sTargetScript) & "<br>")
    FileWriteLine($hFile, "Mode : " & _DD_HtmlEscape($g_sMode) & "<br>")
    FileWriteLine($hFile, "Domaine : " & _DD_HtmlEscape($sDomLine) & "<br>")
    FileWriteLine($hFile, "AutoIt : " & _DD_HtmlEscape($sAutoItLine) & "</p>")

    FileWriteLine($hFile, "<div class='stats'>")
    FileWriteLine($hFile, "<span>Fichiers : " & $g_iFileCount & "</span>")
    FileWriteLine($hFile, "<span>Includes : " & $g_iIncludeEdgeCount & "</span>")
    FileWriteLine($hFile, "<span>Fonctions : " & $g_iFuncCount & "</span>")
    FileWriteLine($hFile, "<span>Variables : " & $g_iVariableCount & "</span>")
    FileWriteLine($hFile, "<span class='ERROR'>Erreurs : " & $g_iCountError & "</span>")
    FileWriteLine($hFile, "<span class='WARNING'>Avertissements : " & $g_iCountWarning & "</span>")
    FileWriteLine($hFile, "<span class='INFO'>Infos : " & $g_iCountInfo & "</span>")
    FileWriteLine($hFile, "</div>")

    FileWriteLine($hFile, "<h2>Diagnostics</h2>")
    FileWriteLine($hFile, "<table><tr><th>Severite</th><th>Code</th><th>Domaine</th><th>Fichier</th><th>Ligne</th>" & _
        "<th>Fonction</th><th>Symbole</th><th>Message</th><th>Suggestion</th><th>Confiance</th></tr>")

    Local $i = 0
    For $i = 0 To $g_iDiagCount - 1
        If Not _DD_DiagnosticPassesFilter($i) Then ContinueLoop
        FileWriteLine($hFile, "<tr class='" & $g_aDiagnostics[$i][$DCOL_SEVERITY] & "'>" & _
            "<td>" & $g_aDiagnostics[$i][$DCOL_SEVERITY] & "</td>" & _
            "<td>" & $g_aDiagnostics[$i][$DCOL_CODE] & "</td>" & _
            "<td>" & $g_aDiagnostics[$i][$DCOL_DOMAIN] & "</td>" & _
            "<td>" & _DD_HtmlEscape(_DD_RelativePath($g_aDiagnostics[$i][$DCOL_FILE], $g_sProjectRoot)) & "</td>" & _
            "<td>" & $g_aDiagnostics[$i][$DCOL_LINE] & "</td>" & _
            "<td>" & _DD_HtmlEscape($g_aDiagnostics[$i][$DCOL_FUNCTION]) & "</td>" & _
            "<td>" & _DD_HtmlEscape($g_aDiagnostics[$i][$DCOL_SYMBOL]) & "</td>" & _
            "<td>" & _DD_HtmlEscape($g_aDiagnostics[$i][$DCOL_MESSAGE]) & "</td>" & _
            "<td>" & _DD_HtmlEscape($g_aDiagnostics[$i][$DCOL_SUGGESTION]) & "</td>" & _
            "<td>" & $g_aDiagnostics[$i][$DCOL_CONFIDENCE] & "</td></tr>")
    Next

    FileWriteLine($hFile, "</table>")

    FileWriteLine($hFile, "<h2>Fichiers inclus</h2><ul>")
    For $i = 0 To $g_iFileCount - 1
        FileWriteLine($hFile, "<li>" & _DD_HtmlEscape($g_aFiles[$i][$FCOL_RELPATH]) & " [" & $g_aFiles[$i][$FCOL_DOMAIN] & "]</li>")
    Next
    FileWriteLine($hFile, "</ul>")

    FileWriteLine($hFile, "</body></html>")
    FileClose($hFile)
EndFunc   ;==>_DD_WriteHtmlReport

#endregion

; ============================================================================
#region Client HTTP (pour le Test CMR)
; ============================================================================

; Test rapide et non bloquant : le port TCP est-il ouvert localement ?
; Utilise pour savoir si un serveur Dispatch ecoute deja avant de tester,
; et pour savoir quand l'instance lancee par DispatchDoctor est prete.
Func _DD_IsPortOpen($sHost, $iPort)
    If Not $g_bTcpStarted Then
        TCPStartup()
        $g_bTcpStarted = True
    EndIf

    Local $iSocket = TCPConnect($sHost, $iPort)
    If $iSocket = -1 Then Return False
    TCPCloseSocket($iSocket)
    Return True
EndFunc   ;==>_DD_IsPortOpen

; Envoie une requete HTTP POST minimale et retourne le corps de la
; reponse (ou "" en cas d'echec, avec le detail dans $g_sLastHttpError).
; Implementation volontairement simple (TCP brut, pas de gestion du
; chunked-encoding) : suffisant pour dialoguer avec le serveur maison de
; Dispatch (HttpRouter.au3), qui repond toujours en Content-Length simple.
Func _DD_HttpPost($sHost, $iPort, $sPath, $sBody)
    $g_sLastHttpError = ""

    If Not $g_bTcpStarted Then
        TCPStartup()
        $g_bTcpStarted = True
    EndIf

    Local $iSocket = TCPConnect($sHost, $iPort)
    If $iSocket = -1 Then
        $g_sLastHttpError = "connexion impossible a " & $sHost & ":" & $iPort
        Return ""
    EndIf

    Local $sReq = "POST " & $sPath & " HTTP/1.1" & @CRLF & _
        "Host: " & $sHost & ":" & $iPort & @CRLF & _
        "Content-Type: application/json" & @CRLF & _
        "Content-Length: " & StringLen($sBody) & @CRLF & _
        "Connection: close" & @CRLF & @CRLF & $sBody

    TCPSend($iSocket, StringToBinary($sReq, 4))
    If @error Then
        $g_sLastHttpError = "echec d'envoi de la requete HTTP"
        TCPCloseSocket($iSocket)
        Return ""
    EndIf

    Local $sResp = ""
    Local $hTimer = TimerInit()
    While TimerDiff($hTimer) < 10000
        Local $sChunk = TCPRecv($iSocket, 65536)
        If @error Then ExitLoop
        If $sChunk <> "" Then
            $sResp &= $sChunk
        Else
            Sleep(20)
        EndIf
    WEnd

    TCPCloseSocket($iSocket)

    Local $iPos = StringInStr($sResp, @CRLF & @CRLF)
    If $iPos > 0 Then Return StringMid($sResp, $iPos + 4)
    Return $sResp
EndFunc   ;==>_DD_HttpPost

; Extraction naive d'une valeur chaine depuis un petit JSON plat, du style
; de ceux produits par HttpRouter.au3 ($sKey":"valeur"). Suffisant pour
; lire "status"/"message" dans les reponses de _CMR_StatusJSON() ; pas un
; parseur JSON general.
Func _DD_JsonGetValue($sJson, $sKey)
    Local $aMatch = StringRegExp($sJson, '"' & $sKey & '"\s*:\s*"([^"]*)"', 1)
    If Not @error Then Return $aMatch[0]

    Local $aMatchBool = StringRegExp($sJson, '"' & $sKey & '"\s*:\s*(true|false)', 1)
    If Not @error Then Return $aMatchBool[0]

    Return ""
EndFunc   ;==>_DD_JsonGetValue

#endregion

; ============================================================================
#region Test d'action reelle (CMR / COMAT / ETMS ...)
; ============================================================================

; Lance une VRAIE instance de Dispatch (dediee au test, avec /ErrorStdOut
; pour que les erreurs runtime partent en texte au lieu d'ouvrir le popup
; "AutoIt Error" bloquant), declenche l'action choisie via son API HTTP
; reelle (POST /api/action, voir HttpRouter.au3), surveille le resultat,
; et si ca plante, croise l'erreur avec tout ce que l'analyse statique a
; deja trouve dans la meme fonction (doublons, ReDim dangereux, etc.).
;
; Attention : ceci execute reellement l'action (pas une simulation). Voir
; l'avertissement specifique de chaque action dans $g_aTestableActions.
Func _DD_RunActionTest($sActionLabel, $sParam1, $sParam2)
    If Not $g_bGuiActive Then Return
    If $g_bActionTestRunning Then
        _DD_GuiLog("[WARNING] Un test est deja en cours, attends qu'il se termine avant d'en relancer un.")
        Return
    EndIf

    Local $iRow = -1
    Local $t = 0
    For $t = 0 To UBound($g_aTestableActions, 1) - 1
        If $g_aTestableActions[$t][0] = $sActionLabel Then
            $iRow = $t
            ExitLoop
        EndIf
    Next
    If $iRow = -1 Then
        _DD_GuiLog("[ERROR] Action de test inconnue : '" & $sActionLabel & "'.")
        Return
    EndIf

    Local $sActionName = $g_aTestableActions[$iRow][1]
    Local $sField1 = $g_aTestableActions[$iRow][2]
    Local $sField2 = $g_aTestableActions[$iRow][3]
    Local $sWarning = $g_aTestableActions[$iRow][4]

    If StringStripWS($sParam1, 3) = "" Then
        _DD_GuiLog("[ERROR] Test '" & $sActionName & "' : le parametre principal (" & $sField1 & ") est vide.")
        Return
    EndIf
    If $sField2 <> "" And StringStripWS($sParam2, 3) = "" Then
        _DD_GuiLog("[ERROR] Test '" & $sActionName & "' : le parametre secondaire (" & $sField2 & ") est requis pour cette action.")
        Return
    EndIf

    $g_bActionTestRunning = True
    _DD_GuiLog("")
    _DD_GuiLog("=== Test '" & $sActionName & "' ===")
    _DD_GuiLog("Attention : " & $sWarning)

    _DD_GuiSetStatus("Test '" & $sActionName & "' : verification du port " & $DD_DISPATCH_PORT & "...")
    If _DD_IsPortOpen("127.0.0.1", $DD_DISPATCH_PORT) Then
        _DD_GuiLog("[ERROR] Le port " & $DD_DISPATCH_PORT & " est deja utilise (une instance de Dispatch tourne peut-etre deja). " & _
            "Ferme-la avant de lancer un test : je ne peux capturer les erreurs que d'une instance que je lance moi-meme avec /ErrorStdOut.")
        _DD_GuiSetStatus("Test annule : port deja occupe.")
        $g_bActionTestRunning = False
        Return
    EndIf

    If $g_sAutoItPath = "" Then _DD_DetectAutoIt()
    If $g_sAutoItPath = "" Then
        _DD_GuiLog("[ERROR] AutoIt3.exe introuvable, impossible de lancer le test.")
        $g_bActionTestRunning = False
        Return
    EndIf

    Local $sMainScript = $g_sProjectRoot & "\" & $DD_DEFAULT_MAIN
    If Not FileExists($sMainScript) Then
        _DD_GuiLog("[ERROR] " & $DD_DEFAULT_MAIN & " introuvable dans " & $g_sProjectRoot & ".")
        $g_bActionTestRunning = False
        Return
    EndIf

    _DD_GuiSetStatus("Test '" & $sActionName & "' : demarrage de Dispatch (instance dediee)...")
    Local $sCmd = '"' & $g_sAutoItPath & '" /ErrorStdOut "' & $sMainScript & '"'
    Local $iPid = Run($sCmd, $g_sProjectRoot, @SW_HIDE, BitOR($DD_STDOUT_CHILD, $DD_STDERR_CHILD))
    If @error Or $iPid = 0 Then
        _DD_GuiLog("[ERROR] Impossible de lancer Dispatch.")
        $g_bActionTestRunning = False
        Return
    EndIf
    $g_iActionTestPid = $iPid

    Local $sOutput = ""
    Local $hBootTimer = TimerInit()
    Local $bPortReady = False
    While TimerDiff($hBootTimer) < 15000
        $sOutput &= StdoutRead($iPid)
        $sOutput &= StderrRead($iPid)
        If Not ProcessExists($iPid) Then ExitLoop
        If _DD_IsPortOpen("127.0.0.1", $DD_DISPATCH_PORT) Then
            $bPortReady = True
            ExitLoop
        EndIf
        Sleep(150)
        _DD_GuiPump()
    WEnd

    If Not ProcessExists($iPid) Then
        $sOutput &= StdoutRead($iPid)
        $sOutput &= StderrRead($iPid)
        $g_iActionTestPid = 0
        _DD_FinishActionTest($sActionName, $sField1, $sParam1, $sField2, $sParam2, "CRASH_DEMARRAGE", $sOutput, "")
        $g_bActionTestRunning = False
        Return
    EndIf

    If Not $bPortReady Then
        ProcessClose($iPid)
        $g_iActionTestPid = 0
        _DD_FinishActionTest($sActionName, $sField1, $sParam1, $sField2, $sParam2, "TIMEOUT_DEMARRAGE", $sOutput, "")
        $g_bActionTestRunning = False
        Return
    EndIf

    Local $sParam2Info = ""
    If $sField2 <> "" Then $sParam2Info = ", " & $sField2 & "=" & $sParam2
    _DD_GuiSetStatus("Test '" & $sActionName & "' : envoi de l'action (" & $sField1 & "=" & $sParam1 & $sParam2Info & ")...")

    Local $sPayload = '{"action":"' & $sActionName & '","' & $sField1 & '":"' & _DD_JsonEscape($sParam1) & '"'
    If $sField2 <> "" Then $sPayload &= ',"' & $sField2 & '":"' & _DD_JsonEscape($sParam2) & '"'
    $sPayload &= '}'

    _DD_HttpPost("127.0.0.1", $DD_DISPATCH_PORT, "/api/action", $sPayload)
    If $g_sLastHttpError <> "" Then
        ProcessClose($iPid)
        $g_iActionTestPid = 0
        _DD_FinishActionTest($sActionName, $sField1, $sParam1, $sField2, $sParam2, "ECHEC_HTTP", $sOutput, "")
        $g_bActionTestRunning = False
        Return
    EndIf

    _DD_GuiSetStatus("Test '" & $sActionName & "' : execution en cours, surveillance...")
    Local $hTimer = TimerInit()
    Local $sOutcome = "TIMEOUT"
    Local $sFinalStatusJson = ""

    While TimerDiff($hTimer) < 60000
        $sOutput &= StdoutRead($iPid)
        $sOutput &= StderrRead($iPid)

        If Not ProcessExists($iPid) Then
            $sOutcome = "CRASH"
            ExitLoop
        EndIf

        ; CMR_STATUS est la seule action qui expose un vrai suivi d'etat
        ; aujourd'hui ; pour les autres, on se contente de surveiller un
        ; eventuel crash pendant la duree de surveillance.
        If $sActionName = "CMR_GENERATE" Then
            Local $sStatusResp = _DD_HttpPost("127.0.0.1", $DD_DISPATCH_PORT, "/api/action", '{"action":"CMR_STATUS"}')
            If $g_sLastHttpError = "" And $sStatusResp <> "" Then
                If _DD_JsonGetValue($sStatusResp, "running") = "false" Then
                    $sOutcome = "OK"
                    $sFinalStatusJson = $sStatusResp
                    ExitLoop
                EndIf
            EndIf
        EndIf

        Sleep(500)
        _DD_GuiPump()
    WEnd

    $sOutput &= StdoutRead($iPid)
    $sOutput &= StderrRead($iPid)

    If $sOutcome = "TIMEOUT" And $sActionName <> "CMR_GENERATE" Then $sOutcome = "AUCUN_CRASH_DETECTE"

    If ProcessExists($iPid) Then ProcessClose($iPid)
    $g_iActionTestPid = 0

    _DD_FinishActionTest($sActionName, $sField1, $sParam1, $sField2, $sParam2, $sOutcome, $sOutput, $sFinalStatusJson)
    $g_bActionTestRunning = False
EndFunc   ;==>_DD_RunActionTest

; Analyse le resultat d'un test : si crash, reutilise _DD_ParseCompilerOutput
; (meme moteur que le mode "compile") pour identifier fichier/ligne/cause,
; puis recoupe avec les diagnostics DEJA connus (analyse statique faite
; plus tot) pour la meme fonction -- exactement ce qui permet de dire
; "cette erreur est probablement liee a X, deja repere avant meme le test".
Func _DD_FinishActionTest($sActionName, $sField1, $sParam1, $sField2, $sParam2, $sOutcome, $sOutput, $sFinalStatusJson)
    Local $iDiagBefore = $g_iDiagCount
    Local $sCrashDetail = ""
    Local $sRelated = ""

    If $sOutcome = "CRASH" Or $sOutcome = "CRASH_DEMARRAGE" Then
        _DD_ParseCompilerOutput($sOutput)
        Local $d = 0
        For $d = $iDiagBefore To $g_iDiagCount - 1
            $sCrashDetail &= @CRLF & "  [" & $g_aDiagnostics[$d][$DCOL_CODE] & "] " & $g_aDiagnostics[$d][$DCOL_MESSAGE]
            If $g_aDiagnostics[$d][$DCOL_SUGGESTION] <> "" Then $sCrashDetail &= @CRLF & "    -> " & $g_aDiagnostics[$d][$DCOL_SUGGESTION]

            Local $sCrashFile = $g_aDiagnostics[$d][$DCOL_FILE]
            Local $iCrashLine = $g_aDiagnostics[$d][$DCOL_LINE]
            Local $iFuncRow = _DD_FindEnclosingFunction($sCrashFile, $iCrashLine)
            If $iFuncRow >= 0 Then
                Local $iFStart = $g_aFunctions[$iFuncRow][$UCOL_LINE]
                Local $iFEnd = $g_aFunctions[$iFuncRow][$UCOL_ENDLINE]
                Local $r = 0
                For $r = 0 To $iDiagBefore - 1
                    If StringLower($g_aDiagnostics[$r][$DCOL_FILE]) = StringLower($sCrashFile) Then
                        Local $iRLine = $g_aDiagnostics[$r][$DCOL_LINE]
                        If $iRLine >= $iFStart And $iRLine <= $iFEnd Then
                            $sRelated &= @CRLF & "  [deja detecte en analyse statique] [" & $g_aDiagnostics[$r][$DCOL_CODE] & _
                                "] ligne " & $iRLine & " : " & $g_aDiagnostics[$r][$DCOL_MESSAGE]
                        EndIf
                    EndIf
                Next
            EndIf
        Next
        If $sCrashDetail = "" Then $sCrashDetail = @CRLF & "  (erreur non reconnue automatiquement, voir la sortie brute dans le rapport)"
    EndIf

    Local $sSummary = ""
    Switch $sOutcome
        Case "OK"
            $sSummary = "Termine sans crash. " & _DD_JsonGetValue($sFinalStatusJson, "message")
            _DD_GuiSetStatus("Test '" & $sActionName & "' : termine sans crash.")
            _DD_GuiLog("[OK] " & $sSummary)

        Case "CRASH", "CRASH_DEMARRAGE"
            $sSummary = "Dispatch a plante pendant le test." & $sCrashDetail
            If $sRelated <> "" Then $sSummary &= @CRLF & "Lien possible avec l'analyse statique deja faite :" & $sRelated
            _DD_GuiSetStatus("Test '" & $sActionName & "' : CRASH detecte (voir journal et rapport).")
            _DD_GuiLog("[ERROR] Dispatch a plante pendant le test '" & $sActionName & "'." & $sCrashDetail)
            If $sRelated <> "" Then _DD_GuiLog("[INFO] Lien avec l'analyse statique :" & $sRelated)

        Case "TIMEOUT"
            $sSummary = "CMR toujours 'running' apres le delai de surveillance (pas de crash detecte, mais pas confirme termine)."
            _DD_GuiSetStatus("Test '" & $sActionName & "' : pas de crash, mais toujours en cours au bout du delai.")
            _DD_GuiLog("[WARNING] " & $sSummary)

        Case "AUCUN_CRASH_DETECTE"
            $sSummary = "Aucun crash pendant la surveillance. Cette action n'a pas de statut interrogeable comme CMR : " & _
                "ce n'est pas une confirmation de succes, seulement l'absence de plantage observe."
            _DD_GuiSetStatus("Test '" & $sActionName & "' : aucun crash observe.")
            _DD_GuiLog("[OK] " & $sSummary)

        Case "TIMEOUT_DEMARRAGE"
            $sSummary = "Le serveur Dispatch n'a jamais ouvert le port " & $DD_DISPATCH_PORT & " (demarrage anormalement long ou bloque)."
            _DD_GuiSetStatus("Test '" & $sActionName & "' : demarrage bloque.")
            _DD_GuiLog("[ERROR] " & $sSummary)

        Case "ECHEC_HTTP"
            $sSummary = "Impossible de contacter Dispatch en HTTP : " & $g_sLastHttpError
            _DD_GuiSetStatus("Test '" & $sActionName & "' : echec de connexion HTTP.")
            _DD_GuiLog("[ERROR] " & $sSummary)

        Case Else
            $sSummary = "Resultat : " & $sOutcome
    EndSwitch

    _DD_WriteActionTestReport($sActionName, $sField1, $sParam1, $sField2, $sParam2, $sOutcome, $sSummary, $sOutput)
EndFunc   ;==>_DD_FinishActionTest

Func _DD_WriteActionTestReport($sActionName, $sField1, $sParam1, $sField2, $sParam2, $sOutcome, $sSummary, $sRawOutput)
    Local $sReportsDir = @ScriptDir & "\reports"
    If Not FileExists($sReportsDir) Then DirCreate($sReportsDir)

    Local $sTimestamp = @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC
    Local $sPath = $sReportsDir & "\DispatchDoctor-test-" & $sActionName & "-" & $sTimestamp & ".txt"

    Local $hFile = FileOpen($sPath, 2)
    If $hFile = -1 Then Return

    FileWriteLine($hFile, "========================================================")
    FileWriteLine($hFile, $DD_TOOLNAME & " - Rapport de test d'action reelle")
    FileWriteLine($hFile, "========================================================")
    FileWriteLine($hFile, "Date       : " & _DD_NowString())
    FileWriteLine($hFile, "Action     : " & $sActionName)
    FileWriteLine($hFile, $sField1 & " : " & $sParam1)
    If $sField2 <> "" Then FileWriteLine($hFile, $sField2 & " : " & $sParam2)
    FileWriteLine($hFile, "Resultat   : " & $sOutcome)
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "--- Resume ---")
    FileWriteLine($hFile, $sSummary)
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "--- Sortie brute capturee (stdout/stderr, /ErrorStdOut) ---")
    FileWriteLine($hFile, $sRawOutput)

    FileClose($hFile)

    _DD_GuiLog("[*] Rapport de test ecrit : " & $sPath)
    ConsoleWrite("[*] Rapport de test ecrit : " & $sPath & @CRLF)
EndFunc   ;==>_DD_WriteActionTestReport

#endregion

; ============================================================================
#region Interface graphique (fenetre de progression)
; ============================================================================

; Fenetre de progression optionnelle, dans le meme style visuel que
; MainDispatch.au3 (fond sombre, Segoe UI). Toutes les fonctions _DD_Gui*
; sont sans effet tant que $g_bGuiActive est a False, ce qui permet au
; mode "selftest" de rester 100% console (rapide, non interactif). On
; evite deliberement toute dependance a GUIConstantsEx.au3 : le code de
; fermeture de fenetre -3 et les styles par defaut de GUICtrlCreateEdit
; suffisent, comme le fait deja MainDispatch.au3 (If $iMsg = -3 Then Exit).
Func _DD_GuiInit()
    $g_hGui = GUICreate($DD_TOOLNAME & " v" & $DD_VERSION & "  -  " & $g_sMode, 640, 700)
    GUISetBkColor(0x1E1E1E, $g_hGui)
    GUISetFont(10, 400, 0, "Segoe UI", $g_hGui)

    Local $idTitle = GUICtrlCreateLabel($DD_TOOLNAME & " - Diagnostic Dispatch", 20, 15, 600, 24)
    GUICtrlSetFont($idTitle, 13, 700, 0, "Segoe UI")
    GUICtrlSetColor($idTitle, 0x66CCFF)
    GUICtrlSetBkColor($idTitle, 0x1E1E1E)

    $g_idLblStatus = GUICtrlCreateLabel("Initialisation...", 20, 48, 600, 20)
    GUICtrlSetColor($g_idLblStatus, 0x00FF88)
    GUICtrlSetBkColor($g_idLblStatus, 0x1E1E1E)

    $g_idLblStats = GUICtrlCreateLabel("", 20, 72, 600, 20)
    GUICtrlSetColor($g_idLblStats, 0xCCCCCC)
    GUICtrlSetBkColor($g_idLblStats, 0x1E1E1E)

    ; Styles par defaut de GUICtrlCreateEdit (multiline + ascenseur vertical
    ; automatique) : suffisants pour un journal en lecture visuelle, pas
    ; besoin de deviner des constantes de style non verifiees ici.
    $g_idEditLog = GUICtrlCreateEdit("", 20, 98, 600, 170)
    GUICtrlSetFont($g_idEditLog, 9, 400, 0, "Consolas")
    GUICtrlSetColor($g_idEditLog, 0xDDDDDD)
    GUICtrlSetBkColor($g_idEditLog, 0x111111)

    Local $idBrowseLabel = GUICtrlCreateLabel("Parcourir les fonctions detectees :", 20, 276, 400, 18)
    GUICtrlSetColor($idBrowseLabel, 0x66CCFF)
    GUICtrlSetBkColor($idBrowseLabel, 0x1E1E1E)

    ; Combo : liste des domaines (peuplee une fois l'indexation terminee).
    ; Style par defaut de GUICtrlCreateCombo (liste deroulante simple),
    ; suffisant ici, pas besoin de style explicite. Le filtrage se fait
    ; via le bouton "Filtrer" (et non uniquement au changement de
    ; selection dans le combo) pour rester fiable quelle que soit la
    ; maniere dont l'utilisateur valide son choix (clic, clavier, ...).
    $g_idComboDomain = GUICtrlCreateCombo("", 20, 298, 115, 24)
    $g_idBtnFilterDomain = GUICtrlCreateButton("Filtrer", 140, 297, 65, 26)

    $g_idInputSearch = GUICtrlCreateInput("", 210, 298, 205, 24)
    $g_idBtnSearch = GUICtrlCreateButton("Chercher", 420, 297, 70, 26)
    $g_idBtnOpenFile = GUICtrlCreateButton("Ouvrir fichier", 495, 297, 125, 26)

    ; Style par defaut de GUICtrlCreateList (liste simple a selection unique) :
    ; suffisant pour parcourir des noms de fonctions, pas de style a deviner.
    $g_idListFunc = GUICtrlCreateList("", 20, 328, 600, 250)
    GUICtrlSetFont($g_idListFunc, 9, 400, 0, "Consolas")
    GUICtrlSetColor($g_idListFunc, 0xDDDDDD)
    GUICtrlSetBkColor($g_idListFunc, 0x111111)

    Local $idTestLabel = GUICtrlCreateLabel("Tester une action reelle (sur une instance de Dispatch dediee au test) : action, param. principal, param. secondaire", 20, 586, 600, 18)
    GUICtrlSetColor($idTestLabel, 0x66CCFF)
    GUICtrlSetBkColor($idTestLabel, 0x1E1E1E)

    $g_idComboTestAction = GUICtrlCreateCombo("", 20, 606, 290, 24)
    Local $t = 0
    Local $sActionData = ""
    For $t = 0 To UBound($g_aTestableActions, 1) - 1
        If $sActionData <> "" Then $sActionData &= "|"
        $sActionData &= $g_aTestableActions[$t][0]
    Next
    GUICtrlSetData($g_idComboTestAction, $sActionData, $g_aTestableActions[0][0])

    $g_idInputParam1 = GUICtrlCreateInput("", 315, 606, 100, 26)
    $g_idInputParam2 = GUICtrlCreateInput("", 420, 606, 80, 26)
    $g_idBtnRunTest = GUICtrlCreateButton("Lancer le test", 505, 605, 115, 28)

    ; $g_idBtnOpenHtml est cree plus tard (une fois les rapports generes),
    ; a cet emplacement fixe reserve pour lui.
    $g_idBtnClose = GUICtrlCreateButton("Fermer", 480, 650, 140, 28)

    GUISetState(@SW_SHOW, $g_hGui)
    $g_bGuiActive = True
    _DD_GuiPump()
EndFunc   ;==>_DD_GuiInit

; Peuple le menu deroulant des domaines (COMMON, CMR, COMAT, ETMS, ...)
; et affiche la liste des fonctions correspondantes. A appeler une fois
; que _DD_IndexFunctions() a tourne.
Func _DD_GuiPopulateBrowser()
    If Not $g_bGuiActive Then Return

    Local $oDomains = ObjCreate("Scripting.Dictionary")
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        Local $k = $g_aFunctions[$i][$UCOL_DOMAIN]
        If Not $oDomains.Exists($k) Then $oDomains.Add($k, True)
    Next

    Local $aKeys = $oDomains.Keys()
    Local $n = UBound($aKeys)
    Local $a = 0
    Local $b = 0
    For $a = 0 To $n - 2
        For $b = 0 To $n - 2 - $a
            If $aKeys[$b] > $aKeys[$b + 1] Then
                Local $sTmp = $aKeys[$b]
                $aKeys[$b] = $aKeys[$b + 1]
                $aKeys[$b + 1] = $sTmp
            EndIf
        Next
    Next

    Local $sData = "TOUS LES DOMAINES"
    For $i = 0 To $n - 1
        $sData &= "|" & $aKeys[$i]
    Next
    GUICtrlSetData($g_idComboDomain, $sData, "TOUS LES DOMAINES")

    _DD_GuiShowFunctionsForDomain("TOUS LES DOMAINES")
EndFunc   ;==>_DD_GuiPopulateBrowser

; Construit une ligne d'affichage pour une fonction, avec un separateur
; "::" facile a re-decouper (utilise par le bouton "Ouvrir fichier").
; Le caractere "|" est neutralise car il sert de separateur entre entrees
; dans les controles List d'AutoIt.
Func _DD_FormatFuncEntry($iRow, $bShowDomain = True)
    Local $sEntry = $g_aFunctions[$iRow][$UCOL_NAME] & "(" & $g_aFunctions[$iRow][$UCOL_PARAMS] & ")  ->  " & _
        _DD_RelativePath($g_aFunctions[$iRow][$UCOL_FILE], $g_sProjectRoot) & "::" & $g_aFunctions[$iRow][$UCOL_LINE]
    If $bShowDomain Then $sEntry &= "::" & $g_aFunctions[$iRow][$UCOL_DOMAIN]
    Return StringReplace($sEntry, "|", "/")
EndFunc   ;==>_DD_FormatFuncEntry

; Filtre la liste sur un domaine precis. Quand un seul domaine est
; affiche, son nom n'est plus repete sur chaque ligne (deja connu via le
; combo) pour garder l'affichage lisible dans la petite fenetre de liste.
Func _DD_GuiShowFunctionsForDomain($sDomain)
    If Not $g_bGuiActive Then Return
    Local $bAll = ($sDomain = "TOUS LES DOMAINES" Or $sDomain = "")
    Local $sData = ""
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        If Not $bAll And $g_aFunctions[$i][$UCOL_DOMAIN] <> $sDomain Then ContinueLoop
        If $sData <> "" Then $sData &= "|"
        $sData &= _DD_FormatFuncEntry($i, $bAll)
    Next
    If $sData = "" Then $sData = "(aucune fonction dans ce domaine)"
    GUICtrlSetData($g_idListFunc, $sData)
EndFunc   ;==>_DD_GuiShowFunctionsForDomain

Func _DD_GuiSearchFunctions()
    If Not $g_bGuiActive Then Return
    Local $sQuery = StringLower(StringStripWS(GUICtrlRead($g_idInputSearch), BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING)))
    If $sQuery = "" Then
        _DD_GuiShowFunctionsForDomain(GUICtrlRead($g_idComboDomain))
        Return
    EndIf

    Local $sData = ""
    Local $i = 0
    For $i = 0 To $g_iFuncCount - 1
        If StringInStr(StringLower($g_aFunctions[$i][$UCOL_NAME]), $sQuery) = 0 Then ContinueLoop
        If $sData <> "" Then $sData &= "|"
        $sData &= _DD_FormatFuncEntry($i, True)
    Next
    If $sData = "" Then $sData = "(aucun resultat)"
    GUICtrlSetData($g_idListFunc, $sData)
EndFunc   ;==>_DD_GuiSearchFunctions

; Ouvre, dans l'application associee (SciTE en general pour .au3), le
; fichier de la fonction actuellement selectionnee dans la liste.
Func _DD_GuiOpenSelectedFile()
    If Not $g_bGuiActive Then Return
    Local $sSelected = GUICtrlRead($g_idListFunc)
    If $sSelected = "" Or StringLeft($sSelected, 1) = "(" Then Return

    Local $aParts = StringSplit($sSelected, "::", BitOR($STR_NOCOUNT, $STR_ENTIRESPLIT))
    If UBound($aParts) < 2 Then Return

    Local $sRelPath = StringStripWS($aParts[1], BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING))
    Local $sFullPath = $g_sProjectRoot & "\" & $sRelPath
    If FileExists($sFullPath) Then ShellExecute($sFullPath)
EndFunc   ;==>_DD_GuiOpenSelectedFile

Func _DD_GuiSetStatus($sText)
    If Not $g_bGuiActive Then Return
    GUICtrlSetData($g_idLblStatus, $sText)
    ConsoleWrite("[*] " & $sText & @CRLF)
    _DD_GuiLog("[*] " & $sText)
EndFunc   ;==>_DD_GuiSetStatus

Func _DD_GuiRefreshStats()
    If Not $g_bGuiActive Then Return
    Local $sStats = "Fichiers: " & $g_iFileCount & "   Fonctions: " & $g_iFuncCount & "   Variables: " & $g_iVariableCount & _
        "   Erreurs: " & $g_iCountError & "   Avertissements: " & $g_iCountWarning & "   Infos: " & $g_iCountInfo
    GUICtrlSetData($g_idLblStats, $sStats)
    _DD_GuiPump()
EndFunc   ;==>_DD_GuiRefreshStats

; Ajoute une ligne au journal visible dans la fenetre. Le tampon est borne
; pour rester fluide meme si le projet grandit beaucoup (des centaines de
; diagnostics) : on ne garde que les ~50 Ko les plus recents.
Func _DD_GuiLog($sLine)
    If Not $g_bGuiActive Then Return
    $g_sGuiLogBuffer &= $sLine & @CRLF
    If StringLen($g_sGuiLogBuffer) > 60000 Then
        $g_sGuiLogBuffer = "[... journal tronque ...]" & @CRLF & StringRight($g_sGuiLogBuffer, 50000)
    EndIf
    GUICtrlSetData($g_idEditLog, $g_sGuiLogBuffer)
    _DD_GuiPump()
EndFunc   ;==>_DD_GuiLog

; Vide la file de messages Windows pour garder la fenetre reactive
; (deplacement, redessin, clic sur Fermer) pendant les phases d'analyse.
; Gere aussi la fermeture anticipee par l'utilisateur (croix ou bouton
; Fermer) sans jamais planter : le code de sortie reste base sur les
; erreurs deja detectees a cet instant.
Func _DD_GuiPump()
    If Not $g_bGuiActive Then Return
    Local $iMsg = GUIGetMsg()
    If $iMsg = -3 Then
        _DD_GuiCloseAndExit()
    ElseIf $iMsg = $g_idBtnClose Then
        _DD_GuiCloseAndExit()
    ElseIf $g_idBtnOpenHtml <> 0 And $iMsg = $g_idBtnOpenHtml Then
        If $g_sHtmlReportPath <> "" And FileExists($g_sHtmlReportPath) Then ShellExecute($g_sHtmlReportPath)
    ElseIf $iMsg = $g_idComboDomain Then
        _DD_GuiShowFunctionsForDomain(GUICtrlRead($g_idComboDomain))
    ElseIf $iMsg = $g_idBtnFilterDomain Then
        _DD_GuiShowFunctionsForDomain(GUICtrlRead($g_idComboDomain))
    ElseIf $iMsg = $g_idBtnSearch Then
        _DD_GuiSearchFunctions()
    ElseIf $iMsg = $g_idBtnOpenFile Then
        _DD_GuiOpenSelectedFile()
    ElseIf $iMsg = $g_idBtnRunTest Then
        _DD_RunActionTest(GUICtrlRead($g_idComboTestAction), GUICtrlRead($g_idInputParam1), GUICtrlRead($g_idInputParam2))
    EndIf
EndFunc   ;==>_DD_GuiPump

Func _DD_GuiCloseAndExit()
    Local $iCode = 0
    If $g_iCountError > 0 Then $iCode = 1
    ; Evite de laisser une instance de test orpheline si la fenetre est
    ; fermee pendant un test d'action en cours.
    If $g_iActionTestPid <> 0 And ProcessExists($g_iActionTestPid) Then ProcessClose($g_iActionTestPid)
    If $g_hGui <> 0 Then GUIDelete($g_hGui)
    Exit $iCode
EndFunc   ;==>_DD_GuiCloseAndExit

; Affiche le bouton "Ouvrir le rapport HTML" (si un rapport a ete genere)
; et garde la fenetre ouverte jusqu'a ce que l'utilisateur la ferme, pour
; qu'il ait le temps de lire le journal et les statistiques finales.
Func _DD_GuiWaitClose()
    If Not $g_bGuiActive Then Return

    If $g_sHtmlReportPath <> "" Then
        $g_idBtnOpenHtml = GUICtrlCreateButton("Ouvrir le rapport HTML", 20, 650, 220, 28)
    EndIf

    _DD_GuiSetStatus("Termine. Vous pouvez consulter le journal ci-dessus ou fermer cette fenetre.")

    While 1
        _DD_GuiPump()
        Sleep(50)
    WEnd
EndFunc   ;==>_DD_GuiWaitClose

#endregion

; ============================================================================
#region Utilitaires
; ============================================================================

Func _DD_GrowTable(ByRef $aArr, $iCount)
    If $iCount >= UBound($aArr, 1) Then
        Local $iCols = UBound($aArr, 2)
        Local $iNewSize = UBound($aArr, 1) * 2
        ReDim $aArr[$iNewSize][$iCols]
    EndIf
EndFunc   ;==>_DD_GrowTable

Func _DD_AddFileRow($sPath, $sRel, $sParent, $iLine, $iDepth, $sDomain)
    _DD_GrowTable($g_aFiles, $g_iFileCount)
    $g_aFiles[$g_iFileCount][$FCOL_PATH] = $sPath
    $g_aFiles[$g_iFileCount][$FCOL_RELPATH] = $sRel
    $g_aFiles[$g_iFileCount][$FCOL_PARENT] = $sParent
    $g_aFiles[$g_iFileCount][$FCOL_LINE] = $iLine
    $g_aFiles[$g_iFileCount][$FCOL_DEPTH] = $iDepth
    $g_aFiles[$g_iFileCount][$FCOL_DOMAIN] = $sDomain
    $g_iFileCount += 1
EndFunc   ;==>_DD_AddFileRow

Func _DD_AddFuncRow($sName, $sFile, $iLine, $iEndLine, $sParams, $iParamCount, $sDomain)
    _DD_GrowTable($g_aFunctions, $g_iFuncCount)
    $g_aFunctions[$g_iFuncCount][$UCOL_NAME] = $sName
    $g_aFunctions[$g_iFuncCount][$UCOL_FILE] = $sFile
    $g_aFunctions[$g_iFuncCount][$UCOL_LINE] = $iLine
    $g_aFunctions[$g_iFuncCount][$UCOL_ENDLINE] = $iEndLine
    $g_aFunctions[$g_iFuncCount][$UCOL_PARAMS] = $sParams
    $g_aFunctions[$g_iFuncCount][$UCOL_PARAMCOUNT] = $iParamCount
    $g_aFunctions[$g_iFuncCount][$UCOL_DOMAIN] = $sDomain
    $g_iFuncCount += 1
EndFunc   ;==>_DD_AddFuncRow

Func _DD_AddDiagnostic($sSeverity, $sCode, $sDomain, $sFile, $iLine, $sFunction, $sSymbol, $sMessage, $sSuggestion, $sConfidence)
    _DD_GrowTable($g_aDiagnostics, $g_iDiagCount)
    $g_aDiagnostics[$g_iDiagCount][$DCOL_SEVERITY] = $sSeverity
    $g_aDiagnostics[$g_iDiagCount][$DCOL_CODE] = $sCode
    $g_aDiagnostics[$g_iDiagCount][$DCOL_DOMAIN] = $sDomain
    $g_aDiagnostics[$g_iDiagCount][$DCOL_FILE] = $sFile
    $g_aDiagnostics[$g_iDiagCount][$DCOL_LINE] = $iLine
    $g_aDiagnostics[$g_iDiagCount][$DCOL_FUNCTION] = $sFunction
    $g_aDiagnostics[$g_iDiagCount][$DCOL_SYMBOL] = $sSymbol
    $g_aDiagnostics[$g_iDiagCount][$DCOL_MESSAGE] = $sMessage
    $g_aDiagnostics[$g_iDiagCount][$DCOL_SUGGESTION] = $sSuggestion
    $g_aDiagnostics[$g_iDiagCount][$DCOL_CONFIDENCE] = $sConfidence
    $g_iDiagCount += 1

    Switch $sSeverity
        Case $DD_SEV_ERROR
            $g_iCountError += 1
        Case $DD_SEV_WARNING
            $g_iCountWarning += 1
        Case $DD_SEV_INFO
            $g_iCountInfo += 1
        Case Else
            $g_iCountOk += 1
    EndSwitch

    Local $sLineInfo = ""
    If $iLine > 0 Then $sLineInfo = ":" & $iLine
    Local $sLogLine = "[" & $sSeverity & "] " & $sCode & " - " & _DD_RelativePath($sFile, $g_sProjectRoot) & $sLineInfo & " - " & $sMessage
    ConsoleWrite($sLogLine & @CRLF)
    _DD_GuiLog($sLogLine)
    _DD_GuiRefreshStats()
EndFunc   ;==>_DD_AddDiagnostic

Func _DD_HasDiagnosticCode($sCode)
    Local $i = 0
    For $i = 0 To $g_iDiagCount - 1
        If $g_aDiagnostics[$i][$DCOL_CODE] = $sCode Then Return True
    Next
    Return False
EndFunc   ;==>_DD_HasDiagnosticCode

Func _DD_NormalizePath($sPath)
    Local $s = StringReplace($sPath, "/", "\")
    Local $bUNC = (StringLeft($s, 2) = "\\")
    $s = StringRegExpReplace($s, "\\{2,}", "\")
    If $bUNC And StringLeft($s, 2) <> "\\" Then $s = "\" & $s
    $s = StringRegExpReplace($s, "\\$", "")
    Return $s
EndFunc   ;==>_DD_NormalizePath

Func _DD_ParentDir($sPath)
    Local $iPos = StringInStr($sPath, "\", 0, -1)
    If $iPos = 0 Then Return $sPath
    Return StringLeft($sPath, $iPos - 1)
EndFunc   ;==>_DD_ParentDir

Func _DD_RelativePath($sPath, $sRoot)
    If StringLeft(StringLower($sPath), StringLen($sRoot)) = StringLower($sRoot) Then
        Local $sRel = StringMid($sPath, StringLen($sRoot) + 1)
        $sRel = StringRegExpReplace($sRel, "^\\+", "")
        If $sRel = "" Then Return $sPath
        Return $sRel
    EndIf
    Return $sPath
EndFunc   ;==>_DD_RelativePath

Func _DD_NowString()
    Return @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc   ;==>_DD_NowString

; Distance de Levenshtein classique (programmation dynamique, deux lignes).
; Utilisee pour detecter les fautes de frappe entre noms de fonctions et
; suggerer des variables proches pour les diagnostics de variables.
Func _DD_Levenshtein($s1, $s2)
    Local $iLen1 = StringLen($s1)
    Local $iLen2 = StringLen($s2)
    If $iLen1 = 0 Then Return $iLen2
    If $iLen2 = 0 Then Return $iLen1

    Local $aPrev[$iLen2 + 1]
    Local $aCurr[$iLen2 + 1]
    Local $j = 0
    For $j = 0 To $iLen2
        $aPrev[$j] = $j
    Next

    Local $i = 0
    For $i = 1 To $iLen1
        $aCurr[0] = $i
        Local $cA = StringMid($s1, $i, 1)
        For $j = 1 To $iLen2
            Local $cB = StringMid($s2, $j, 1)
            Local $iCost = 1
            If $cA = $cB Then $iCost = 0

            Local $iDel = $aPrev[$j] + 1
            Local $iIns = $aCurr[$j - 1] + 1
            Local $iSub = $aPrev[$j - 1] + $iCost

            Local $iMin = $iDel
            If $iIns < $iMin Then $iMin = $iIns
            If $iSub < $iMin Then $iMin = $iSub

            $aCurr[$j] = $iMin
        Next

        For $j = 0 To $iLen2
            $aPrev[$j] = $aCurr[$j]
        Next
    Next

    Return $aPrev[$iLen2]
EndFunc   ;==>_DD_Levenshtein

Func _DD_ResetState()
    Global $g_aFiles[8][$FCOL_COLS]
    $g_iFileCount = 0
    Global $g_aFunctions[8][$UCOL_COLS]
    $g_iFuncCount = 0
    Global $g_aDiagnostics[8][$DCOL_COLS]
    $g_iDiagCount = 0
    $g_iCountError = 0
    $g_iCountWarning = 0
    $g_iCountInfo = 0
    $g_iCountOk = 0
    $g_iIncludeEdgeCount = 0
    $g_iVariableCount = 0
    $oDD_FileTextCache = ObjCreate("Scripting.Dictionary")
    $oDD_VisitedFiles = ObjCreate("Scripting.Dictionary")
    $oDD_IncludeCount = ObjCreate("Scripting.Dictionary")
    $oDD_IncludedBy = ObjCreate("Scripting.Dictionary")
    $oDD_Ancestors = ObjCreate("Scripting.Dictionary")
    $oDD_GlobalVars = ObjCreate("Scripting.Dictionary")
    $oDD_AllProjectFiles = ObjCreate("Scripting.Dictionary")
EndFunc   ;==>_DD_ResetState

#endregion

; ============================================================================
#region Auto-tests internes (--mode selftest)
; ============================================================================

; Cree de petits fichiers .au3 temporaires dans @TempDir, verifie que
; chaque cas attendu par le cahier des charges est bien detecte (ou bien
; absent quand il ne doit pas y avoir de faux positif), puis nettoie tout.
; Ne touche jamais aux fichiers du projet Dispatch.
Func _DD_RunSelfTests()
    ConsoleWrite("=== Auto-tests internes " & $DD_TOOLNAME & " ===" & @CRLF)

    Local $sTestDir = @TempDir & "\DispatchDoctorSelfTest_" & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC
    DirCreate($sTestDir)

    Local $iPass = 0
    Local $iFail = 0

    ; Cas 1 : fonction valide + variable Local valide + variable Global valide
    Local $sMain1 = $sTestDir & "\Main1.au3"
    FileWriteLine($sMain1, '#include "Sub1.au3"')
    FileWriteLine($sMain1, "Global $g_iValue = 0")
    FileWriteLine($sMain1, "Func _Test_Valid($sParam)")
    FileWriteLine($sMain1, "    Local $sLocal = $sParam")
    FileWriteLine($sMain1, "    $g_iValue = $g_iValue + 1")
    FileWriteLine($sMain1, "    Return $sLocal")
    FileWriteLine($sMain1, "EndFunc")
    FileWriteLine($sTestDir & "\Sub1.au3", "; fichier vide intentionnellement")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain1)
    _DD_IndexFunctions()
    _DD_IndexGlobalVariables()
    _DD_AnalyzeAllVariables()
    _DD_SelfTestReport("Cas 1 : fonction/variables valides", $g_iFuncCount = 1 And $g_iCountError = 0, $iPass, $iFail)

    ; Cas 2 : variable non declaree
    Local $sMain2 = $sTestDir & "\Main2.au3"
    FileWriteLine($sMain2, "Func _Test_Undeclared()")
    FileWriteLine($sMain2, "    Local $a = 1")
    FileWriteLine($sMain2, "    Return $a + $bNotDeclaredXyz")
    FileWriteLine($sMain2, "EndFunc")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain2)
    _DD_IndexFunctions()
    _DD_IndexGlobalVariables()
    _DD_AnalyzeAllVariables()
    _DD_SelfTestReport("Cas 2 : variable non declaree", _DD_HasDiagnosticCode($DD_CODE_UNDECLARED_VARIABLE), $iPass, $iFail)

    ; Cas 3 : fonction dupliquee
    Local $sMain3 = $sTestDir & "\Main3.au3"
    FileWriteLine($sMain3, '#include "Dup_A.au3"')
    FileWriteLine($sMain3, '#include "Dup_B.au3"')
    FileWriteLine($sTestDir & "\Dup_A.au3", "Func _Test_Dup()" & @CRLF & "    Return 1" & @CRLF & "EndFunc")
    FileWriteLine($sTestDir & "\Dup_B.au3", "Func _Test_Dup()" & @CRLF & "    Return 2" & @CRLF & "EndFunc")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain3)
    _DD_IndexFunctions()
    _DD_SelfTestReport("Cas 3 : fonction dupliquee", _DD_HasDiagnosticCode($DD_CODE_DUPLICATE_FUNCTION), $iPass, $iFail)

    ; Cas 4 : include manquant
    Local $sMain4 = $sTestDir & "\Main4.au3"
    FileWriteLine($sMain4, '#include "DoesNotExist_Xyz.au3"')

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain4)
    _DD_SelfTestReport("Cas 4 : include manquant", _DD_HasDiagnosticCode($DD_CODE_MISSING_INCLUDE), $iPass, $iFail)

    ; Cas 5 : include circulaire (verifie aussi l'absence de boucle infinie :
    ; si ce test se termine, c'est deja la preuve que le parcours est borne)
    Local $sCircA = $sTestDir & "\Circ_A.au3"
    Local $sCircB = $sTestDir & "\Circ_B.au3"
    FileWriteLine($sCircA, '#include "Circ_B.au3"')
    FileWriteLine($sCircB, '#include "Circ_A.au3"')

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sCircA)
    _DD_SelfTestReport("Cas 5 : inclusion circulaire (sans boucle infinie)", _DD_HasDiagnosticCode($DD_CODE_CIRCULAR_INCLUDE), $iPass, $iFail)

    ; Cas 6 : reconnaissance de "Unknown function name" dans la sortie compilateur
    _DD_ResetState()
    _DD_ParseCompilerOutput('"' & $sMain2 & '" (3) : ==> Unknown function name.:' & @CRLF & _
        "_FonctionInconnueXyz()" & @CRLF & "~~~~~~~~~~~~~~~~~~~~~~^" & @CRLF)
    _DD_SelfTestReport("Cas 6 : sortie compilateur 'Unknown function name'", _DD_HasDiagnosticCode($DD_CODE_UNKNOWN_FUNCTION), $iPass, $iFail)

    ; Cas 7 : reconnaissance de "Incorrect number of parameters"
    _DD_ResetState()
    _DD_ParseCompilerOutput('"' & $sMain2 & '" (5) : ==> Incorrect number of parameters in function call.:' & @CRLF)
    _DD_SelfTestReport("Cas 7 : sortie compilateur 'Incorrect number of parameters'", _DD_HasDiagnosticCode($DD_CODE_BAD_PARAM_COUNT), $iPass, $iFail)

    ; Cas 8 : variables Local valides, pas de faux positif
    Local $sMain8 = $sTestDir & "\Main8.au3"
    FileWriteLine($sMain8, "Func _Test_LocalOnly()")
    FileWriteLine($sMain8, "    Local $x = 1, $y = 2")
    FileWriteLine($sMain8, "    Local $z")
    FileWriteLine($sMain8, "    $z = $x + $y")
    FileWriteLine($sMain8, "    Return $z")
    FileWriteLine($sMain8, "EndFunc")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain8)
    _DD_IndexFunctions()
    _DD_IndexGlobalVariables()
    _DD_AnalyzeAllVariables()
    _DD_SelfTestReport("Cas 8 : variables Local valides (pas de faux positif)", Not _DD_HasDiagnosticCode($DD_CODE_UNDECLARED_VARIABLE), $iPass, $iFail)

    ; Cas 9 : variable Global valide utilisee dans une autre fonction
    Local $sMain9 = $sTestDir & "\Main9.au3"
    FileWriteLine($sMain9, "Global $g_bReady = False")
    FileWriteLine($sMain9, "Func _Test_SetReady()")
    FileWriteLine($sMain9, "    $g_bReady = True")
    FileWriteLine($sMain9, "EndFunc")
    FileWriteLine($sMain9, "Func _Test_IsReady()")
    FileWriteLine($sMain9, "    Return $g_bReady")
    FileWriteLine($sMain9, "EndFunc")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain9)
    _DD_IndexFunctions()
    _DD_IndexGlobalVariables()
    _DD_AnalyzeAllVariables()
    _DD_SelfTestReport("Cas 9 : variable Global valide entre fonctions", Not _DD_HasDiagnosticCode($DD_CODE_UNDECLARED_VARIABLE), $iPass, $iFail)

    ; Cas 10 : ReDim en toute premiere reference (piege reel decouvert sur
    ; _QueueClearRows / generation CMR) doit etre detecte
    Local $sMain10 = $sTestDir & "\Main10.au3"
    FileWriteLine($sMain10, "Global $g_aTest10[0]")
    FileWriteLine($sMain10, "Func _Test_ReDimUnsafe()")
    FileWriteLine($sMain10, "    ReDim $g_aTest10[5]")
    FileWriteLine($sMain10, "EndFunc")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain10)
    _DD_IndexFunctions()
    _DD_IndexGlobalVariables()
    _DD_AnalyzeAllVariables()
    _DD_SelfTestReport("Cas 10 : ReDim en premiere reference detecte (piege _QueueClearRows)", _DD_HasDiagnosticCode($DD_CODE_REDIM_UNSAFE), $iPass, $iFail)

    ; Cas 11 : ReDim precede d'une lecture (UBound) ne doit jamais etre signale
    Local $sMain11 = $sTestDir & "\Main11.au3"
    FileWriteLine($sMain11, "Global $g_aTest11[0]")
    FileWriteLine($sMain11, "Func _Test_ReDimSafe()")
    FileWriteLine($sMain11, "    Local $n = UBound($g_aTest11)")
    FileWriteLine($sMain11, "    ReDim $g_aTest11[$n + 1]")
    FileWriteLine($sMain11, "EndFunc")

    _DD_ResetState()
    $g_sProjectRoot = $sTestDir
    _DD_BuildIncludeGraph($sMain11)
    _DD_IndexFunctions()
    _DD_IndexGlobalVariables()
    _DD_AnalyzeAllVariables()
    _DD_SelfTestReport("Cas 11 : ReDim precede d'une lecture (pas de faux positif)", Not _DD_HasDiagnosticCode($DD_CODE_REDIM_UNSAFE), $iPass, $iFail)

    DirRemove($sTestDir, 1)

    ConsoleWrite(@CRLF & "Auto-tests : " & $iPass & " reussi(s), " & $iFail & " echoue(s)." & @CRLF)
    If $iFail > 0 Then Exit 1
EndFunc   ;==>_DD_RunSelfTests

Func _DD_SelfTestReport($sLabel, $bCondition, ByRef $iPass, ByRef $iFail)
    If $bCondition Then
        ConsoleWrite("[PASS] " & $sLabel & @CRLF)
        $iPass += 1
    Else
        ConsoleWrite("[FAIL] " & $sLabel & @CRLF)
        $iFail += 1
    EndIf
EndFunc   ;==>_DD_SelfTestReport

#endregion

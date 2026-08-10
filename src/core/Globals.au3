; ============================================================================
; Globals.au3
; Includes, variables globales, bootstrap serveur et boucle principale extraits du monolithe.
; -----------------------------------------------------------------------------
; Extrait automatiquement depuis Dispatch.txt.
; MainDispatch.au3 reste la version executable complete pour eviter toute casse.
; ============================================================================

; ╔══════════════════════════════════════════════════════════════════════════╗
; ║  DispatchMaster 2.0 - Serveur Local & Hub AutoIt                         ║
; ╚══════════════════════════════════════════════════════════════════════════╝
#include <File.au3>
#include <String.au3>
#include <Date.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <ListBoxConstants.au3>
#include <ComboConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>

; ═══════════════════════════════════════════════════════════════════════════
; VARIABLES GLOBALES
; ═══════════════════════════════════════════════════════════════════════════
Global $g_iTrackCount   = 0
Global $g_aTrackIDs[1]
Global $g_hTracker      = 0
Global $g_idTrackProg = 0
Global $g_idTrackLbl  = 0
Global $g_idTrackLV   = 0
Global $bFC_Stop        = False
Global $bFC_Pause       = False
Global $bFC_Skip        = False
Global $iFC_StepCurrent = 0
Global $g_sFC_AuditLog  = ""
Global $g_bFC_Audit     = True
Global $bCOMAT_Stop  = False
Global $bCOMAT_Pause = False
Global $bCOMAT_Skip  = False

; GUI Batch control buttons
Global $g_idBtnPause  = 0
Global $g_idBtnPlay   = 0
Global $g_idBtnSkip   = 0
Global $g_idBtnStop   = 0
Global $g_idBatchInfo = 0

Global $sClassFileOpen  = "[CLASS:#32770; TITLE:Open]"
Global $sClassMenu      = "[CLASS:TfmMenuSelection]"
Global $sClassInput     = "[CLASS:TfmInput]"
Global $idToolbar       = "[CLASS:TToolBar; INSTANCE:1]"
Global $DELAY_MEDIUM    = 500
Global $DELAY_LONG      = 1000
Global $DELAY_ETMS_LOAD = 1500

Global Const $COMAT_LOG_CTRL   = "[CLASS:TEIEdit; INSTANCE:91]"
Global Const $COMAT_DELAY_S    = 80
Global Const $COMAT_DELAY_M    = 150
Global Const $COMAT_DELAY_L    = 300
Global Const $COMAT_DELAY_LOAD = 1500

; ── CONSTANTES DE SÉCURITÉ (dispatch_patch) ──
Global Const $MAX_BODY_SIZE       = 2097152 ; 2 MB max par requête
Global Const $MAX_PATH_LENGTH     = 500     ; Longueur max d'un chemin réseau
Global Const $ALLOWED_NET_PREFIX  = "F:\"   ; Seul préfixe autorisé pour les chemins réseau
Global Const $ID_PATTERN          = "^[a-zA-Z0-9_\-\.\+\s]{1,200}$" ; Pattern valide pour un ID

; ── Checksum réseau (pour /api/net-check) ──
Global $g_sLastNetChecksum = ""
Global $g_sLastNetModified = ""

; ── Declarations globales dupliquees depuis d'autres fichiers (voir note ──
; ci-dessus): tout #include place APRES ce fichier dans MainDispatch.au3 a ses
; propres "Global $x = ..." de haut de fichier qui ne s'executent jamais tant
; que la boucle bloquante ci-dessous tourne -- exactement le bug qui a fait
; planter $gsConfigFile, $g_oOutlook/$g_oNamespace puis $HPE_INI_PATH, un a la
; fois, au fur et a mesure que chaque fonctionnalite etait utilisee pour la
; premiere fois. On les redeclare ici (sans Const, pour eviter tout conflit de
; redeclaration si la ligne d'origine finit par s'executer) pour que TOUTES les
; fonctions appelees depuis HttpServer_HandleClient les trouvent des le depart,
; au lieu de decouvrir les cas un par un a chaque nouveau crash utilisateur.

; (giPort/giMainSocket omis : deja declares Global en tete de MainDispatch.au3,
; qui s'execute avant l'#include de ce fichier -- pas besoin de doublon.)

; -- src/core/Constants.au3 --
Global $CFG_FILE = @ScriptDir & "\robot_v15_config.ini"

; -- src/services/StateService.au3 --
Global $APP_TITLE = "HPE BL Queue Snapshot"
Global $BASE_PATH = "F:\Scripting\Export\EXPORT_HPE_CMR_001\"
Global $INPUT_PATH = $BASE_PATH & "Database\INPUT\"
Global $OUTPUT_PATH = $BASE_PATH & "Database\OUTPUT\"
Global $HTML_PATH = $BASE_PATH & "Database\HTML\"
Global $PDF_PATH = @ScriptDir & "\PDF\"
Global $SNAPSHOT_PATH = @ScriptDir & "\Snapshots\"
Global $INPUT_CSV = $INPUT_PATH & "input.csv"
Global $INPUT_GEN = $INPUT_PATH & "inputGEN.csv"
Global $LASTBL_MARKER = $INPUT_PATH & "last_bl_queue_snapshot.ini"
Global $EDS_PATH = $BASE_PATH & "EXPORT_HPE_CMR_006.eds"
Global $INI_PATH = @ScriptDir & "\Config\HPE_BL_QUEUE_SNAPSHOT.ini"
Global $MAIL_AUTOSEND = False
Global $OUTPUT_WAIT_MS = 180000
Global $OUTPUT_STABLE_MS = 1500
Global $ETMS_WINDOW = "[CLASS:TfmBrowser]"
Global $ETMS_TOOLBAR = "[CLASS:TRzToolbar; INSTANCE:1]"
Global $EDOC_UPLOAD_TITLE = "Upload Documents CDG"
Global $EDOC_UPLOAD_BTN = "TButton2"
Global $FILEOPEN_WIN = "[CLASS:TRzShellOpenSaveForm]"
Global $FILEOPEN_EDIT = "[CLASS:TRzEdit; INSTANCE:1]"
Global $TOOLBAR_X = 54
Global $TOOLBAR_Y = 9
Global $g_sOperator = ""
Global $g_bStopQueue = False
Global $g_aBL[0], $g_aHTML[0], $g_aPDF[0], $g_aCarrier[0], $g_aDelivery[0], $g_aCompany[0], $g_aSnap[0]
Global $g_aQueueBlocks[0], $g_aQueueStatus[0]
Global $idQueue = 0, $idQueueList = 0, $idStatus = 0, $idResults = 0, $idLastBL = 0
Global $idSPCarrier = 0, $idSPTo = 0, $idSPCC = 0, $idSPBCC = 0, $idSPSubject = 0, $idSPPdf = 0, $idSPBody = 0, $idSPSign = 0, $idSPKnown = 0
Global $g_hGui = 0
Global $g_bGenerating = False
Global $g_bForceNext = False
Global $g_bForceSkipCurrent = False
Global $g_bHardClose = False
Global $idBtnStopQueue = 0
Global $idBtnForceNext = 0
Global $idBtnMail = 0
Global $idBtnEdocOnly = 0
Global $idBtnEdocMail = 0
Global $idBtnRefreshResults = 0
Global $idBtnRefreshResultsTop = 0
Global $g_aHoverIds[0], $g_aHoverNormal[0], $g_aHoverHover[0], $g_aHoverIsHover[0]
Global $g_oCsvCache = ObjCreate("Scripting.Dictionary")
Global $g_sBrowserCache = ""
Global $g_iLastHoverTick = 0
Global $HOVER_THROTTLE_MS = 60

; -- src/features/cmr/CMR_EDOC.au3 --
Global $EDOC_APP_TITLE = "EDOC Master Bot V15.6"
Global $TEMP_DIR = @TempDir & "\Temp_EDOC_Robot_V15"
Global $EDOC_WM_NOTIFY = 0x004E, $EDOC_NM_DBLCLK = -3
Global $C_BG = 0xF3F4F6, $C_CARD = 0xFFFFFF, $C_BORDER = 0xCBD5E1, $C_TEXT = 0x111827, $C_MUTED = 0x6B7280, $C_ACCENT = 0xDCFCE7, $C_BUTTON = 0xE5E7EB
Global $g_oOutlook = 0, $g_oNamespace = 0, $g_sHistorique = "|"
Global $g_aDynCtrls[30][6], $g_iDynCount = 0
Global $g_aGrouped[800][15], $g_iGroupedCount = 0
Global $g_idSelList = 0, $g_hSelList = 0

; -- src/features/hpe/HPE_Tool.au3 --
Global $HPE_INI_PATH = @ScriptDir & "\Config\HPE_Tool.ini"

Opt("TrayIconDebug", 0)
Opt("TrayMenuMode", 3)        ; pas de menu Pause/Exit par défaut
Opt("TrayAutoPause", 0)       ; ne jamais auto-pauser le script
TCPStartup()

Global $g_iPort       = 9500
Global $g_iMainSocket = -1

; Fermer les instances en double
Local $iPID  = @AutoItPID
Local $aList = ProcessList(@ScriptName)
For $i = 1 To $aList[0][0]
    If $aList[$i][1] <> $iPID Then ProcessClose($aList[$i][1])
Next
Sleep(400)

; Essai port 9500, puis ports aléatoires si occupé
$g_iMainSocket = TCPListen("127.0.0.1", 9500)
If $g_iMainSocket <> -1 Then
    $g_iPort = 9500
Else
    Local $iFound = 0
    Local $iP = 0
    For $iP = 1 To 50
        Local $iTryPort = Random(10000, 60000, 1)
        $g_iMainSocket = TCPListen("127.0.0.1", $iTryPort)
        If $g_iMainSocket <> -1 Then
            $g_iPort = $iTryPort
            $iFound  = 1
            ExitLoop
        EndIf
    Next
    If $iFound = 0 Then
        MsgBox(16, "Erreur Serveur", "Impossible de démarrer le serveur TCP." & @CRLF & "Vérifiez votre pare-feu ou redémarrez le PC.")
        Exit
    EndIf
EndIf

Global $g_sSaveFile = @ScriptDir & "\historique_dispatch.json"
Global $g_sAuditLog = @ScriptDir & "\logs\dispatch_audit.log"
Global $g_iAuditCheckTimer = 0
Global Const $AUDIT_CHECK_INTERVAL = 60000 ; 60 secondes entre chaque vérification silencieuse
Global $g_sStatusFile   = @ScriptDir & "\dispatch_status.json"
Global $g_sDataFile     = @ScriptDir & "\dispatch_data.json"
Global $g_sContactsFile       = @ScriptDir & "\dispatch_contacts.tsv"
Global $g_sContactsLegacyFile = @ScriptDir & "\dispatch_contacts.json"
Global $g_sHtmlFile = @ScriptDir & "\interface.html"
; CMR_FULL_FUSION_V1 - moteur CMR fusionné dans le serveur Dispatch
Global $g_sCMRStatus = "Prêt."
Global $g_bCMRRunning = False
Global $g_sCMRLastJSON = "{}"
Global $g_sCMRQueueFile = @ScriptDir & "\cmr_queue.txt"

If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\Interface.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\Interface_v2.html"
If Not FileExists($g_sHtmlFile) Then $g_sHtmlFile = @ScriptDir & "\DispatchInterface.html"
If Not FileExists($g_sHtmlFile) Then
    MsgBox(16, "Erreur", "Fichier HTML introuvable dans : " & @ScriptDir & @CRLF & "Attendu : interface.html")
    Exit
EndIf
Global $g_sHTML = FileRead($g_sHtmlFile)
; Pré-calculer le HTML avec le bon port (une seule fois)
$g_sHTML = StringReplace($g_sHTML, "127.0.0.1:9500", "127.0.0.1:" & $g_iPort)
$g_sHTML = StringReplace($g_sHTML, "127.0.0.1:8080", "127.0.0.1:" & $g_iPort)

ShellExecute("http://127.0.0.1:" & $g_iPort)

; Créer les dossiers nécessaires
If Not FileExists(@ScriptDir & "\logs") Then DirCreate(@ScriptDir & "\logs")
If Not FileExists(@ScriptDir & "\backups") Then DirCreate(@ScriptDir & "\backups")
; \Config : stockage ini de plusieurs outils (HPE_Tool.au3, CMR_SP.au3...). Sans
; ce dossier, IniWrite echoue silencieusement (il ne cree pas les dossiers parents
; manquants) -- les liens crees via l'onglet HPE ne survivaient donc jamais a un
; redemarrage du serveur.
If Not FileExists(@ScriptDir & "\Config") Then DirCreate(@ScriptDir & "\Config")
_AuditLog("INFO", "Serveur démarré sur le port " & $g_iPort)
$g_iAuditCheckTimer = TimerInit()

; Boucle principale
While 1
    Local $iClientSocket = TCPAccept($g_iMainSocket)
  If $iClientSocket <> -1 Then HttpServer_HandleClient($iClientSocket)

    ; Vérification silencieuse en arrière-plan (chaque minute)
    If TimerDiff($g_iAuditCheckTimer) > $AUDIT_CHECK_INTERVAL Then
        _SilentHealthCheck()
        $g_iAuditCheckTimer = TimerInit()
    EndIf

    Sleep(10)
WEnd

; ==============================================================================
; SERVEUR HTTP
; ==============================================================================

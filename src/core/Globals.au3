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
_AuditLog("INFO", "Serveur démarré sur le port " & $g_iPort)
$g_iAuditCheckTimer = TimerInit()

; Boucle principale
While 1
    Local $iClientSocket = TCPAccept($g_iMainSocket)
    If $iClientSocket <> -1 Then _HandleClient($iClientSocket)

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

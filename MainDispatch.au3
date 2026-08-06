#NoTrayIcon
Opt("MustDeclareVars", 0)
Opt("GUIOnEventMode", 0)

#include <File.au3>
#include <String.au3>
#include <Date.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <WinAPI.au3>
#include <IE.au3>

; ========== CORE ==========
#include "src\core\Globals.au3"
#include "src\core\Constants.au3"
#include "src\core\JsonUtils.au3"
#include "src\core\DateUtils.au3"
#include "src\core\FileUtils.au3"
#include "src\core\BackupUtils.au3"
#include "src\core\Security.au3"
#include "src\core\Audit.au3"
#include "src\core\Config.au3"
#include "src\core\HttpResponse.au3"
#include "src\core\HttpRouter.au3"
#include "src\core\HttpServer.au3"

; ========== SERVICES ==========
#include "src\services\StateService.au3"
#include "src\services\ContactsService.au3"
#include "src\services\NetworkStateService.au3"
#include "src\services\ConfigService.au3"
#include "src\services\JobService.au3"

; ========== FEATURES ==========
; ETMS
#include "src\features\etms\Action_ETMS.au3"

; COMAT
#include "src\features\comat\Action_COMAT.au3"
#include "src\features\comat\Batch_COMAT.au3"

; FC
#include "src\features\fc\Action_FC.au3"
#include "src\features\fc\Batch_FC.au3"
#include "src\features\fc\Audit_FC.au3"

; CMR
#include "src\features\cmr\Action_CMR.au3"
#include "src\features\cmr\CMR_Engine.au3"
#include "src\features\cmr\CMR_Mail.au3"
#include "src\features\cmr\CMR_EDOC.au3"
#include "src\features\cmr\CMR_PDF.au3"

; EDOC
#include "src\features\edoc\Action_EDOC.au3"
#include "src\features\edoc\EDOC_Web.au3"
#include "src\features\edoc\EDOC_Master.au3"

; MAIL
#include "src\features\mail\Mail_Common.au3"
#include "src\features\mail\Mail_RDV.au3"
#include "src\features\mail\Mail_Alertes.au3"
#include "src\features\mail\Mail_CP.au3"

; DIAGNOSTICS
#include "src\features\diagnostics\Diagnostic.au3"
#include "src\features\diagnostics\StorageInfo.au3"
#include "src\features\diagnostics\ContactsCleanup.au3"

; ========== API ==========
#include "src\api\ApiPing.au3"
#include "src\api\ApiState.au3"
#include "src\api\ApiContacts.au3"
#include "src\api\ApiNetwork.au3"
#include "src\api\ApiConfig.au3"
#include "src\api\ApiJobs.au3"
#include "src\api\ApiActions.au3"

; ========== MAIN ==========
MainDispatch()

Func MainDispatch()
    DispatchInitialize()
    DispatchStartServer()
    DispatchOpenInterface()
    While 1
        DispatchProcessClients()
        DispatchProcessBackgroundJobs()
        Sleep(10)
    WEnd
EndFunc

Func DispatchInitialize()
    $giPort = 9500
    $gsSaveFile = @ScriptDir & "\dispatch.json"
    $gsDataFile = @ScriptDir & "\data.json"
    $gsStatusFile = @ScriptDir & "\status.json"
    $gsContactsFile = @ScriptDir & "\contacts.tsv"
    $gsConfigFile = @ScriptDir & "\config.ini"
    $gsHtmlFile = @ScriptDir & "\ui\index.html"
    AuditLog("Dispatch initialisation")
EndFunc

Func DispatchStartServer()
    $giMainSocket = TCPListen("0.0.0.0", $giPort, 100)
    If @error Then
        MsgBox(16, "Erreur", "Impossible de démarrer le serveur HTTP sur le port " & $giPort)
        Exit
    EndIf
    AuditLog("Serveur HTTP démarréĺ§ sur port " & $giPort)
EndFunc

Func DispatchOpenInterface()
    Local $sURL = "http://localhost:" & $giPort & "/"
    ShellExecute($sURL)
    AuditLog("Interface ouverte: " & $sURL)
EndFunc

Func DispatchProcessClients()
    Local $iSocket = TCPAccept($giMainSocket)
    If $iSocket = -1 Then Return
    HttpServer_HandleClient($iSocket)
    TCPCloseSocket($iSocket)
EndFunc

Func DispatchProcessBackgroundJobs()
EndFunc

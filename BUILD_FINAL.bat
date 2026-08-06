@echo off
echo ================================================
echo          DISPATCH - COMPILATION FINALE
echo ================================================
echo.

set AUTOIT_PATH=C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2Exe64.exe

if not exist "%AUTOIT_PATH%" (
    echo ERREUR: AutoIt non trouve
    echo Chemin attendu: %AUTOIT_PATH%
    echo.
    echo Veuillez installer AutoIt3 depuis:
    echo https://www.autoitscript.com/site/autoit/downloads/
    pause
    exit /b 1
)

echo 1. Nettoyage...
if exist Dispatch.exe del Dispatch.exe
if exist dispatch.exe.tmp del dispatch.exe.tmp

echo 2. Compilation de MainDispatch.au3...
echo.

"%AUTOIT_PATH%" /in MainDispatch.au3 /out Dispatch.exe /console

if exist Dispatch.exe (
    echo.
    echo ================================================
    echo                    SUCCES !
    echo ================================================
    echo.
    echo Dispatch.exe a ete genere avec succes.
    echo.
    echo POUR LANCER L'APPLICATION:
    echo   Double-cliquer sur Dispatch.exe
    echo.
    echo L'interface s'ouvrira automatiquement sur:
    echo   http://localhost:9500
    echo.
    echo ================================================
    echo.
    pause
    start Dispatch.exe
) else (
    echo.
    echo ================================================
    echo                   ECHEC
    echo ================================================
    echo.
    echo La compilation a echoue.
    echo.
    echo VERIFICATIONS:
    echo   1. AutoIt3 est-il installe ?
    echo   2. MainDispatch.au3 existe-t-il ?
    echo   3. Y a-t-il des erreurs dans le fichier ?
    echo.
    pause
)

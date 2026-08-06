@echo off
echo Compilation de Dispatch.exe...
echo.

set AUTOIT_PATH=C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2Exe64.exe

if not exist "%AUTOIT_PATH%" (
    echo ERREUR: AutoIt non trouve a: %AUTOIT_PATH%
    echo Veuillez installer AutoIt ou modifier le chemin dans ce script.
    pause
    exit /b 1
)

"%AUTOIT_PATH%" /in MainDispatch.au3 /out Dispatch.exe /icon dispatch.ico /console

if exist Dispatch.exe (
    echo.
    echo SUCCES: Dispatch.exe genere avec succes !
    echo.
    echo Fichiers a copier pour la production:
    echo   - Dispatch.exe
    echo   - ui\
    echo   - dispatch.json
    echo   - data.json
    echo   - status.json
    echo   - contacts.tsv
    echo   - config.ini
    echo.
) else (
    echo.
    echo ERREUR: La compilation a echoue.
    echo.
)

pause

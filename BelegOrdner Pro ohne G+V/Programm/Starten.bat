@echo off
setlocal EnableDelayedExpansion
title BelegOrdner Pro - ohne G+V
mode con: cols=74 lines=30
color 0B

set "PROG=%~dp0"
for %%I in ("%~dp0..") do set "BASIS=%%~fI"

cls
echo.
echo    ==============================================================
echo.
echo         B E L E G O R D N E R   P R O   -   O H N E   G + V
echo.
echo         Belege sortieren - Buchhaltung vorbereiten
echo.
echo    ==============================================================
echo.
color 0E
echo      Programmordner: %BASIS%
color 0B
echo.
echo    --------------------------------------------------------------
echo.

:: HINWEIS: Der E-Mail-Import laeuft hier BEWUSST NICHT automatisch.
:: Er wuerde sonst beim ersten Start ungefragt das Postfach des jeweiligen
:: Rechners auslesen. Der Import wird stattdessen im Programm selbst
:: gestartet - ueber den Knopf "E-Mails importieren" auf der Startseite.

color 0A
echo      [1/3]  Auf Neuerungen pruefen
color 07
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%BASIS%\update.ps1"
echo.

color 0A
echo      [2/3]  Dateiliste erstellen
color 07
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%BASIS%\manifest.ps1"
echo.

color 0A
echo      [3/3]  Programm starten
color 07

set "SRVSTATUS=NONE"
set "SRVROOT="
for /f "usebackq delims=" %%S in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$mein = '%PROG%'.TrimEnd([char]92); try { $r = (Invoke-RestMethod 'http://localhost:8044/serverinfo' -TimeoutSec 2).root; if ($r -and ($r.TrimEnd([char]92) -ieq $mein)) { 'OK' } else { 'ANDER' } } catch { 'NONE' }"`) do set "SRVSTATUS=%%S"

if "!SRVSTATUS!"=="OK" (
    echo             Laeuft bereits.
)
if "!SRVSTATUS!"=="ANDER" (
    color 0E
    echo             Anderer Ordner aktiv - wird beendet...
    color 07
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'powershell.exe' -and $_.ProcessId -ne $PID -and $_.CommandLine -like '*webserver.ps1*' -and $_.CommandLine -notlike '*%BASIS%*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    timeout /t 2 /nobreak >nul
    set "SRVSTATUS=NONE"
)
if "!SRVSTATUS!"=="NONE" (
    echo             Wird gestartet...
    start "BelegOrdner_Server" /min powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Minimized -File "%BASIS%\webserver.ps1"
    timeout /t 3 /nobreak >nul
)

set "BROWSER="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "BROWSER=%LocalAppData%\Google\Chrome\Application\chrome.exe"

if defined BROWSER (
    start "" "%BROWSER%" "http://localhost:8044/FinanzBuch.html"
) else (
    echo             Chrome nicht gefunden - Standardbrowser wird genutzt
    start "" "http://localhost:8044/FinanzBuch.html"
)

echo.
color 0B
echo    --------------------------------------------------------------
color 0A
echo.
echo      Fertig - das Programm oeffnet sich im Browser.
color 0E
echo.
echo      Beim ersten Start:
echo        1. Links auf "Einstellungen" - Firma und Bank eintragen
echo        2. Rechnungen in den Ordner "Rechnungen" legen
echo           ODER auf "E-Mails importieren" klicken, um sie
echo           aus dem eigenen Postfach zu holen
color 0B
echo.
echo    --------------------------------------------------------------
echo.
timeout /t 6 /nobreak >nul
exit

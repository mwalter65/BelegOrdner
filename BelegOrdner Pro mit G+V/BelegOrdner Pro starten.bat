@echo off
:: ============================================================
::  BelegOrdner Pro - Starter
::
::  Diese Datei startet das Programm.
::  Sie liegt ganz oben im Ordner, damit sie sofort zu finden ist.
::
::  WEITERGABE: Einfach diesen kompletten Ordner kopieren.
::  Das Programm laeuft von jedem Laufwerk und jedem Ordner.
:: ============================================================
title BelegOrdner Pro wird gestartet

if not exist "%~dp0Programm\Starten.bat" (
    color 0C
    echo.
    echo   FEHLER: Der Unterordner "Programm" fehlt.
    echo.
    echo   Bitte pruefen Sie, ob der Ordner vollstaendig kopiert wurde.
    echo   Er muss so aufgebaut sein:
    echo.
    echo      BelegOrdner Pro\
    echo         ^|-- BelegOrdner Pro starten.bat   ^(diese Datei^)
    echo         ^|-- webserver.ps1
    echo         ^|-- Programm\
    echo                ^|-- Starten.bat
    echo                ^|-- FinanzBuch.html
    echo.
    pause
    exit /b 1
)

call "%~dp0Programm\Starten.bat"

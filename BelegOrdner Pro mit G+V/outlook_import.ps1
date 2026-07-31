# outlook_import.ps1 â€“ Startet Outlook automatisch und importiert PDF-Rechnungen

param([int]$TageRueckwirkend = 90)

# Pfad relativ zum Skript – dadurch läuft das Programm von jedem Laufwerk (C:, D:, USB-Stick …)
$zielOrdner  = Join-Path $PSScriptRoot "Programm\Rechnungen"
$mailOrdner  = Join-Path $PSScriptRoot "Programm\EMails"
$mailIndex   = Join-Path $mailOrdner "_email_index.csv"
$vonDatum    = (Get-Date).AddDays(-$TageRueckwirkend)
$outlookPfad = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$mailsNeu    = 0

# Es werden ausschliesslich KOPIEN abgelegt – die Originale bleiben in Outlook.
if (-not (Test-Path $mailOrdner)) { New-Item -ItemType Directory -Force -Path $mailOrdner | Out-Null }
if (-not (Test-Path $mailIndex)) {
  Set-Content -Path $mailIndex -Value '"Datum";"Absender";"Betreff";"EMail-Datei";"PDF-Datei"' -Encoding UTF8
}

# â”€â”€ Outlook starten falls nicht aktiv â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
$proc = Get-Process outlook -ErrorAction SilentlyContinue
if (-not $proc) {
    Write-Host "  Outlook wird gestartet..."
    if (Test-Path $outlookPfad) {
        Start-Process $outlookPfad
    } else {
        Start-Process "outlook.exe"   # Fallback: Ã¼ber PATH
    }
    Write-Host "  Warte bis Outlook bereit ist (15 Sekunden)..."
    Start-Sleep -Seconds 15
} else {
    Write-Host "  Outlook lÃ¤uft bereits."
}

# â”€â”€ Nochmal sicherstellen dass der COM-Server bereit ist â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
$versuche = 0
$outlook  = $null
while ($versuche -lt 5 -and $outlook -eq $null) {
    try {
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
    } catch {
        $versuche++
        Write-Host "  Warte auf Outlook COM ($versuche/5)..."
        Start-Sleep -Seconds 5
    }
}

if ($outlook -eq $null) {
    Write-Host "  FEHLER: Outlook konnte nicht gestartet werden."
    exit 1
}

Write-Host "  Durchsuche Posteingang (letzte $TageRueckwirkend Tage)..."

try {
    $mapi  = $outlook.GetNamespace("MAPI")
    $inbox = $mapi.GetDefaultFolder(6)   # 6 = olFolderInbox
    $neu   = 0

    function SucheInOrdner($ordner) {
        foreach ($mail in $ordner.Items) {
            try {
                if ($mail.ReceivedTime -lt $vonDatum) { continue }
                if ($mail.Attachments.Count -eq 0)    { continue }

                $pdfsDieserMail = @()
                foreach ($att in $mail.Attachments) {
                    if ($att.FileName -notmatch '\.pdf$') { continue }
                    $sauber = $att.FileName -replace '[\\/:*?"<>|]','_'
                    $ziel   = Join-Path $zielOrdner $sauber
                    # KOPIE herunterladen - das Original bleibt unveraendert in Outlook
                    if (-not (Test-Path $ziel)) {
                        $att.SaveAsFile($ziel)
                        $neu++
                        Write-Host "    OK  $sauber"
                    }
                    $pdfsDieserMail += $sauber
                }

                # E-Mail selbst als Kopie sichern (GoBD: E-Mail kann Handelsbrief sein).
                # Das Original bleibt ebenfalls unberuehrt in Outlook.
                if ($pdfsDieserMail.Count -gt 0) {
                    try {
                        $dat  = $mail.ReceivedTime.ToString('yyyy-MM-dd')
                        $von  = ''
                        try { $von = $mail.SenderName } catch { $von = 'unbekannt' }
                        $betr = ''
                        try { $betr = $mail.Subject } catch { $betr = 'ohne Betreff' }
                        $namensteil = ("$dat" + '_' + "$von" + '_' + "$betr")
                        $namensteil = $namensteil -replace '[\\/:*?"<>|\r\n\t]','_'
                        if ($namensteil.Length -gt 110) { $namensteil = $namensteil.Substring(0,110) }
                        $msgZiel = Join-Path $mailOrdner ($namensteil + '.msg')
                        if (-not (Test-Path $msgZiel)) {
                            $mail.SaveAs($msgZiel, 3)   # 3 = olMSG
                            $mailsNeu++
                            # Verknuepfung E-Mail <-> PDF fuer das Monatsarchiv festhalten
                            foreach ($pdfName in $pdfsDieserMail) {
                                $zeile = '"' + $dat + '";"' + ($von -replace '"','') + '";"' + ($betr -replace '"','') + '";"' + (Split-Path $msgZiel -Leaf) + '";"' + $pdfName + '"'
                                Add-Content -Path $mailIndex -Value $zeile -Encoding UTF8
                            }
                        }
                    } catch { }
                }
            } catch { }
        }
        foreach ($sub in $ordner.Folders) { SucheInOrdner $sub }
    }

    SucheInOrdner $inbox
    try { SucheInOrdner ($mapi.GetDefaultFolder(4))  } catch {}  # 4  = Junk/Spam
    try { SucheInOrdner ($mapi.GetDefaultFolder(5))  } catch {}  # 5  = Gesendete
    try { SucheInOrdner ($mapi.GetDefaultFolder(3))  } catch {}  # 3  = GelÃ¶schte
    try { SucheInOrdner ($mapi.GetDefaultFolder(23)) } catch {}  # 23 = Archiv
    # Alle weiteren Top-Level-Ordner (benutzerdefinierte Ordner)
    foreach ($topOrdner in $mapi.Folders) {
        try { SucheInOrdner $topOrdner } catch {}
    }
    Write-Host "  $neu neue PDF-Rechnung(en) als Kopie abgelegt."
    Write-Host "  $mailsNeu zugehoerige E-Mail(s) als Kopie gesichert (Originale bleiben in Outlook)."

} catch {
    Write-Host "  Fehler beim Lesen: $_"
}


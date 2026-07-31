# thunderbird_import.ps1 â€“ Extrahiert PDF-AnhÃ¤nge aus Thunderbird mbox-Dateien

param([int]$TageRueckwirkend = 90)

# Pfad relativ zum Skript – dadurch läuft das Programm von jedem Laufwerk (C:, D:, USB-Stick …)
$zielOrdner = Join-Path $PSScriptRoot "Programm\Rechnungen"
$vonDatum   = (Get-Date).AddDays(-$TageRueckwirkend)
# Thunderbird-Profil des aktuell angemeldeten Benutzers (nicht fest verdrahtet)
$profileDir = Join-Path $env:APPDATA "Thunderbird\Profiles"
$gesamt_neu = 0
$gesamt_emails = 0

if (-not (Test-Path $zielOrdner)) { New-Item -ItemType Directory $zielOrdner | Out-Null }

$profile = Get-ChildItem $profileDir -Directory -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $profile) { Write-Host "Kein Thunderbird-Profil gefunden."; exit 1 }
Write-Host "Thunderbird-Profil: $($profile.Name)"

$mboxFiles = Get-ChildItem $profile.FullName -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -eq '' -and $_.Length -gt 10000 -and $_.Name -notmatch '^\.' } |
    Sort-Object Length -Descending

Write-Host "Durchsuche $($mboxFiles.Count) Ordner (letzte $TageRueckwirkend Tage)..."

function SpeicherePDF($headerBlock, $daten) {
    if ($daten.Count -eq 0) { return }

    # Dateiname aus Header-Block rekonstruieren
    # 1) filename*N= Teile zusammensetzen (RFC 2231 split)
    $teile = @{}
    $headerBlock | Select-String 'filename\*(\d+)\s*=\s*["\s]*([^";\r\n]+)' -AllMatches |
        ForEach-Object { $_.Matches } | ForEach-Object {
            $idx  = [int]$_.Groups[1].Value
            $teil = $_.Groups[2].Value.Trim().Trim('"')
            $teile[$idx] = $teil
        }
    $pdfName = ''
    if ($teile.Count -gt 0) {
        $pdfName = ($teile.Keys | Sort-Object | ForEach-Object { $teile[$_] }) -join ''
    }

    # 2) filename= (einfach, kein Split)
    if (-not $pdfName) {
        $m = [regex]::Match($headerBlock, 'filename\s*=\s*["\s]*([^";\r\n]+\.pdf)', 'IgnoreCase')
        if ($m.Success) { $pdfName = $m.Groups[1].Value.Trim().Trim('"') }
    }

    # 3) name= aus Content-Type
    if (-not $pdfName) {
        $m = [regex]::Match($headerBlock, 'name\s*=\s*["\s]*([^";\r\n]+)', 'IgnoreCase')
        if ($m.Success) {
            $n = $m.Groups[1].Value.Trim().Trim('"')
            if ($n) { $pdfName = if ($n -match '\.pdf$') { $n } else { $n + '.pdf' } }
        }
    }

    if (-not $pdfName) { $pdfName = "Rechnung_Thunderbird_$($script:gesamt_neu + 1).pdf" }

    # Nur .pdf Endung
    if ($pdfName -notmatch '\.pdf$') { return }

    $sauber = $pdfName -replace '[\\/:*?"<>|#%]', '_'
    $ziel   = Join-Path $zielOrdner $sauber
    if (Test-Path $ziel) { return }
    try {
        $b64   = ($daten -join '') -replace '\s', ''
        $bytes = [Convert]::FromBase64String($b64)
        if ($bytes.Length -lt 200) { return }
        [System.IO.File]::WriteAllBytes($ziel, $bytes)
        $script:gesamt_neu++
        Write-Host "  OK $sauber"
    } catch { }
}

function ParseMbox($pfad) {
    $enc    = [System.Text.Encoding]::GetEncoding(28591)
    # FileShare.ReadWrite erlaubt Lesen auch wenn Thunderbird die Datei geöffnet hat
    $fs     = [System.IO.File]::Open($pfad, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $reader = [System.IO.StreamReader]::new($fs, $enc)

    $aktuell     = $false
    $datumGelesen = $false
    $state        = 'HEADER'   # HEADER | BODY | MIME_HEADER | PDF_DATA

    # MIME-Part Header-Puffer (logische Zeilen zusammengefÃ¼gt)
    $mimeHeaderBuf = [System.Text.StringBuilder]::new()
    $istPdfPart    = $false
    $istBase64     = $false
    $pdfDaten      = [System.Collections.Generic.List[string]]::new()

    while (-not $reader.EndOfStream) {
        $z = $reader.ReadLine()

        # â”€â”€ Neue E-Mail â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if ($z.StartsWith('From ') -and $z.Length -gt 10) {
            if ($state -eq 'PDF_DATA' -and $aktuell -and $pdfDaten.Count -gt 0) {
                SpeicherePDF $mimeHeaderBuf.ToString() $pdfDaten
            }
            $aktuell       = $false
            $datumGelesen  = $false
            $state         = 'HEADER'
            $mimeHeaderBuf.Clear() | Out-Null
            $istPdfPart    = $false
            $istBase64     = $false
            $pdfDaten.Clear()
            $script:gesamt_emails++
            continue
        }

        # â”€â”€ E-Mail Haupt-Header â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if ($state -eq 'HEADER') {
            if ($z -eq '') {
                if (-not $datumGelesen) { $aktuell = $true }
                $state = 'BODY'
            } elseif (-not $datumGelesen -and $z.StartsWith('Date:')) {
                $datumGelesen = $true
                try {
                    $d = [System.DateTime]::Parse($z.Substring(5).Trim(),
                         [System.Globalization.CultureInfo]::InvariantCulture,
                         [System.Globalization.DateTimeStyles]::None)
                    $aktuell = ($d -ge $vonDatum)
                } catch { $aktuell = $true }
            }
            continue
        }

        if (-not $aktuell) { continue }

        # â”€â”€ MIME-Body â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if ($state -eq 'BODY') {
            # MIME-Grenze â†’ neuer Part beginnt
            if ($z.StartsWith('--')) {
                if ($istPdfPart -and $pdfDaten.Count -gt 0) {
                    SpeicherePDF $mimeHeaderBuf.ToString() $pdfDaten
                }
                $mimeHeaderBuf.Clear() | Out-Null
                $istPdfPart = $false
                $istBase64  = $false
                $pdfDaten.Clear()
                $state = 'MIME_HEADER'
                continue
            }
            # Kein Boundary â€“ inline body (kein Attachment) â†’ weiter
            continue
        }

        # â”€â”€ MIME-Part Header lesen â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if ($state -eq 'MIME_HEADER') {
            if ($z -eq '') {
                # Ende der MIME-Header: auswerten
                $hdr = $mimeHeaderBuf.ToString()
                $istPdfPart = $hdr -match 'Content-Type:\s*(application/pdf|application/octet-stream)'
                $istBase64  = $hdr -match 'Content-Transfer-Encoding:\s*base64'
                if ($istPdfPart -and $istBase64) {
                    $state = 'PDF_DATA'
                } else {
                    $state = 'BODY'
                    $istPdfPart = $false
                }
                continue
            }
            # Header-Zeilen sammeln (Fortsetzungszeilen beginnen mit Leerzeichen/Tab)
            $mimeHeaderBuf.Append($z) | Out-Null
            $mimeHeaderBuf.Append("`n") | Out-Null
            continue
        }

        # â”€â”€ Base64-Daten sammeln â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if ($state -eq 'PDF_DATA') {
            if ($z.StartsWith('--')) {
                # Boundary = Ende dieses Parts
                SpeicherePDF $mimeHeaderBuf.ToString() $pdfDaten
                $mimeHeaderBuf.Clear() | Out-Null
                $istPdfPart = $false
                $istBase64  = $false
                $pdfDaten.Clear()
                $state = 'MIME_HEADER'
                continue
            }
            $clean = $z.Trim()
            if ($clean -match '^[A-Za-z0-9+/=]+$') {
                $pdfDaten.Add($clean)
            }
            continue
        }
    }

    if ($state -eq 'PDF_DATA' -and $aktuell -and $pdfDaten.Count -gt 0) {
        SpeicherePDF $mimeHeaderBuf.ToString() $pdfDaten
    }
    $reader.Dispose()
    $fs.Dispose()
}

foreach ($mbox in $mboxFiles) {
    $mb = [math]::Round($mbox.Length / 1MB, 0)
    Write-Host "[$mb MB] $($mbox.Name)..."
    try { ParseMbox $mbox.FullName }
    catch { Write-Host "  Fehler: $_" }
}

Write-Host ""
Write-Host "$gesamt_neu neue PDF aus Thunderbird ($gesamt_emails E-Mails geprueft)."


# datensicherung.ps1 - erstellt ein abgeschlossenes Monatspaket
#
# Inhalt des Pakets:
#   - Programmstand (Zuordnungen, Salden, Einstellungen)
#   - Kopien der zugeordneten Rechnungs-PDFs
#   - Kopien der zugehoerigen E-Mails (.msg)
#   - Kontoauszug
#   - Zuordnungsliste + Aenderungsprotokoll
#   - Pruefsummen (SHA-256) ueber jede Datei
#   - Verfahrensdokumentation
#
# Ablage:  <Ziel>\Datensicherung\<Jahr>\<Jahr-Monat>_Datensicherung.zip
#
# Die Originale bleiben unangetastet - es werden ausschliesslich Kopien abgelegt.

$basis     = Join-Path $PSScriptRoot "Programm"
$auftrag   = Join-Path $basis "_sicherung_auftrag.json"
$statusDat = Join-Path $basis "_sicherung_status.json"

function Schreibe-Status($phase, $text, $prozent, $fertig, $ergebnis) {
    $obj = [ordered]@{
        phase    = $phase
        text     = $text
        prozent  = $prozent
        fertig   = $fertig
        zeit     = (Get-Date).ToString('o')
    }
    if ($ergebnis) { $obj.ergebnis = $ergebnis }
    # Ohne Byte-Markierung schreiben, damit Browser und Skripte es sauber lesen koennen
    try {
        $json = $obj | ConvertTo-Json -Depth 6
        [IO.File]::WriteAllText($statusDat, $json, (New-Object Text.UTF8Encoding($false)))
    } catch {}
}

try {
    if (-not (Test-Path $auftrag)) { Schreibe-Status 'fehler' 'Kein Auftrag gefunden' 0 $true $null; exit 1 }
    Schreibe-Status 'start' 'Sicherung wird vorbereitet' 2 $false $null

    $d = Get-Content $auftrag -Raw -Encoding UTF8 | ConvertFrom-Json

    $monat = if ($d.monat) { ($d.monat -replace '[^0-9\-]','') } else { 'unbekannt' }
    $jahr  = if ($monat -match '^(\d{4})') { $matches[1] } else { 'unbekannt' }
    $label = if ($d.label) { $d.label } else { $monat }

    # Zielordner: <Ziel>\Datensicherung\<Jahr>\
    $zielBasis = if ($d.zielBasis -and (Test-Path $d.zielBasis)) { $d.zielBasis } else { $basis }
    $jahrOrdner = Join-Path (Join-Path $zielBasis 'Datensicherung') $jahr
    if (-not (Test-Path $jahrOrdner)) { New-Item -ItemType Directory -Force -Path $jahrOrdner | Out-Null }

    # Arbeitsordner – FullName verwenden, damit Kurz-/Langform des Temp-Pfads
    # nicht zu falschen Pfadangaben in den Pruefsummen fuehrt
    $bau = Join-Path $env:TEMP ("sicherung_" + $monat + "_" + (Get-Random))
    New-Item -ItemType Directory -Force -Path $bau | Out-Null
    $bau = (Get-Item $bau).FullName
    New-Item -ItemType Directory -Force -Path (Join-Path $bau 'Rechnungen') | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $bau 'EMails') | Out-Null

    # E-Mail-Zuordnung einlesen (welche E-Mail gehoert zu welchem PDF)
    $mailZuPdf = @{}
    $idx = Join-Path $basis 'EMails\_email_index.csv'
    if (Test-Path $idx) {
        try {
            Import-Csv $idx -Delimiter ';' -Encoding UTF8 | ForEach-Object {
                if ($_.'PDF-Datei') { $mailZuPdf[$_.'PDF-Datei'] = $_.'EMail-Datei' }
            }
        } catch {}
    }

    # Rechnungen und E-Mails kopieren
    Schreibe-Status 'kopieren' 'Rechnungen werden kopiert' 10 $false $null
    $liste = @($d.belege)
    $anz = $liste.Count
    $kopiert = 0; $mails = 0; $fehlend = @(); $i = 0

    foreach ($nr in $liste) {
        $i++
        if (-not $nr) { continue }
        if ($i % 10 -eq 0) {
            $p = 10 + [int](50 * $i / [Math]::Max($anz,1))
            Schreibe-Status 'kopieren' "Rechnung $i von $anz" $p $false $null
        }
        # Der Eintrag kann einen Unterordner enthalten (z.B. "Pdf/Juni 2026/Rechnung.pdf").
        # Backslashes verwenden und die Ordnerstruktur im Paket beibehalten.
        $rel = ($nr -replace '/','\')
        if ($rel -notmatch '\.pdf$') { $rel = "$rel.pdf" }
        $quelle = Join-Path (Join-Path $basis 'Rechnungen') $rel
        if (-not (Test-Path $quelle)) {
            # Falls nicht gefunden: nach dem reinen Dateinamen im ganzen Baum suchen
            $nurName = Split-Path $rel -Leaf
            $treffer = @(Get-ChildItem (Join-Path $basis 'Rechnungen') -File -Recurse -Filter $nurName -ErrorAction SilentlyContinue)
            if ($treffer.Count -eq 1) { $quelle = $treffer[0].FullName; $rel = $nurName }
        }
        if (Test-Path $quelle) {
            $ziel = Join-Path (Join-Path $bau 'Rechnungen') $rel
            $zielOrdner = Split-Path $ziel -Parent
            if (-not (Test-Path $zielOrdner)) { New-Item -ItemType Directory -Force -Path $zielOrdner | Out-Null }
            Copy-Item $quelle $ziel -Force
            $kopiert++
            $mailDatei = $mailZuPdf[(Split-Path $rel -Leaf)]
            if ($mailDatei) {
                $mq = Join-Path $basis "EMails\$mailDatei"
                $mz = Join-Path $bau "EMails\$mailDatei"
                if ((Test-Path $mq) -and -not (Test-Path $mz)) { Copy-Item $mq $mz -Force; $mails++ }
            }
        } else {
            $fehlend += $rel
        }
    }

    # Kontoauszug mitnehmen
    if ($d.kontoauszug) {
        $ka = Join-Path $basis ("Kontoauszuege\" + $d.kontoauszug)
        if (Test-Path $ka) { Copy-Item $ka (Join-Path $bau (Split-Path $ka -Leaf)) -Force }
    }

    # Textdateien schreiben
    Schreibe-Status 'listen' 'Listen und Protokolle werden geschrieben' 65 $false $null
    if ($d.zuordnungCsv) { Set-Content -Path (Join-Path $bau 'Zuordnung.csv') -Value $d.zuordnungCsv -Encoding UTF8 }
    if ($d.protokollCsv) { Set-Content -Path (Join-Path $bau 'Aenderungsprotokoll.csv') -Value $d.protokollCsv -Encoding UTF8 }
    if ($d.verfahrensdok) { Set-Content -Path (Join-Path $bau 'Verfahrensdokumentation.html') -Value $d.verfahrensdok -Encoding UTF8 }
    if ($d.programmstand) {
        ($d.programmstand | ConvertTo-Json -Depth 12) | Set-Content -Path (Join-Path $bau 'Programmstand.json') -Encoding UTF8
    }

    # Kurzuebersicht
    $info = @(
        "Datensicherung $label",
        "Erstellt am : $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')",
        "Rechnungen  : $kopiert",
        "E-Mails     : $mails",
        "Fehlend     : $($fehlend.Count)",
        "",
        "Hinweis: Dieses Paket enthaelt ausschliesslich KOPIEN.",
        "Die Originale verbleiben im E-Mail-Programm und im Programmordner."
    )
    if ($fehlend.Count -gt 0) {
        $info += ""
        $info += "NICHT GEFUNDENE BELEGE"
        $info += "----------------------"
        $info += "Diese Buchungen haben eine Zuordnung, aber die Datei fehlt:"
        $info += ""
        foreach ($f in $fehlend) {
            $nurName = Split-Path $f -Leaf
            # Passende Buchungsdaten suchen, damit der Beleg zuzuordnen ist
            $treffer = @($d.belegeInfo | Where-Object {
                $_.datei -and ((Split-Path ($_.datei -replace '/','\') -Leaf) -eq $nurName)
            })
            if ($treffer.Count -gt 0) {
                foreach ($tr in $treffer) {
                    $dat = if ($tr.datum) { try { ([datetime]$tr.datum).ToString('dd.MM.yyyy') } catch { $tr.datum } } else { '' }
                    $info += ("  {0,-11} {1,-38} {2,12} EUR" -f $dat, $tr.name, $tr.betrag)
                    $info += ("              fehlender Beleg: {0}" -f $f)
                    $info += ""
                }
            } else {
                $info += ("  {0}" -f $f)
                $info += ""
            }
        }
    }
    Set-Content -Path (Join-Path $bau 'Inhalt.txt') -Value $info -Encoding UTF8

    # Pruefsummen ueber alle Dateien
    Schreibe-Status 'pruefsummen' 'Pruefsummen werden berechnet' 78 $false $null
    $ps = @("Pruefsummen (SHA-256) - Datensicherung $label", "Erstellt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')", "")
    Get-ChildItem $bau -Recurse -File | Sort-Object FullName | ForEach-Object {
        $rel = $_.FullName.Substring($bau.Length + 1)
        $h = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
        $ps += ("{0}  {1}" -f $h, $rel)
    }
    Set-Content -Path (Join-Path $bau 'Pruefsummen.txt') -Value $ps -Encoding UTF8

    # Packen
    Schreibe-Status 'packen' 'Paket wird erstellt' 88 $false $null
    $zip = Join-Path $jahrOrdner ($monat + '_Datensicherung.zip')
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $bau '*') -DestinationPath $zip -CompressionLevel Optimal
    Remove-Item $bau -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $auftrag -Force -ErrorAction SilentlyContinue

    $mb = [math]::Round((Get-Item $zip).Length / 1MB, 2)
    $erg = [ordered]@{
        datei      = $zip
        groesseMB  = $mb
        rechnungen = $kopiert
        emails     = $mails
        fehlend    = @($fehlend)
    }
    Schreibe-Status 'fertig' "Datensicherung erstellt ($mb MB)" 100 $true $erg
    Write-Host "Datensicherung erstellt: $zip ($mb MB)"

} catch {
    Schreibe-Status 'fehler' ("Fehler: " + $_.Exception.Message) 0 $true $null
    Write-Host "FEHLER: $($_.Exception.Message)"
    exit 1
}

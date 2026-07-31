# update.ps1 - holt vor dem Start eine neuere Fassung, falls es eine gibt
#
# Ablauf:
#   1. Eigene Version und GitHub-Adresse aus FinanzBuch.html lesen
#   2. version.json abrufen (hoechstens 8 Sekunden warten)
#   3. Ist dort etwas Neueres? Dann herunterladen
#   4. Heruntergeladenes pruefen, bevor irgendetwas ersetzt wird
#   5. Alte Fassung zur Seite legen, neue einsetzen
#
# Grundsatz: Dieses Skript darf den Start NIEMALS verhindern.
# Kein Netz, kein Server, kaputte Datei - es meldet das und geht weiter.
# Eine Buchhaltung darf nicht daran scheitern, dass ein Server schweigt.

$ErrorActionPreference = 'SilentlyContinue'
$basis   = Join-Path $PSScriptRoot 'Programm'
$prog    = Join-Path $basis 'FinanzBuch.html'
$archiv  = Join-Path $basis '_alte_Fassungen'

function Melde($text, $farbe) {
    if ($farbe) { Write-Host "             $text" -ForegroundColor $farbe }
    else        { Write-Host "             $text" }
}

if (-not (Test-Path $prog)) { Melde 'Programmdatei nicht gefunden - uebersprungen'; exit 0 }

try {
    $utf8 = New-Object Text.UTF8Encoding($false)
    $html = [IO.File]::ReadAllText($prog, $utf8)

    $mVer = [regex]::Match($html, "PROGRAMM_VERSION\s*=\s*'([^']+)'")
    $mBas = [regex]::Match($html, "UPDATE_BASIS\s*=\s*'([^']*)'")
    $mKen = [regex]::Match($html, "PROGRAMM_KENNUNG\s*=\s*'([^']+)'")

    if (-not $mVer.Success) { Melde 'Keine Versionsangabe im Programm - uebersprungen'; exit 0 }
    $meine = $mVer.Groups[1].Value
    $adr   = if ($mBas.Success) { $mBas.Groups[1].Value.TrimEnd('/') } else { '' }
    $kennung = if ($mKen.Success) { $mKen.Groups[1].Value } else { 'windows' }

    # Die Adresse getrennt merken.
    #
    # Grund: Beim Update wird die Programmdatei ersetzt. Steht in der
    # neuen Datei keine Adresse, waere die Kette danach unterbrochen -
    # das Geraet bekaeme nie wieder ein Update und niemand merkte es.
    # Deshalb liegt sie zusaetzlich hier und wird bei Bedarf wieder
    # in die neue Datei eingesetzt.
    $merkzettel = Join-Path $PSScriptRoot 'update-adresse.txt'
    if ([string]::IsNullOrWhiteSpace($adr) -and (Test-Path $merkzettel)) {
        $adr = ([IO.File]::ReadAllText($merkzettel, $utf8)).Trim().TrimEnd('/')
        if ($adr) { Melde 'Adresse aus dem Merkzettel uebernommen' }
    }
    if (-not [string]::IsNullOrWhiteSpace($adr)) {
        [IO.File]::WriteAllText($merkzettel, $adr, $utf8)
    }

    if ([string]::IsNullOrWhiteSpace($adr)) {
        Melde "Version $meine (keine Update-Adresse hinterlegt)"
        exit 0
    }

    Melde "Version $meine - pruefe auf Neuerungen..."

    # ── version.json holen ────────────────────────────────────
    $alteVorgabe = [Net.ServicePointManager]::SecurityProtocol
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $info = $null
    try {
        $roh = Invoke-WebRequest -Uri "$adr/version.json" -UseBasicParsing -TimeoutSec 8
        $info = $roh.Content | ConvertFrom-Json
    } catch {
        Melde 'Kein Netz oder Server nicht erreichbar - Start wie gewohnt' Yellow
        exit 0
    }

    # Angaben zur eigenen Ausgabe, sonst die allgemeinen
    $meins = $null
    if ($info.ausgaben -and $info.ausgaben.$kennung) { $meins = $info.ausgaben.$kennung } else { $meins = $info }
    $neueste = $meins.version
    if (-not $neueste) { $neueste = $info.version }
    if (-not $neueste) { Melde 'Angaben unvollstaendig - uebersprungen'; exit 0 }

    # ── Vergleich Zahl fuer Zahl (1.2.10 ist neuer als 1.2.9) ─
    function IstNeuer($a, $b) {
        $x = @($a -split '\.' | ForEach-Object { [int]($_ -replace '\D','0') })
        $y = @($b -split '\.' | ForEach-Object { [int]($_ -replace '\D','0') })
        for ($i = 0; $i -lt [Math]::Max($x.Count, $y.Count); $i++) {
            $xa = if ($i -lt $x.Count) { $x[$i] } else { 0 }
            $ya = if ($i -lt $y.Count) { $y[$i] } else { 0 }
            if ($xa -ne $ya) { return ($xa -gt $ya) }
        }
        return $false
    }

    if (-not (IstNeuer $neueste $meine)) {
        Melde "Version $meine ist aktuell" Green
        exit 0
    }

    Melde "Neue Fassung $neueste gefunden - wird geladen..." Cyan

    # ── Herunterladen ─────────────────────────────────────────
    $dateiname = $meins.datei
    if (-not $dateiname) { $dateiname = 'FinanzBuch.html' }
    $tmp = Join-Path $env:TEMP ("belegordner_" + $neueste + ".html")
    try {
        Invoke-WebRequest -Uri "$adr/$dateiname" -OutFile $tmp -UseBasicParsing -TimeoutSec 90
    } catch {
        Melde 'Herunterladen fehlgeschlagen - Start mit bisheriger Fassung' Yellow
        exit 0
    }

    # ── Pruefen, BEVOR etwas ersetzt wird ─────────────────────
    # Eine halb geladene oder falsche Datei darf niemals die
    # laufende Fassung ueberschreiben.
    if (-not (Test-Path $tmp)) { Melde 'Datei kam nicht an - abgebrochen' Yellow; exit 0 }
    $groesse = (Get-Item $tmp).Length
    if ($groesse -lt 150000) {
        Melde "Geladene Datei ist zu klein ($([math]::Round($groesse/1KB)) KB) - abgebrochen" Yellow
        Remove-Item $tmp -Force; exit 0
    }
    $neu = [IO.File]::ReadAllText($tmp, $utf8)
    if ($neu -notmatch 'PROGRAMM_VERSION' -or $neu -notmatch '</html>') {
        Melde 'Geladene Datei ist unvollstaendig - abgebrochen' Yellow
        Remove-Item $tmp -Force; exit 0
    }
    $mNeu = [regex]::Match($neu, "PROGRAMM_VERSION\s*=\s*'([^']+)'")
    if (-not $mNeu.Success -or -not (IstNeuer $mNeu.Groups[1].Value $meine)) {
        Melde 'Geladene Datei ist nicht neuer - abgebrochen' Yellow
        Remove-Item $tmp -Force; exit 0
    }

    # ── Alte Fassung aufheben, neue einsetzen ─────────────────
    if (-not (Test-Path $archiv)) { New-Item -ItemType Directory -Force -Path $archiv | Out-Null }
    $sicherung = Join-Path $archiv ("FinanzBuch_" + $meine + "_" + (Get-Date -Format 'yyyy-MM-dd_HHmm') + ".html")
    Copy-Item $prog $sicherung -Force

    # Fehlt in der neuen Datei die Adresse, wird sie eingesetzt -
    # sonst gaebe es nach diesem Update nie wieder eines.
    if ($neu -match "UPDATE_BASIS\s*=\s*''") {
        $neu = [regex]::Replace($neu, "const UPDATE_BASIS = '';", "const UPDATE_BASIS = '$adr';")
        Melde 'Update-Adresse in die neue Fassung uebernommen'
    }
    [IO.File]::WriteAllText($prog, $neu, $utf8)
    Remove-Item $tmp -Force

    Melde "Aktualisiert: $meine  ->  $neueste" Green
    if ($meins.hinweis) { Melde $meins.hinweis }
    Melde "Vorherige Fassung liegt in _alte_Fassungen"
    Melde "Ihre Zuordnungen und Einstellungen bleiben erhalten"

} catch {
    Melde 'Update uebersprungen - Start wie gewohnt' Yellow
}
exit 0

# Manifest erstellen: listet ALLE Dateien je Ordner
# Pfade relativ zum Skript – dadurch läuft das Programm von jedem Laufwerk (C:, D:, USB-Stick …)
$basis = Join-Path $PSScriptRoot "Programm"

$k = Get-ChildItem (Join-Path $basis 'Kontoauszuege')  -File -ErrorAction SilentlyContinue |
     Where-Object { $_.Name -notlike 'HIER_*' } |
     Sort-Object LastWriteTime -Descending |
     Select-Object -ExpandProperty Name

$m = Get-ChildItem (Join-Path $basis 'Monatsberichte') -File -ErrorAction SilentlyContinue |
     Where-Object { $_.Name -notlike 'HIER_*' } |
     Sort-Object LastWriteTime -Descending |
     Select-Object -ExpandProperty Name

# Rechnungen AUCH aus Unterordnern erfassen. Der Pfad wird relativ zum
# Rechnungen-Ordner gespeichert (mit / als Trenner), damit der Server ihn
# direkt ausliefern kann. Ohne das waeren abgelegte Unterordner unsichtbar.
$rOrdner = Join-Path $basis 'Rechnungen'
$r = @()
if (Test-Path $rOrdner) {
    $laenge = (Get-Item $rOrdner).FullName.Length + 1
    $r = Get-ChildItem $rOrdner -File -Recurse -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -notlike 'HIER_*' } |
         ForEach-Object { $_.FullName.Substring($laenge) -replace '\\','/' }
}

$manifest = [ordered]@{
    kontoauszuege  = @($k | Where-Object { $_ })
    monatsberichte = @($m | Where-Object { $_ })
    rechnungen     = @($r | Where-Object { $_ })
}

$manifest | ConvertTo-Json | Set-Content (Join-Path $basis 'manifest.json') -Encoding UTF8
Write-Host "Manifest OK - $(@($k).Count) Kontoauszuege, $(@($m).Count) Monatsberichte, $(@($r).Count) Rechnungen"

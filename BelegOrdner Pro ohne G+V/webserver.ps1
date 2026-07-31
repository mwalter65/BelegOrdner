$port = 8044
# Pfad relativ zum Skript – dadurch läuft das Programm von jedem Laufwerk (C:, D:, USB-Stick …)
$root = Join-Path $PSScriptRoot "Programm"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

$mime = @{
  '.html'  = 'text/html; charset=utf-8'
  '.js'    = 'application/javascript'
  '.css'   = 'text/css'
  '.pdf'   = 'application/pdf'
  '.csv'   = 'text/plain; charset=utf-8'
  '.json'  = 'application/json; charset=utf-8'
  '.mt940' = 'text/plain; charset=utf-8'
  '.sta'   = 'text/plain; charset=utf-8'
}

while ($true) {
  try { $ctx = $listener.GetContext() } catch { break }
  $req = $ctx.Request
  $res = $ctx.Response
  $res.Headers.Add("Access-Control-Allow-Origin","*")
  $res.Headers.Add("Access-Control-Allow-Methods","GET, POST, OPTIONS")
  $res.Headers.Add("Access-Control-Allow-Headers","Content-Type,X-Filename")

  if ($req.HttpMethod -eq 'OPTIONS') { $res.StatusCode = 204; $res.Close(); continue }

  # RawUrl statt Url.LocalPath verwenden: LocalPath deutet ein kodiertes '#' im
  # Dateinamen als Trennzeichen und schneidet alles danach ab. Ausserdem war das
  # doppelte Entschluesseln fehleranfaellig - jetzt wird genau einmal dekodiert.
  $roh = $req.RawUrl
  $fz = $roh.IndexOf('?')
  if ($fz -ge 0) { $roh = $roh.Substring(0, $fz) }
  $path = [Uri]::UnescapeDataString($roh.TrimStart('/'))

  # Upload-Endpoint: POST /upload?name=Dateiname.pdf
  if ($req.HttpMethod -eq 'POST' -and $path -eq 'upload') {
    try {
      $fname = [Uri]::UnescapeDataString(($req.QueryString['name'] -replace '[\\/:*?"<>|]','_'))
      if ($fname -and $fname -match '\.pdf$') {
        $ziel = Join-Path "$root\Rechnungen" $fname
        $buf = New-Object byte[] $req.ContentLength64
        $req.InputStream.Read($buf, 0, $buf.Length) | Out-Null
        [IO.File]::WriteAllBytes($ziel, $buf)
        # Manifest neu erstellen
        Start-Process "powershell" -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$root\manifest.ps1`"" -WindowStyle Hidden
        $b = [Text.Encoding]::UTF8.GetBytes('{"ok":true}')
        $res.ContentType = 'application/json'; $res.ContentLength64 = $b.Length
        $res.OutputStream.Write($b, 0, $b.Length)
      } else {
        $res.StatusCode = 400
      }
    } catch { $res.StatusCode = 500 }
    $res.Close(); continue
  }

  # Auskunft: welchen Ordner liefert dieser Server aus? (wird von Starten.bat geprüft)
  if ($path -eq 'serverinfo') {
    $info = "{""root"":""" + ($root -replace '\\','\\') + """}"
    $b = [Text.Encoding]::UTF8.GetBytes($info)
    $res.ContentType = 'application/json; charset=utf-8'; $res.ContentLength64 = $b.Length
    $res.OutputStream.Write($b, 0, $b.Length); $res.Close(); continue
  }

  # Windows-Ordnerdialog oeffnen, damit der Ablageort sichtbar gewaehlt werden kann
  if ($path -eq 'ordner-waehlen') {
    $gewaehlt = ''
    try {
      $start = $req.QueryString['start']
      if (-not $start) { $start = $root }
      $shell = New-Object -ComObject Shell.Application
      $ordner = $shell.BrowseForFolder(0, 'Ordner fuer die Datensicherung waehlen', 0, $start)
      if ($ordner) { $gewaehlt = $ordner.Self.Path }
    } catch {}
    $antwort = '{"pfad":"' + ($gewaehlt -replace '\\','\\') + '"}'
    $b = [Text.Encoding]::UTF8.GetBytes($antwort)
    $res.ContentType = 'application/json; charset=utf-8'; $res.ContentLength64 = $b.Length
    $res.OutputStream.Write($b, 0, $b.Length); $res.Close(); continue
  }

  # Explorer am Zielort oeffnen und die Datei markieren
  if ($path -eq 'ordner-oeffnen') {
    try {
      $ziel = [Uri]::UnescapeDataString($req.QueryString['pfad'])
      if ($ziel -and (Test-Path $ziel)) {
        if (Test-Path $ziel -PathType Leaf) { Start-Process explorer.exe -ArgumentList ('/select,"' + $ziel + '"') }
        else { Start-Process explorer.exe -ArgumentList ('"' + $ziel + '"') }
      }
    } catch {}
    $b = [Text.Encoding]::UTF8.GetBytes('{"ok":true}')
    $res.ContentType = 'application/json'; $res.ContentLength64 = $b.Length
    $res.OutputStream.Write($b, 0, $b.Length); $res.Close(); continue
  }

  # Datensicherung anstossen: Auftrag ablegen, Skript im Hintergrund starten
  if ($req.HttpMethod -eq 'POST' -and $path -eq 'datensicherung') {
    try {
      $buf = New-Object byte[] $req.ContentLength64
      $gelesen = 0
      while ($gelesen -lt $buf.Length) {
        $n = $req.InputStream.Read($buf, $gelesen, $buf.Length - $gelesen)
        if ($n -le 0) { break }
        $gelesen += $n
      }
      $inhalt = [Text.Encoding]::UTF8.GetString($buf, 0, $gelesen)
      [IO.File]::WriteAllText((Join-Path $root '_sicherung_auftrag.json'), $inhalt, (New-Object Text.UTF8Encoding($false)))
      $statusDat = Join-Path $root '_sicherung_status.json'
      if (Test-Path $statusDat) { Remove-Item $statusDat -Force -ErrorAction SilentlyContinue }
      Start-Process "powershell" -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSScriptRoot\datensicherung.ps1`"" -WindowStyle Hidden
      $b = [Text.Encoding]::UTF8.GetBytes('{"ok":true}')
      $res.ContentType = 'application/json'; $res.ContentLength64 = $b.Length
      $res.OutputStream.Write($b, 0, $b.Length)
    } catch { $res.StatusCode = 500 }
    $res.Close(); continue
  }

  # Import-Endpoint: Outlook + Thunderbird + Manifest im Hintergrund starten
  # Die Skripte liegen NEBEN webserver.ps1 (eine Ebene über $root) – daher $PSScriptRoot
  if ($path -eq 'run-import') {
    Start-Process "powershell" -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSScriptRoot\outlook_import.ps1`" -TageRueckwirkend 90" -WindowStyle Hidden
    Start-Process "powershell" -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSScriptRoot\thunderbird_import.ps1`" -TageRueckwirkend 90" -WindowStyle Hidden
    Start-Process "powershell" -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSScriptRoot\manifest.ps1`"" -WindowStyle Hidden
    $b = [Text.Encoding]::UTF8.GetBytes('{"ok":true}')
    $res.ContentType = 'application/json'; $res.ContentLength64 = $b.Length
    $res.OutputStream.Write($b, 0, $b.Length); $res.Close(); continue
  }

  $full = Join-Path $root $path

  # Sicherheitscheck: kein Pfad ausserhalb von C:\claude
  if (-not $full.StartsWith($root)) {
    $res.StatusCode = 403; $res.Close(); continue
  }

  if (Test-Path $full -PathType Leaf) {
    try {
      $bytes = [IO.File]::ReadAllBytes($full)
      $ext   = [IO.Path]::GetExtension($full).ToLower()
      $res.ContentType     = if ($mime[$ext]) { $mime[$ext] } else { 'application/octet-stream' }
      $res.ContentLength64 = $bytes.Length
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
    } catch { $res.StatusCode = 500 }
  } else {
    $res.StatusCode = 404
    $b = [Text.Encoding]::UTF8.GetBytes("404")
    $res.ContentLength64 = $b.Length
    $res.OutputStream.Write($b, 0, $b.Length)
  }
  $res.Close()
}
$listener.Stop()

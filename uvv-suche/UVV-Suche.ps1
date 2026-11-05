<#
UVV-Suche.ps1

Durchsucht das klassische Outlook (eigenes Postfach, Gruppenpostfach, Online-Archiv)
nach Unterlagen zur UVV-Unterweisung und erstellt einen Bericht pro Mitarbeiter.

  - Keine Installation und keine Administratorrechte noetig.
  - Liest nur: In Outlook wird nichts veraendert, verschoben, geloescht oder gesendet.
  - Alle Daten bleiben auf diesem Laptop. Es wird nichts ins Internet uebertragen.

Ablauf:
  1. Alle Mails ab einem Datum werden durchsucht (alle Ordner der gewaehlten Postfaecher).
  2. Dokument-Anhaenge (PDF, Word, Excel, Bilder, ...) werden in einen Ordner gespeichert.
     Doppelte Anhaenge (z. B. weitergeleitet) werden nur einmal gespeichert.
  3. Der Text der Dokumente wird gelesen - bei PDFs und Bildern (auch Scans) mit der
     Texterkennung, die in Windows eingebaut ist.
  4. Gesucht wird nach den Namen aus namen.txt und nach UVV-Begriffen
     (UVV, Unterweisung, Belehrung, Teilnehmerliste, ...).
  5. Ergebnis: Bericht.html (pro Mitarbeiter), Treffer.csv und Zusammenfassung.csv (Excel)
     sowie Kopien der Treffer in einem Ordner pro Mitarbeiter.

Start:     Rechtsklick auf diese Datei -> "Mit PowerShell ausfuehren"
Abbrechen: jederzeit mit Strg+C - der Bericht wird mit dem bisherigen Stand geschrieben.

Diese Datei enthaelt absichtlich keine Umlaute, damit sie mit jeder Kodierung funktioniert.
#>

# =====================================================================================
#  EINSTELLUNGEN - bei Bedarf hier anpassen (bitte keine Umlaute verwenden)
# =====================================================================================

# Datei mit den Namen der Mitarbeiter, eine Person pro Zeile: "Nachname, Vorname".
# Leer lassen = namen.txt im selben Ordner wie dieses Skript.
$NamenDatei = ''

# Hauptordner fuer die Ergebnisse. Leer lassen = C:\Users\<Ihr Name>\UVV-Suche
# (bewusst nicht Desktop/Dokumente, weil diese oft mit OneDrive synchronisiert werden).
$AusgabeHauptordner = ''

# Vorschlag fuer "Mails ab Datum" beim Start (TT.MM.JJJJ).
$StandardSeit = '01.01.2024'

# Gruppenpostfaecher, die NICHT links in der Outlook-Ordnerliste stehen, hier mit
# Mailadresse eintragen, z. B.  @('gruppenpostfach@firma.de')
# Normalerweise nicht noetig: Postfaecher aus der Ordnerliste werden automatisch angeboten.
$ZusaetzlichePostfaecher = @()

# Suchbegriffe fuer den UVV-Bezug. Gross-/Kleinschreibung ist egal, Wortteile reichen
# ("Unterweis" findet Unterweisung, Unterweisungsnachweis, Sicherheitsunterweisung ...).
# Bitte Wortstaemme ohne Umlaute verwenden ("Teilnahmebest" statt "Teilnahmebestaetigung").
$Suchbegriffe = @(
    'UVV', 'Unterweis', 'Unterwiesen', 'Belehrung', 'Teilnehmerliste', 'Teilnahmeliste',
    'Teilnahmebest', 'Teilnahmenachweis', 'Schulungsnachweis', 'Unfallverh', 'DGUV'
)

# $true  = alle gefundenen Dokument-Anhaenge behalten (Ordner "Anhaenge")
# $false = am Ende nur die Dokumente behalten, in denen etwas gefunden wurde
$AlleAnhaengeBehalten = $true

# Texterkennung: hoechstens so viele Seiten pro PDF bzw. pro mehrseitigem Bild lesen
$OcrMaxSeiten = 5

# Groessere Anhaenge werden uebersprungen (MB)
$MaxAnhangMB = 60

# Kleinere Bilder werden uebersprungen, z. B. Logos und Signaturbilder (KB)
$MinBildKB = 40

# =====================================================================================
#  Ab hier nichts mehr aendern
# =====================================================================================

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'
$script:KB = 1024
$script:MB = 1048576

function New-Set([string[]]$werte) {
    $h = @{}
    foreach ($w in $werte) { $h[$w] = $true }
    return $h
}

$script:DokSet = New-Set @('pdf', 'doc', 'docx', 'docm', 'dot', 'dotx', 'rtf', 'odt', 'txt', 'csv',
    'xls', 'xlsx', 'xlsm', 'ods', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'odp',
    'jpg', 'jpeg', 'png', 'tif', 'tiff', 'bmp', 'gif', 'heic', 'msg', 'eml', 'zip', 'htm', 'html')
$script:BildSet = New-Set @('jpg', 'jpeg', 'png', 'tif', 'tiff', 'bmp', 'gif', 'heic')
$script:BinaerSet = New-Set @('doc', 'dot', 'xls', 'ppt', 'pps')
$script:TextSet = New-Set @('txt', 'csv', 'htm', 'html', 'eml')
# Diese Typen brauchen Texterkennung oder enthalten weitere Dateien -> werden zuletzt gelesen
$script:LangsamSet = New-Set @('pdf', 'jpg', 'jpeg', 'png', 'tif', 'tiff', 'bmp', 'gif', 'heic', 'msg', 'zip')
$script:ZipXmlTeile = @{
    'docx' = @('word/*.xml'); 'docm' = @('word/*.xml'); 'dotx' = @('word/*.xml')
    'xlsx' = @('xl/sharedStrings.xml', 'xl/worksheets/*.xml', 'xl/comments*.xml')
    'xlsm' = @('xl/sharedStrings.xml', 'xl/worksheets/*.xml', 'xl/comments*.xml')
    'pptx' = @('ppt/slides/*.xml', 'ppt/notesSlides/*.xml')
    'pptm' = @('ppt/slides/*.xml', 'ppt/notesSlides/*.xml')
    'ppsx' = @('ppt/slides/*.xml', 'ppt/notesSlides/*.xml')
    'odt'  = @('content.xml'); 'ods' = @('content.xml'); 'odp' = @('content.xml')
}
# Outlook-Ordner mit Synchronisierungskonflikten (Kopien) nicht durchsuchen
$script:IgnorierOrdner = New-Set @('Synchronisierungsprobleme', 'Sync Issues')

$script:ChAe = [string][char]0x00E4
$script:ChOe = [string][char]0x00F6
$script:ChUe = [string][char]0x00FC
$script:ChSz = [string][char]0x00DF

$script:LogDatei = $null
$script:Personen = @()
$script:Begriffe = @()

# ------------------------------------------------------------------ Hilfsfunktionen

# Ersetzt {ae} {oe} {ue} {Ae} {Oe} {Ue} {ss} durch echte Umlaute (Datei bleibt reines ASCII)
function T([string]$s) {
    if (-not $s) { return '' }
    $s = $s.Replace('{ae}', [string][char]0x00E4).Replace('{oe}', [string][char]0x00F6)
    $s = $s.Replace('{ue}', [string][char]0x00FC).Replace('{Ae}', [string][char]0x00C4)
    $s = $s.Replace('{Oe}', [string][char]0x00D6).Replace('{Ue}', [string][char]0x00DC)
    $s = $s.Replace('{ss}', [string][char]0x00DF)
    return $s
}

function Write-LogDatei([string]$text) {
    if (-not $script:LogDatei) { return }
    try {
        $zeile = (Get-Date).ToString('dd.MM.yyyy HH:mm:ss') + '  ' + $text + "`r`n"
        [System.IO.File]::AppendAllText($script:LogDatei, $zeile, [System.Text.Encoding]::UTF8)
    } catch { }
}

function Log([string]$text, [string]$farbe = '') {
    $t = T $text
    if ($farbe) { Write-Host $t -ForegroundColor $farbe } else { Write-Host $t }
    Write-LogDatei $t
}

function Release($obj) {
    if ($null -eq $obj) { return }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
}

function Get-Endung([string]$name) {
    if (-not $name) { return '' }
    $i = $name.LastIndexOf('.')
    if ($i -lt 0 -or $i -eq $name.Length - 1) { return '' }
    $e = $name.Substring($i + 1).Trim().ToLowerInvariant()
    if ($e.Length -gt 6 -or $e -match '[^a-z0-9]') { return '' }
    return $e
}

function Get-SichererName([string]$s, [int]$max) {
    if (-not $s) { return 'unbekannt' }
    $s = [regex]::Replace($s, '[\\/:*?"<>|\x00-\x1f]', '_')
    $s = [regex]::Replace($s, '\s+', ' ').Trim().TrimEnd('.')
    if ($s.Length -gt $max) {
        $ext = ''
        $i = $s.LastIndexOf('.')
        if ($i -gt 0 -and ($s.Length - $i) -le 6) { $ext = $s.Substring($i); $s = $s.Substring(0, $i) }
        $laenge = [Math]::Max(1, $max - $ext.Length)
        if ($s.Length -gt $laenge) { $s = $s.Substring(0, $laenge) }
        $s = $s.TrimEnd(' ', '.') + $ext
    }
    if (-not $s) { return 'unbekannt' }
    return $s
}

function Get-FreierPfad([string]$ordner, [string]$datei) {
    $p = [System.IO.Path]::Combine($ordner, $datei)
    if (-not [System.IO.File]::Exists($p)) { return $p }
    $i = $datei.LastIndexOf('.')
    $basis = $datei; $ext = ''
    if ($i -gt 0) { $basis = $datei.Substring(0, $i); $ext = $datei.Substring($i) }
    for ($n = 2; $n -lt 10000; $n++) {
        $p = [System.IO.Path]::Combine($ordner, ($basis + '_' + $n + $ext))
        if (-not [System.IO.File]::Exists($p)) { return $p }
    }
    return [System.IO.Path]::Combine($ordner, ([guid]::NewGuid().ToString('N') + $ext))
}

function Read-Bytes([string]$pfad, [long]$max) {
    $fs = [System.IO.File]::Open($pfad, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $len = [int][Math]::Min([long]$fs.Length, $max)
        $buf = New-Object byte[] $len
        $pos = 0
        while ($pos -lt $len) {
            $r = $fs.Read($buf, $pos, $len - $pos)
            if ($r -le 0) { break }
            $pos += $r
        }
        if ($pos -lt $len) {
            $kurz = New-Object byte[] $pos
            [System.Array]::Copy($buf, $kurz, $pos)
            $buf = $kurz
        }
        return , $buf
    } finally { $fs.Dispose() }
}

# Text aus Bytes: erkennt UTF-16, UTF-8 (mit/ohne BOM), sonst Windows-1252 (ANSI)
function ConvertFrom-BytesSmart([byte[]]$b) {
    if ($null -eq $b -or $b.Length -eq 0) { return '' }
    if ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode.GetString($b, 2, $b.Length - 2) }
    if ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode.GetString($b, 2, $b.Length - 2) }
    $start = 0
    if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { $start = 3 }
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8.GetString($b, $start, $b.Length - $start)
    } catch {
        return [System.Text.Encoding]::GetEncoding(1252).GetString($b, $start, $b.Length - $start)
    }
}

# ------------------------------------------------------------------ Text vereinheitlichen

# Kleinschreibung, Umlaute -> ae/oe/ue/ss, Akzente weg, alles ausser a-z/0-9 -> Leerzeichen
function ConvertTo-Normal([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    if ($text.Length -gt 4000000) { $text = $text.Substring(0, 4000000) }
    $t = [regex]::Replace($text, '[\p{Cs}\p{Cc}\p{Co}\p{Cn}]', ' ')
    try { $t = $t.Normalize([System.Text.NormalizationForm]::FormC) } catch { }
    $t = $t.ToLowerInvariant()
    $t = $t.Replace($script:ChAe, 'ae').Replace($script:ChOe, 'oe').Replace($script:ChUe, 'ue').Replace($script:ChSz, 'ss')
    try { $t = $t.Normalize([System.Text.NormalizationForm]::FormD) } catch { }
    $t = [regex]::Replace($t, '\p{Mn}+', '')
    $t = [regex]::Replace($t, '[^a-z0-9]+', ' ')
    return $t.Trim()
}

# Schreibweisen eines Namens: "Loehr" und "Lohr" fuer Loehr mit Umlaut (Texterkennung verliert Punkte)
function Get-Varianten([string]$s) {
    $v = New-Object System.Collections.Generic.List[string]
    if (-not $s) { return , $v.ToArray() }
    $a = ConvertTo-Normal $s
    if ($a) { $v.Add($a) }
    $low = $s
    try { $low = $s.Normalize([System.Text.NormalizationForm]::FormC) } catch { }
    $low = $low.ToLowerInvariant().Replace($script:ChAe, 'a').Replace($script:ChOe, 'o').Replace($script:ChUe, 'u')
    $b = ConvertTo-Normal $low
    if ($b -and -not $v.Contains($b)) { $v.Add($b) }
    return , $v.ToArray()
}

function Join-Regex([string[]]$teile) {
    $esc = New-Object System.Collections.Generic.List[string]
    foreach ($t in $teile) { if ($t) { $esc.Add([regex]::Escape($t)) } }
    return '(?:' + ($esc -join '|') + ')'
}

function New-Person([string]$nachname, [string[]]$vornamen) {
    $nachV = Get-Varianten $nachname
    if ($nachV.Count -eq 0) { return $null }
    $vorV = New-Object System.Collections.Generic.List[string]
    $init = New-Object System.Collections.Generic.List[string]
    foreach ($tok in $vornamen) {
        $n = ConvertTo-Normal $tok
        if (-not $n) { continue }
        if (-not $init.Contains($n.Substring(0, 1))) { $init.Add($n.Substring(0, 1)) }
        foreach ($v in (Get-Varianten $tok)) {
            if ($v.Length -ge 2 -and -not $vorV.Contains($v)) { $vorV.Add($v) }
        }
    }
    if ($vorV.Count -eq 0) { return $null }
    $L = Join-Regex $nachV
    $F = Join-Regex $vorV.ToArray()
    $I = Join-Regex $init.ToArray()
    $opt = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $anrede = '(?:herr|herrn|frau|hr|fr|kollege|kollegin)'
    return [pscustomobject]@{
        Anzeige  = ($nachname + ', ' + ($vornamen -join ' '))
        Nachname = $nachname
        # nur der Nachname als ganzes Wort
        RxNach   = [regex]::new(('\b' + $L + '\b'), $opt)
        # ein Vorname als ganzes Wort (irgendwo im Text)
        RxVor    = [regex]::new(('\b' + $F + '\b'), $opt)
        # Vor- und Nachname direkt beieinander, hoechstens ein Wort dazwischen
        # ("Klatt, Andre", "Andre Klatt", "Klatt | Andre", andre.klatt@...)
        RxStark  = [regex]::new(('\b' + $L + ' (?:[a-z0-9]+ )?' + $F + '\b|\b' + $F + ' (?:[a-z0-9]+ )?' + $L + '\b'), $opt)
        # Nachname mit Anfangsbuchstabe ("Klatt, A.") oder Anrede ("Herr Klatt")
        RxMittel = [regex]::new(('\b' + $L + ' ' + $I + '\b|\b' + $I + ' ' + $L + '\b|\b' + $anrede + ' ' + $L + '\b'), $opt)
        # Nachname im Originaltext, um die Zeile im Dokument anzuzeigen
        RxRoh    = Get-RohMuster $nachname
    }
}

# Liest namen.txt: eine Person pro Zeile, "Nachname, Vorname" oder "Vorname Nachname".
# Aus Excel kopierte Zeilen gehen auch (es zaehlt nur die erste Spalte).
function Read-Namen([string]$pfad) {
    $liste = New-Object System.Collections.Generic.List[object]
    $gesehen = @{}
    $text = ConvertFrom-BytesSmart (Read-Bytes $pfad (5 * $script:MB))
    foreach ($roh in ($text -split "`r?`n")) {
        $zeile = $roh.Trim().TrimStart([char]0xFEFF)
        if (-not $zeile -or $zeile.StartsWith('#')) { continue }
        $zeile = ($zeile -split "`t")[0].Trim()
        $zeile = ($zeile -split ';')[0].Trim().Trim('"').Trim()
        if (-not $zeile) { continue }
        $nach = ''; $vorTeil = ''
        if ($zeile.Contains(',')) {
            $teile = $zeile.Split([char[]]@(','), 2)
            $nach = $teile[0].Trim(); $vorTeil = $teile[1].Trim()
        } else {
            $w = @($zeile -split '\s+' | Where-Object { $_ })
            if ($w.Count -ge 2) { $nach = $w[$w.Count - 1]; $vorTeil = ($w[0..($w.Count - 2)] -join ' ') }
        }
        # Vornamen: nur Woerter aus Buchstaben (bricht z. B. bei "V.IW-MI-N-FFM" ab)
        $vor = New-Object System.Collections.Generic.List[string]
        foreach ($tok in ($vorTeil -split '\s+')) {
            if (-not $tok) { continue }
            if ($tok -notmatch '^[\p{L}][\p{L}\-]*$') { break }
            $vor.Add($tok)
            if ($vor.Count -ge 3) { break }
        }
        if (-not $nach -or $nach -match '\d' -or $vor.Count -eq 0) {
            Log ('  Zeile in der Namensliste nicht verstanden, wird ignoriert: ' + $roh.Trim()) 'Yellow'
            continue
        }
        $p = New-Person $nach $vor.ToArray()
        if ($null -eq $p) { continue }
        if ($gesehen.ContainsKey($p.Anzeige)) { continue }
        $gesehen[$p.Anzeige] = $true
        $liste.Add($p)
    }
    return , $liste.ToArray()
}

function New-NamenVorlage([string]$pfad) {
    $zeilen = @(
        '# Bitte hier die Mitarbeiter eintragen - eine Person pro Zeile, so:  Nachname, Vorname',
        '# Zeilen, die mit # beginnen, werden ignoriert.',
        '# Ganze Zeilen aus einer Excel-Liste gehen auch: verwendet wird nur die erste Spalte.',
        '#',
        '# Beispiel:',
        '# Mustermann, Max',
        ''
    )
    [System.IO.File]::WriteAllText($pfad, ($zeilen -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

function Initialize-Begriffe([string[]]$liste) {
    $b = New-Object System.Collections.Generic.List[object]
    foreach ($w in $liste) {
        $n = ConvertTo-Normal $w
        if ($n) { $b.Add([pscustomobject]@{ Anzeige = $w; Norm = $n }) }
    }
    $script:Begriffe = $b.ToArray()
}

# Welche UVV-Begriffe kommen im (vereinheitlichten) Text vor?
function Find-Begriffe([string]$norm) {
    $treffer = New-Object System.Collections.Generic.List[string]
    if ($norm) {
        foreach ($b in $script:Begriffe) {
            if ($norm.Contains($b.Norm) -and -not $treffer.Contains($b.Anzeige)) { $treffer.Add($b.Anzeige) }
        }
    }
    return , $treffer.ToArray()
}

# Welche Personen kommen vor? Ergebnis: Hashtable Anzeige -> Staerke
#   3 = Vor- und Nachname beieinander, 2 = beide im Text / "Klatt, A." / "Herr Klatt",
#   1 = nur Nachname (nur wenn $nurNachnameErlaubt, z. B. fuer Betreff und Dateiname)
function Find-Personen([string]$norm, [bool]$nurNachnameErlaubt) {
    $res = @{}
    if (-not $norm -or $norm.Length -lt 3) { return $res }
    foreach ($p in $script:Personen) {
        if (-not $p.RxNach.IsMatch($norm)) { continue }
        $s = 1
        if ($p.RxStark.IsMatch($norm)) { $s = 3 }
        elseif ($p.RxVor.IsMatch($norm) -or $p.RxMittel.IsMatch($norm)) { $s = 2 }
        if ($s -ge 2 -or $nurNachnameErlaubt) { $res[$p.Anzeige] = $s }
    }
    return $res
}

# ------------------------------------------------------------------ Text aus Dateien lesen

try { Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop } catch { }
try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }

# Word/Excel/PowerPoint (docx, xlsx, pptx) und OpenDocument: XML-Teile aus dem ZIP lesen
function Get-ZipXmlText([string]$pfad, [string[]]$muster) {
    $sb = New-Object System.Text.StringBuilder
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($pfad)
        foreach ($e in $zip.Entries) {
            $passt = $false
            foreach ($m in $muster) { if ($e.FullName -like $m) { $passt = $true; break } }
            if (-not $passt -or $e.Length -gt (60 * $script:MB)) { continue }
            $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
            try { $xml = $sr.ReadToEnd() } finally { $sr.Dispose() }
            [void]$sb.Append((Convert-XmlZuText $xml))
            [void]$sb.Append("`n")
        }
    } finally { if ($null -ne $zip) { $zip.Dispose() } }
    return $sb.ToString()
}

# Alte Office-Formate (doc, xls, ppt) und Notfall fuer andere: Text grob aus den Bytes holen
function Get-BinaerText([string]$pfad) {
    $b = Read-Bytes $pfad (8 * $script:MB)
    if ($b.Length -lt 2) { return '' }
    $ansi = [System.Text.Encoding]::GetEncoding(1252).GetString($b)
    $u0 = [System.Text.Encoding]::Unicode.GetString($b, 0, $b.Length - ($b.Length % 2))
    $rest = $b.Length - 1
    $u1 = [System.Text.Encoding]::Unicode.GetString($b, 1, $rest - ($rest % 2))
    return ($ansi + ' ' + $u0 + ' ' + $u1)
}

function Get-RtfText([string]$pfad) {
    $roh = [System.Text.Encoding]::GetEncoding(1252).GetString((Read-Bytes $pfad (20 * $script:MB)))
    # \uNNNN mit Ersatzzeichen zuerst, dann \'xx
    $roh = [regex]::Replace($roh, "\\u(-?\d+)(?:\\'[0-9a-fA-F]{2}|\?)?", [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $n = [int]$m.Groups[1].Value
            if ($n -lt 0) { $n += 65536 }
            return [string][char]$n
        })
    $roh = [regex]::Replace($roh, "\\'([0-9a-fA-F]{2})", [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            return [System.Text.Encoding]::GetEncoding(1252).GetString([byte[]]@([System.Convert]::ToByte($m.Groups[1].Value, 16)))
        })
    return $roh
}

function Get-PlainText([string]$pfad, [string]$ext) {
    $t = ConvertFrom-BytesSmart (Read-Bytes $pfad (20 * $script:MB))
    if ($ext -eq 'eml') {
        # quoted-printable (z. B. M=C3=BCnch) grob dekodieren
        $t = [regex]::Replace($t, "=\r?\n", '')
        $t = [regex]::Replace($t, '(?:=[0-9A-Fa-f]{2})+', [System.Text.RegularExpressions.MatchEvaluator] {
                param($m)
                $hex = $m.Value.Replace('=', '')
                $bytes = New-Object byte[] ([int]($hex.Length / 2))
                for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [System.Convert]::ToByte($hex.Substring($i * 2, 2), 16) }
                try { return (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes) }
                catch { return [System.Text.Encoding]::GetEncoding(1252).GetString($bytes) }
            })
    }
    if ($ext -eq 'htm' -or $ext -eq 'html' -or $ext -eq 'eml') {
        $t = [regex]::Replace($t, '<[^>]+>', ' ')
        $t = [System.Net.WebUtility]::HtmlDecode($t)
    }
    return $t
}

# ZIP-Archiv: enthaltene Dokumente auspacken und lesen
function Get-ZipInhaltText([string]$pfad, [int]$tiefe) {
    $sb = New-Object System.Text.StringBuilder
    $ordner = [System.IO.Path]::Combine($script:TmpDir, ('zip_' + [guid]::NewGuid().ToString('N')))
    [void][System.IO.Directory]::CreateDirectory($ordner)
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($pfad)
        $n = 0
        foreach ($e in $zip.Entries) {
            if ($n -ge 100) { break }
            if (-not $e.Name) { continue }
            $ext = Get-Endung $e.Name
            if (-not $script:DokSet.ContainsKey($ext) -or $e.Length -gt ($MaxAnhangMB * $script:MB)) { continue }
            $n++
            $ziel = [System.IO.Path]::Combine($ordner, ($n.ToString() + '.' + $ext))
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, $ziel, $true)
            [void]$sb.AppendLine($e.Name)
            $r = Get-DokumentText $ziel $ext ($tiefe + 1)
            [void]$sb.AppendLine($r.Text)
        }
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
        try { [System.IO.Directory]::Delete($ordner, $true) } catch { }
    }
    return $sb.ToString()
}

# Angehaengte Mail (.msg): Betreff, Text und deren Anhaenge lesen (ueber Outlook)
function Get-MsgText([string]$pfad, [int]$tiefe) {
    $m = $script:Ns.OpenSharedItem($pfad)
    $sb = New-Object System.Text.StringBuilder
    try {
        try { [void]$sb.AppendLine([string]$m.Subject) } catch { }
        try { [void]$sb.AppendLine([string]$m.SenderName) } catch { }
        try { [void]$sb.AppendLine([string]$m.Body) } catch { }
        if ($tiefe -lt 2) {
            $atts = $null
            try { $atts = $m.Attachments } catch { }
            $na = 0
            if ($null -ne $atts) { try { $na = [int]$atts.Count } catch { } }
            for ($k = 1; $k -le $na; $k++) {
                $a = $null
                try {
                    $a = $atts.Item($k)
                    $name = [string]$a.FileName
                    $ext = Get-Endung $name
                    if ([int]$a.Type -eq 5) { $ext = 'msg' }
                    if (-not $script:DokSet.ContainsKey($ext)) { continue }
                    $groesse = [long]$a.Size
                    if ($script:BildSet.ContainsKey($ext) -and $groesse -lt ($MinBildKB * $script:KB)) { continue }
                    if ($groesse -gt ($MaxAnhangMB * $script:MB)) { continue }
                    $tmp = [System.IO.Path]::Combine($script:TmpDir, ([guid]::NewGuid().ToString('N') + '.' + $ext))
                    $a.SaveAsFile($tmp)
                    try {
                        [void]$sb.AppendLine($name)
                        $r = Get-DokumentText $tmp $ext ($tiefe + 1)
                        [void]$sb.AppendLine($r.Text)
                    } finally { try { [System.IO.File]::Delete($tmp) } catch { } }
                } catch {
                    Write-LogDatei ('Anhang in ' + $pfad + ' nicht lesbar: ' + $_.Exception.Message)
                } finally { Release $a }
            }
            Release $atts
        }
    } finally {
        try { $m.Close(1) } catch { }
        Release $m
    }
    return $sb.ToString()
}

# Liest den Text einer Datei. Ergebnis: @{ Text; Quelle; Hinweis }
function Get-DokumentText([string]$pfad, [string]$ext, [int]$tiefe) {
    $r = @{ Text = ''; Quelle = ''; Hinweis = '' }
    if ($ext -eq 'xlsx' -or $ext -eq 'xlsm') {
        $r.Text = Get-XlsxText $pfad; $r.Quelle = 'Office-Datei'
    } elseif ($script:ZipXmlTeile.ContainsKey($ext)) {
        $r.Text = Get-ZipXmlText $pfad $script:ZipXmlTeile[$ext]; $r.Quelle = 'Office-Datei'
    } elseif ($script:BinaerSet.ContainsKey($ext)) {
        $r.Text = Get-BinaerText $pfad; $r.Quelle = 'Office-Datei (altes Format)'
    } elseif ($ext -eq 'rtf') {
        $r.Text = Get-RtfText $pfad; $r.Quelle = 'RTF'
    } elseif ($script:TextSet.ContainsKey($ext)) {
        $r.Text = Get-PlainText $pfad $ext; $r.Quelle = 'Text'
    } elseif ($ext -eq 'pdf') {
        if ($script:OcrAktiv) { $r.Text = Get-PdfOcrText $pfad $OcrMaxSeiten; $r.Quelle = 'Texterkennung' }
        else { $r.Quelle = 'nicht gelesen'; $r.Hinweis = $script:OcrGrund }
    } elseif ($script:BildSet.ContainsKey($ext)) {
        if ($script:OcrAktiv) { $r.Text = Get-BildOcrText $pfad $OcrMaxSeiten; $r.Quelle = 'Texterkennung' }
        else { $r.Quelle = 'nicht gelesen'; $r.Hinweis = $script:OcrGrund }
    } elseif ($ext -eq 'msg') {
        if ($tiefe -ge 2) {
            $r.Text = Get-BinaerText $pfad; $r.Quelle = 'Mail-Anhang (grob)'
        } else {
            try { $r.Text = Get-MsgText $pfad $tiefe; $r.Quelle = 'Mail-Anhang' }
            catch { $r.Text = Get-BinaerText $pfad; $r.Quelle = 'Mail-Anhang (grob)'; $r.Hinweis = $_.Exception.Message }
        }
    } elseif ($ext -eq 'zip') {
        if ($tiefe -ge 2) { $r.Quelle = 'nicht gelesen' }
        else { $r.Text = Get-ZipInhaltText $pfad $tiefe; $r.Quelle = 'ZIP-Inhalt' }
    } else {
        $r.Quelle = 'nicht gelesen'
    }
    if ($null -eq $r.Text) { $r.Text = '' }
    return $r
}

# ------------------------------------------------------------------ Texterkennung (OCR)
# Nutzt die in Windows 10/11 eingebaute Texterkennung (Windows.Media.Ocr) und den
# eingebauten PDF-Renderer (Windows.Data.Pdf). Funktioniert nur in Windows PowerShell 5.1.

$script:OcrAktiv = $false
$script:OcrGrund = 'Texterkennung nicht gestartet'
$script:Ocr = $null
$script:OcrMax = 2600
$script:AsTaskOp = $null
$script:AsTaskAkt = $null

function Wait-WinRt($vorgang, [type]$typ) {
    $task = $script:AsTaskOp.MakeGenericMethod($typ).Invoke($null, @($vorgang))
    if (-not $task.Wait(120000)) { throw 'Zeitueberschreitung bei der Texterkennung' }
    return , $task.Result
}

function Wait-WinRtAktion($vorgang) {
    $task = $script:AsTaskAkt.Invoke($null, @($vorgang))
    if (-not $task.Wait(120000)) { throw 'Zeitueberschreitung bei der Texterkennung' }
}

function Initialize-Ocr {
    $script:OcrAktiv = $false
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') { throw 'Bitte mit "Windows PowerShell" starten (nicht PowerShell 7).' }
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapFrame, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapTransform, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapPixelFormat, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapAlphaMode, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.ExifOrientationMode, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.ColorManagementMode, Windows.Graphics, ContentType = WindowsRuntime]
        $null = [Windows.Media.Ocr.OcrEngine, Windows.Media, ContentType = WindowsRuntime]
        $null = [Windows.Media.Ocr.OcrResult, Windows.Media, ContentType = WindowsRuntime]
        $null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]

        foreach ($mi in [System.WindowsRuntimeSystemExtensions].GetMethods()) {
            if ($mi.Name -ne 'AsTask') { continue }
            $ps = $mi.GetParameters()
            if ($ps.Count -ne 1) { continue }
            $pn = $ps[0].ParameterType.Name
            if ($pn -eq 'IAsyncOperation`1' -and $null -eq $script:AsTaskOp) { $script:AsTaskOp = $mi }
            if ($pn -eq 'IAsyncAction' -and $null -eq $script:AsTaskAkt) { $script:AsTaskAkt = $mi }
        }
        if ($null -eq $script:AsTaskOp -or $null -eq $script:AsTaskAkt) { throw 'WinRT-Hilfsfunktionen nicht gefunden' }

        $engine = $null
        foreach ($sprache in @('de-DE', 'de')) {
            try {
                $lang = [Windows.Globalization.Language]::new($sprache)
                if ([Windows.Media.Ocr.OcrEngine]::IsLanguageSupported($lang)) {
                    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
                    if ($null -ne $engine) { break }
                }
            } catch { }
        }
        if ($null -eq $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
        if ($null -eq $engine) { throw 'Keine Sprache fuer die Windows-Texterkennung installiert' }
        $script:Ocr = $engine
        try { $script:OcrMax = [int][Windows.Media.Ocr.OcrEngine]::MaxImageDimension } catch { $script:OcrMax = 2600 }
        if ($script:OcrMax -lt 1000) { $script:OcrMax = 2600 }
        $script:OcrAktiv = $true
        $script:OcrGrund = ''
        return $true
    } catch {
        $script:OcrAktiv = $false
        $script:OcrGrund = 'Texterkennung nicht verfuegbar: ' + $_.Exception.Message
        return $false
    }
}

# Erkannte Textstuecke auf gleicher Hoehe (eine Tabellenzeile) zu einer Zeile zusammenfassen:
# "Klatt, Andre | 12.12.2025 | ..." statt einzelner Zeilen pro Spalte.
function Invoke-OcrBitmap($bmp) {
    $erg = Wait-WinRt ($script:Ocr.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
    $teile = New-Object System.Collections.Generic.List[object]
    foreach ($z in $erg.Lines) {
        $x1 = [double]::MaxValue; $x2 = 0.0; $ySumme = 0.0; $hSumme = 0.0; $n = 0
        foreach ($w in $z.Words) {
            $r = $w.BoundingRect
            $x1 = [Math]::Min($x1, [double]$r.X)
            $x2 = [Math]::Max($x2, [double]$r.X + [double]$r.Width)
            $ySumme += [double]$r.Y + [double]$r.Height / 2
            $hSumme += [double]$r.Height
            $n++
        }
        if ($n -eq 0) { continue }
        $teile.Add([pscustomobject]@{ X1 = $x1; X2 = $x2; Y = $ySumme / $n; H = $hSumme / $n; Text = [string]$z.Text })
    }
    if ($teile.Count -eq 0) { return '' }
    $hoehen = @($teile | ForEach-Object { $_.H } | Sort-Object)
    $median = [double]$hoehen[[int][Math]::Floor($hoehen.Count / 2)]
    $toleranz = [Math]::Max(3.0, 0.6 * $median)
    $zeilen = New-Object System.Collections.Generic.List[string]
    $gruppe = New-Object System.Collections.Generic.List[object]
    $yGruppe = 0.0
    foreach ($t in @($teile | Sort-Object Y)) {
        if ($gruppe.Count -gt 0 -and [Math]::Abs($t.Y - $yGruppe) -gt $toleranz) {
            $zeilen.Add((Join-OcrZeile $gruppe $median))
            $gruppe.Clear()
        }
        $gruppe.Add($t)
        $summe = 0.0
        foreach ($g in $gruppe) { $summe += $g.Y }
        $yGruppe = $summe / $gruppe.Count
    }
    if ($gruppe.Count -gt 0) { $zeilen.Add((Join-OcrZeile $gruppe $median)) }
    return ($zeilen -join "`n")
}

function Join-OcrZeile($gruppe, [double]$hoehe) {
    $sb = New-Object System.Text.StringBuilder
    $ende = $null
    foreach ($t in @($gruppe | Sort-Object X1)) {
        if ($null -ne $ende) {
            if (($t.X1 - $ende) -gt 2.0 * $hoehe) { [void]$sb.Append(' | ') } else { [void]$sb.Append(' ') }
        }
        [void]$sb.Append($t.Text)
        $ende = $t.X2
    }
    return $sb.ToString()
}

# Skalierung: so gross wie sinnvoll (bessere Erkennung), aber unter der OCR-Grenze und max. ~12 Megapixel
function Get-Skalierung([double]$w, [double]$h, [double]$maxFaktor) {
    if ($w -lt 1 -or $h -lt 1) { return 0.0 }
    $f = [Math]::Min($maxFaktor, ($script:OcrMax - 16) / [Math]::Max($w, $h))
    $f = [Math]::Min($f, [Math]::Sqrt(12000000.0 / ($w * $h)))
    return $f
}

function Get-PdfOcrText([string]$pfad, [int]$maxSeiten) {
    $datei = Wait-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pfad)) ([Windows.Storage.StorageFile])
    $pdf = Wait-WinRt ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($datei)) ([Windows.Data.Pdf.PdfDocument])
    $sb = New-Object System.Text.StringBuilder
    $anzahl = [Math]::Min([int]$pdf.PageCount, $maxSeiten)
    for ($i = 0; $i -lt $anzahl; $i++) {
        $seite = $pdf.GetPage([uint32]$i)
        try {
            $w = [double]$seite.Size.Width
            $h = [double]$seite.Size.Height
            $f = Get-Skalierung $w $h 3.0
            if ($f -le 0) { continue }
            $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $opt.DestinationWidth = [uint32][Math]::Max(1, [Math]::Floor($w * $f))
            $opt.DestinationHeight = [uint32][Math]::Max(1, [Math]::Floor($h * $f))
            $strom = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            try {
                Wait-WinRtAktion ($seite.RenderToStreamAsync($strom, $opt))
                $strom.Seek(0)
                $dec = Wait-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($strom)) ([Windows.Graphics.Imaging.BitmapDecoder])
                $bmp = Wait-WinRt ($dec.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8, [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)) ([Windows.Graphics.Imaging.SoftwareBitmap])
                try { [void]$sb.AppendLine((Invoke-OcrBitmap $bmp)) } finally { try { $bmp.Dispose() } catch { } }
            } finally { try { $strom.Dispose() } catch { } }
        } finally { try { $seite.Dispose() } catch { } }
    }
    try { $pdf.Dispose() } catch { }
    return $sb.ToString()
}

function Get-BildOcrText([string]$pfad, [int]$maxSeiten) {
    $datei = Wait-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pfad)) ([Windows.Storage.StorageFile])
    $strom = Wait-WinRt ($datei.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $dec = Wait-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($strom)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $sb = New-Object System.Text.StringBuilder
        $anzahl = [Math]::Min([int]$dec.FrameCount, $maxSeiten)
        for ($i = 0; $i -lt $anzahl; $i++) {
            $frame = Wait-WinRt ($dec.GetFrameAsync([uint32]$i)) ([Windows.Graphics.Imaging.BitmapFrame])
            $w = [double]$frame.PixelWidth
            $h = [double]$frame.PixelHeight
            $f = [Math]::Min(1.0, (Get-Skalierung $w $h 1.0))
            if ($f -le 0) { continue }
            $tf = New-Object Windows.Graphics.Imaging.BitmapTransform
            $tf.ScaledWidth = [uint32][Math]::Max(1, [Math]::Floor($w * $f))
            $tf.ScaledHeight = [uint32][Math]::Max(1, [Math]::Floor($h * $f))
            $bmp = Wait-WinRt ($frame.GetSoftwareBitmapAsync(
                    [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
                    [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied,
                    $tf,
                    [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
                    [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage)) ([Windows.Graphics.Imaging.SoftwareBitmap])
            try { [void]$sb.AppendLine((Invoke-OcrBitmap $bmp)) } finally { try { $bmp.Dispose() } catch { } }
        }
        return $sb.ToString()
    } finally { try { $strom.Dispose() } catch { } }
}

# Kleine Test-PDF erzeugen (Text "UVV Selbsttest"), um die Texterkennung vorab zu pruefen
function New-TestPdf([string]$pfad) {
    $inhalt = 'BT /F1 40 Tf 20 40 Td (UVV Selbsttest) Tj ET'
    $objekte = @(
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 420 120] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
        ('<< /Length ' + $inhalt.Length + " >>`nstream`n" + $inhalt + "`nendstream"),
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n")
    $pos = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $objekte.Count; $i++) {
        $pos.Add($sb.Length)
        [void]$sb.Append(($i + 1).ToString() + " 0 obj`n" + $objekte[$i] + "`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 " + ($objekte.Count + 1) + "`n0000000000 65535 f `n")
    foreach ($p in $pos) { [void]$sb.Append($p.ToString('0000000000') + " 00000 n `n") }
    [void]$sb.Append("trailer`n<< /Size " + ($objekte.Count + 1) + " /Root 1 0 R >>`nstartxref`n" + $xref + "`n%%EOF`n")
    [System.IO.File]::WriteAllText($pfad, $sb.ToString(), [System.Text.Encoding]::ASCII)
}

function Test-Ocr {
    $pfad = [System.IO.Path]::Combine($script:TmpDir, 'selbsttest.pdf')
    try {
        New-TestPdf $pfad
        $text = ConvertTo-Normal (Get-PdfOcrText $pfad 1)
        if ($text.Contains('selbsttest') -or $text.Contains('uvv')) { return $true }
        $script:OcrGrund = 'Texterkennung liefert keinen Text (Test ergab: "' + $text + '")'
    } catch {
        $script:OcrGrund = 'Texterkennung fehlgeschlagen: ' + $_.Exception.Message
    } finally { try { [System.IO.File]::Delete($pfad) } catch { } }
    $script:OcrAktiv = $false
    return $false
}

# ------------------------------------------------------------------ Outlook durchsuchen

function Initialize-Status {
    $script:Dateien = @{}         # SHA256 -> gespeichertes Dokument
    $script:SchonGesehen = @{}    # "dateiname|groesse" -> SHA256 (spart doppeltes Speichern)
    $script:Funde = New-Object System.Collections.Generic.List[object]
    $script:PostfachInfo = New-Object System.Collections.Generic.List[object]
    $script:Stat = @{ Mails = 0; MailsMitAnhang = 0; Anhaenge = 0; Dokumente = 0; Doppelt = 0
        Verschluesselt = 0; Fehler = 0; Uebersprungen = 0 }
}

function Connect-Outlook {
    $ol = $null
    try { $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop }
    catch {
        $grund = $_.Exception.Message
        throw ((T 'Outlook konnte nicht ge{oe}ffnet werden. Das Skript braucht das klassische Outlook (nicht "Neues Outlook"). Technischer Grund: ') + $grund)
    }
    $script:Outlook = $ol
    $script:Ns = $ol.GetNamespace('MAPI')
}

function Get-TypText([int]$typ) {
    switch ($typ) {
        0 { return 'eigenes Postfach' }
        1 { return 'freigegebenes Postfach' }
        2 { return '{Oe}ffentliche Ordner' }
        3 { return 'Datendatei' }
        4 { return 'zus{ae}tzliches Postfach' }
        default { return '' }
    }
}

function Get-PostfachListe {
    $liste = New-Object System.Collections.Generic.List[object]
    $nr = 0
    foreach ($st in $script:Ns.Stores) {
        $nr++
        $name = ''; try { $name = [string]$st.DisplayName } catch { }
        $typ = -1; try { $typ = [int]$st.ExchangeStoreType } catch { }
        $liste.Add([pscustomobject]@{ Nr = $nr; Name = $name; Typ = $typ; Store = $st; Root = $null })
    }
    foreach ($adresse in $ZusaetzlichePostfaecher) {
        if (-not $adresse) { continue }
        try {
            $empf = $script:Ns.CreateRecipient($adresse)
            [void]$empf.Resolve()
            $posteingang = $script:Ns.GetSharedDefaultFolder($empf, 6)
            $root = $null
            try { $root = $posteingang.Store.GetRootFolder() } catch { }
            if ($null -eq $root) {
                try { $p = $posteingang.Parent; if ([int]$p.Class -eq 2) { $root = $p } } catch { }
            }
            if ($null -eq $root) { $root = $posteingang }
            $nr++
            $liste.Add([pscustomobject]@{ Nr = $nr; Name = $adresse; Typ = 1; Store = $null; Root = $root })
        } catch {
            Log ('  Postfach ' + $adresse + ' nicht erreichbar: ' + $_.Exception.Message) 'Yellow'
        }
    }
    return , $liste
}

function Get-AlleOrdner($root) {
    $liste = New-Object System.Collections.Generic.List[object]
    $stapel = New-Object System.Collections.Stack
    $stapel.Push($root)
    while ($stapel.Count -gt 0) {
        $o = $stapel.Pop()
        $name = ''; try { $name = [string]$o.Name } catch { }
        if ($script:IgnorierOrdner.ContainsKey($name)) { continue }
        $liste.Add($o)
        $unter = New-Object System.Collections.Generic.List[object]
        try { foreach ($u in $o.Folders) { $unter.Add($u) } } catch { }
        for ($i = $unter.Count - 1; $i -ge 0; $i--) { $stapel.Push($unter[$i]) }
    }
    return , $liste
}

function Get-MailDatum($it) {
    foreach ($feld in @('ReceivedTime', 'SentOn', 'CreationTime')) {
        try {
            $d = [datetime]$it.$feld
            if ($d.Year -gt 1950 -and $d.Year -lt 4500) { return $d }
        } catch { }
    }
    return [datetime]::MinValue
}

function Get-Naechste($auswahl) {
    for ($v = 0; $v -lt 3; $v++) {
        try { return $auswahl.GetNext() } catch { Start-Sleep -Milliseconds 500 }
    }
    Write-LogDatei 'Outlook hat das naechste Element nicht geliefert - Rest des Ordners uebersprungen.'
    return $null
}

# Filter fuer Mails OHNE Anhang, deren Betreff oder Text einen Suchbegriff enthaelt
function Get-StichwortFilter {
    $teile = New-Object System.Collections.Generic.List[string]
    foreach ($b in $Suchbegriffe) {
        if (-not $b) { continue }
        $w = $b.Replace("'", "''")
        $teile.Add(('"urn:schemas:httpmail:subject" LIKE ''%{0}%''' -f $w))
        $teile.Add(('"urn:schemas:httpmail:textdescription" LIKE ''%{0}%''' -f $w))
    }
    return '@SQL="urn:schemas:httpmail:hasattachment" = 0 AND (' + ($teile -join ' OR ') + ')'
}

# Speichert einen Anhang (falls Dokument). Ergebnis: @{Hash; Name} oder $null
function Save-Anhang($att, [datetime]$datum, [string]$absender) {
    $typ = 1; try { $typ = [int]$att.Type } catch { }
    if ($typ -eq 6) { return $null }
    $name = ''; try { $name = [string]$att.FileName } catch { }
    if (-not $name) { try { $name = [string]$att.DisplayName } catch { } }
    if (-not $name) { $name = 'Anhang' }
    $ext = Get-Endung $name
    if ($typ -eq 5 -and $ext -ne 'msg') { $name = $name + '.msg'; $ext = 'msg' }
    if (-not $script:DokSet.ContainsKey($ext)) { return $null }
    $groesse = [long]0; try { $groesse = [long]$att.Size } catch { }
    if ($script:BildSet.ContainsKey($ext) -and $groesse -gt 0 -and $groesse -lt ($MinBildKB * $script:KB)) { return $null }
    if ($groesse -gt ($MaxAnhangMB * $script:MB)) {
        $script:Stat.Uebersprungen++
        Write-LogDatei ('Zu gross, uebersprungen: ' + $name + ' (' + [int]($groesse / $script:MB) + ' MB)')
        return $null
    }
    $script:Stat.Anhaenge++
    $schluessel = $name.ToLowerInvariant() + '|' + $groesse
    if ($script:SchonGesehen.ContainsKey($schluessel)) {
        $hash = $script:SchonGesehen[$schluessel]
        $vorhanden = $script:Dateien[$hash]
        if ($datum -gt $vorhanden.Datum) { $vorhanden.Datum = $datum }
        $script:Stat.Doppelt++
        return [pscustomobject]@{ Hash = $hash; Name = $name }
    }
    $tmp = [System.IO.Path]::Combine($script:TmpDir, ([guid]::NewGuid().ToString('N') + '.' + $ext))
    $att.SaveAsFile($tmp)
    if (-not [System.IO.File]::Exists($tmp)) { throw 'Anhang wurde nicht gespeichert' }
    $hash = (Get-FileHash -LiteralPath $tmp -Algorithm SHA256).Hash
    $script:SchonGesehen[$schluessel] = $hash
    if ($script:Dateien.ContainsKey($hash)) {
        try { [System.IO.File]::Delete($tmp) } catch { }
        $vorhanden = $script:Dateien[$hash]
        if ($datum -gt $vorhanden.Datum) { $vorhanden.Datum = $datum }
        $script:Stat.Doppelt++
        return [pscustomobject]@{ Hash = $hash; Name = $name }
    }
    $zielName = $datum.ToString('yyyy-MM-dd') + '_' + (Get-SichererName $absender 25) + '_' + (Get-SichererName $name 90)
    $ziel = Get-FreierPfad $script:AnhangDir $zielName
    [System.IO.File]::Move($tmp, $ziel)
    $script:Stat.Dokumente++
    $script:Dateien[$hash] = [pscustomobject]@{
        Hash = $hash; Pfad = $ziel; Datei = [System.IO.Path]::GetFileName($ziel); Endung = $ext
        Groesse = (New-Object System.IO.FileInfo($ziel)).Length; Datum = $datum; Signal = $false
        Status = 'offen'; Quelle = ''; Hinweis = ''; Zeichen = 0; Begriffe = @(); Namen = @{}
        Kontext = @{}; Anzahl = 0; DokArt = ''
    }
    return [pscustomobject]@{ Hash = $hash; Name = $name }
}

function Add-Fund([string]$art, [string]$hash, [string]$anhang, [datetime]$datum, [string]$absender,
    [string]$betreff, [string]$postfach, [string]$ordner, $mailBegriffe, $mailNamen, $dateiBegriffe, $dateiNamen, $mailKontext) {
    if ($null -eq $mailKontext) { $mailKontext = @{} }
    if ($null -eq $mailBegriffe) { $mailBegriffe = @() }
    if ($null -eq $mailNamen) { $mailNamen = @{} }
    if ($null -eq $dateiBegriffe) { $dateiBegriffe = @() }
    if ($null -eq $dateiNamen) { $dateiNamen = @{} }
    $script:Funde.Add([pscustomobject]@{
            Art = $art; Hash = $hash; Anhang = $anhang; Datum = $datum; Absender = $absender; Betreff = $betreff
            Postfach = $postfach; Ordner = $ordner; MailBegriffe = $mailBegriffe; MailNamen = $mailNamen
            DateiBegriffe = $dateiBegriffe; DateiNamen = $dateiNamen; MailKontext = $mailKontext
        })
}

# Eine Mail pruefen. Rueckgabe $true = Startdatum unterschritten (Ordner ist fertig)
function Search-Mail($it, [string]$postfach, [string]$pfad, [datetime]$seit, [bool]$sortiert, [bool]$mitAnhaengen) {
    $rt = [datetime]::MaxValue
    try { $rt = [datetime]$it.ReceivedTime } catch { }
    if ($sortiert -and $rt.Year -lt 4500 -and $rt -lt $seit) { return $true }
    $datum = Get-MailDatum $it
    if ($datum -lt $seit) { return $false }
    $script:Stat.Mails++
    $klasse = ''; try { $klasse = [string]$it.MessageClass } catch { }
    $betreff = ''; try { $betreff = [string]$it.Subject } catch { }
    $absender = ''; try { $absender = [string]$it.SenderName } catch { }
    $nBetreff = ConvertTo-Normal $betreff
    $betreffNamen = Find-Personen $nBetreff $true

    # Verschluesselte Mails nicht oeffnen (wuerde nach Karte/PIN fragen) - nur Betreff pruefen
    if ($klasse -ieq 'IPM.Note.SMIME') {
        $script:Stat.Verschluesselt++
        $kw = Find-Begriffe $nBetreff
        if ($kw.Count -gt 0 -or $betreffNamen.Count -gt 0) {
            Add-Fund 'Mail verschluesselt' '' '' $datum $absender $betreff $postfach $pfad $kw $betreffNamen $null $null (Get-Kontext $betreff $betreffNamen)
        }
        return $false
    }

    $text = ''; try { $text = [string]$it.Body } catch { }
    if ($text.Length -gt 30000) { $text = $text.Substring(0, 30000) }
    $nMail = ConvertTo-Normal ($betreff + ' ' + $absender + ' ' + $text)
    $mailBegriffe = Find-Begriffe $nMail
    $mailNamen = Find-Personen $nMail $false
    foreach ($k in @($betreffNamen.Keys)) {
        if (-not $mailNamen.ContainsKey($k) -or $mailNamen[$k] -lt $betreffNamen[$k]) { $mailNamen[$k] = $betreffNamen[$k] }
    }
    $mailKontext = Get-Kontext ($betreff + "`n" + $text) $mailNamen

    $dokumente = 0
    if ($mitAnhaengen) {
        $script:Stat.MailsMitAnhang++
        $atts = $null
        try { $atts = $it.Attachments } catch { }
        $na = 0
        if ($null -ne $atts) { try { $na = [int]$atts.Count } catch { } }
        for ($k = 1; $k -le $na; $k++) {
            $att = $null
            try {
                $att = $atts.Item($k)
                $erg = Save-Anhang $att $datum $absender
                if ($null -ne $erg) {
                    $dokumente++
                    $nDatei = ConvertTo-Normal $erg.Name
                    $dBegriffe = Find-Begriffe $nDatei
                    $dNamen = Find-Personen $nDatei $true
                    Add-Fund 'Anhang' $erg.Hash $erg.Name $datum $absender $betreff $postfach $pfad $mailBegriffe $mailNamen $dBegriffe $dNamen $mailKontext
                    if ($mailBegriffe.Count -gt 0 -or $mailNamen.Count -gt 0 -or $dBegriffe.Count -gt 0 -or $dNamen.Count -gt 0) {
                        $script:Dateien[$erg.Hash].Signal = $true
                    }
                }
            } catch {
                $script:Stat.Fehler++
                Write-LogDatei ('Anhang nicht gespeichert (Mail "' + $betreff + '"): ' + $_.Exception.Message)
            } finally { Release $att }
        }
        Release $atts
    }
    if ($dokumente -eq 0 -and $mailBegriffe.Count -gt 0) {
        Add-Fund 'Mail' '' '' $datum $absender $betreff $postfach $pfad $mailBegriffe $mailNamen $null $null $mailKontext
    }
    return $false
}

function Search-Ordner($ordner, [string]$postfach, [datetime]$seit, $info) {
    $pfad = ''; try { $pfad = [string]$ordner.FolderPath } catch { }
    $kurz = ''; try { $kurz = [string]$ordner.Name } catch { }

    # Aelteste vorhandene Mail merken: zeigt, ob Outlook aeltere Mails nur auf dem Server hat
    $alle = $null
    try {
        $alle = $ordner.Items
        if ([int]$alle.Count -gt 0) {
            $alle.Sort('[ReceivedTime]', $false)
            $erste = $alle.GetFirst()
            if ($null -ne $erste) {
                $d = Get-MailDatum $erste
                if ($d -gt [datetime]::MinValue -and ($null -eq $info.Aelteste -or $d -lt $info.Aelteste)) { $info.Aelteste = $d }
                Release $erste
            }
        }
    } catch { }
    Release $alle

    $items = $null
    try { $items = $ordner.Items } catch { return }

    # Durchgang 1: alle Mails mit Anhang, neueste zuerst, bis zum Startdatum
    $auswahl = $null
    try { $auswahl = $items.Restrict('@SQL="urn:schemas:httpmail:hasattachment" = 1') } catch { $auswahl = $items }
    $sortiert = $false
    try { $auswahl.Sort('[ReceivedTime]', $true); $sortiert = $true } catch { }
    $gesamt = 0; try { $gesamt = [int]$auswahl.Count } catch { }
    $n = 0
    $it = $null
    try { $it = $auswahl.GetFirst() } catch { }
    while ($null -ne $it) {
        $n++
        $stop = $false
        try {
            $res = @(Search-Mail $it $postfach $pfad $seit $sortiert $true)
            $stop = ($res.Count -gt 0 -and $res[$res.Count - 1] -eq $true)
        } catch {
            $script:Stat.Fehler++
            Write-LogDatei ('Mail nicht lesbar in ' + $pfad + ': ' + $_.Exception.Message)
        }
        Release $it
        $it = $null
        if ($stop) { break }
        if ($n % 20 -eq 0) {
            $prozent = 0
            if ($gesamt -gt 0) { $prozent = [int][Math]::Min(100, 100 * $n / $gesamt) }
            Write-Progress -Id 2 -ParentId 1 -Activity ($kurz + ': Mails mit Anhang') -Status ($n.ToString() + ' / ' + $gesamt + '   gespeicherte Dokumente: ' + $script:Stat.Dokumente) -PercentComplete $prozent
        }
        if ($n % 200 -eq 0) { [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers() }
        $it = Get-Naechste $auswahl
    }
    if (-not [object]::ReferenceEquals($auswahl, $items)) { Release $auswahl }

    # Durchgang 2: Mails ohne Anhang mit einem UVV-Begriff in Betreff oder Text
    $auswahl = $null
    try { $auswahl = $items.Restrict($script:StichwortFilter) }
    catch { $auswahl = $null; Write-LogDatei ('Stichwortsuche in ' + $pfad + ' nicht moeglich: ' + $_.Exception.Message) }
    if ($null -ne $auswahl) {
        $it = $null
        try { $it = $auswahl.GetFirst() } catch { }
        while ($null -ne $it) {
            try { [void]@(Search-Mail $it $postfach $pfad $seit $false $false) }
            catch {
                $script:Stat.Fehler++
                Write-LogDatei ('Mail nicht lesbar in ' + $pfad + ': ' + $_.Exception.Message)
            }
            Release $it
            $it = Get-Naechste $auswahl
        }
        Release $auswahl
    }
    Release $items
    $info.Ordner++
}

function Search-Postfach($eintrag, [datetime]$seit) {
    $root = $eintrag.Root
    if ($null -eq $root) { $root = $eintrag.Store.GetRootFolder() }
    $info = [pscustomobject]@{ Name = $eintrag.Name; Aelteste = $null; Ordner = 0; Mails = 0 }
    $script:PostfachInfo.Add($info)
    $mailsVorher = $script:Stat.Mails
    Log ''
    Log ('Postfach: ' + $eintrag.Name) 'Cyan'
    $ordnerListe = Get-AlleOrdner $root
    $i = 0
    foreach ($o in $ordnerListe) {
        $i++
        $typ = -1; try { $typ = [int]$o.DefaultItemType } catch { }
        if ($typ -ne 0) { continue }
        $name = ''; try { $name = [string]$o.Name } catch { }
        Write-Progress -Id 1 -Activity ('Postfach ' + $eintrag.Name) -Status ('Ordner ' + $i + ' von ' + $ordnerListe.Count + ': ' + $name) -PercentComplete ([int](100 * ($i - 1) / [Math]::Max(1, $ordnerListe.Count)))
        $vorherMails = $script:Stat.Mails
        $vorherDok = $script:Stat.Dokumente
        Search-Ordner $o $eintrag.Name $seit $info
        $neu = $script:Stat.Mails - $vorherMails
        if ($neu -gt 0) { Log ('  ' + $name + ': ' + $neu + ' Mails, ' + ($script:Stat.Dokumente - $vorherDok) + ' neue Dokumente') }
    }
    $info.Mails = $script:Stat.Mails - $mailsVorher
    Write-Progress -Id 2 -ParentId 1 -Activity 'Mails' -Completed
    if ($null -ne $info.Aelteste) {
        Log ('  {Ae}lteste Mail, die Outlook hier anzeigt: ' + $info.Aelteste.ToString('dd.MM.yyyy'))
        if ($info.Aelteste -gt $seit.AddDays(14)) {
            Log ('  Achtung: Outlook hat hier nur Mails ab ' + $info.Aelteste.ToString('dd.MM.yyyy') + '. {Ae}ltere liegen vermutlich nur auf dem Server (Hinweis im Bericht).') 'Yellow'
        }
    }
}

# ------------------------------------------------------------------ Zeilen, Datumsangaben, Sammeldokumente

# Muster fuer den Nachnamen im Originaltext (Gross/klein egal, Umlaut auch als ae oder a)
function Get-RohMuster([string]$name) {
    $sb = New-Object System.Text.StringBuilder
    $s = $name
    try { $s = $s.Normalize([System.Text.NormalizationForm]::FormC) } catch { }
    foreach ($ch in $s.ToLowerInvariant().ToCharArray()) {
        $c = [int]$ch
        if ($c -eq 0xE4) { [void]$sb.Append('(?:' + [char]0xE4 + '|ae|a)') }
        elseif ($c -eq 0xF6) { [void]$sb.Append('(?:' + [char]0xF6 + '|oe|o)') }
        elseif ($c -eq 0xFC) { [void]$sb.Append('(?:' + [char]0xFC + '|ue|u)') }
        elseif ($c -eq 0xDF) { [void]$sb.Append('(?:' + [char]0xDF + '|ss)') }
        elseif ([char]::IsWhiteSpace($ch) -or $ch -eq [char]'-') { [void]$sb.Append('[\s\-]+') }
        else {
            $basis = ([string]$ch).Normalize([System.Text.NormalizationForm]::FormD).Substring(0, 1)
            if ($basis -cne [string]$ch) { [void]$sb.Append('(?:' + [regex]::Escape([string]$ch) + '|' + [regex]::Escape($basis) + ')') }
            else { [void]$sb.Append([regex]::Escape([string]$ch)) }
        }
    }
    $opt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    return [regex]::new(('(?<![\p{L}\p{N}])' + $sb.ToString() + '(?![\p{L}\p{N}])'), $opt)
}

# Datumsangaben (TT.MM.JJJJ, TT.MM.JJ, JJJJ-MM-TT) aus einem Text, einheitlich als TT.MM.JJJJ
function Get-Daten([string]$text) {
    $liste = New-Object System.Collections.Generic.List[string]
    if (-not $text) { return , $liste.ToArray() }
    $grenze = (Get-Date).AddYears(2)
    $muster = '(?<!\d)(\d{1,2})\s?\.\s?(\d{1,2})\s?\.\s?(\d{4}|\d{2})(?!\d)|(?<!\d)(\d{4})-(\d{2})-(\d{2})(?!\d)'
    foreach ($m in [regex]::Matches($text, $muster)) {
        try {
            if ($m.Groups[1].Success) {
                $j = [int]$m.Groups[3].Value
                if ($j -lt 100) { $j += 2000 }
                $d = New-Object DateTime ($j, [int]$m.Groups[2].Value, [int]$m.Groups[1].Value)
            } else {
                $d = New-Object DateTime ([int]$m.Groups[4].Value, [int]$m.Groups[5].Value, [int]$m.Groups[6].Value)
            }
            if ($d.Year -ge 2000 -and $d -le $grenze) {
                $t = $d.ToString('dd.MM.yyyy')
                if (-not $liste.Contains($t)) { $liste.Add($t) }
            }
        } catch { }
    }
    return , $liste.ToArray()
}

# Fuer jede gefundene Person bis zu 2 Zeilen aus dem Text, in denen ihr Name steht (Zeilen mit
# Vor- und Nachname bevorzugt), plus die Datumsangaben aus diesen Zeilen.
# So sieht man bei Sammeldokumenten, was in der Zeile der Person steht.
function Get-Kontext([string]$roh, $namen) {
    $erg = @{}
    if (-not $roh -or $null -eq $namen -or $namen.Count -eq 0) { return $erg }
    foreach ($p in $script:Personen) {
        if (-not $namen.ContainsKey($p.Anzeige)) { continue }
        $kandidaten = New-Object System.Collections.Generic.List[object]
        $nr = 0
        foreach ($m in $p.RxRoh.Matches($roh)) {
            $nr++
            if ($nr -gt 40) { break }
            $start = 0
            $nl = $roh.LastIndexOf([char]10, $m.Index)
            if ($nl -ge 0) { $start = $nl + 1 }
            $start = [Math]::Max($start, $m.Index - 120)
            $ende = $roh.IndexOf([char]10, $m.Index)
            if ($ende -lt 0) { $ende = $roh.Length }
            $ende = [Math]::Min($ende, $m.Index + $m.Length + 120)
            $zeile = $roh.Substring($start, $ende - $start)
            $zeile = [regex]::Replace($zeile, '[\p{C}]+', ' ')
            $zeile = [regex]::Replace($zeile, '\s+', ' ').Trim()
            if (-not $zeile) { continue }
            $punkte = 0
            if ($p.RxVor.IsMatch((ConvertTo-Normal $zeile))) { $punkte = 2 }
            $kandidaten.Add([pscustomobject]@{ Text = $zeile; Punkte = $punkte; Nr = $nr })
        }
        $zeilen = New-Object System.Collections.Generic.List[string]
        $daten = New-Object System.Collections.Generic.List[string]
        $sortiert = @($kandidaten | Sort-Object -Property @{ Expression = { $_.Punkte }; Descending = $true }, @{ Expression = { $_.Nr } })
        foreach ($k in $sortiert) {
            if ($zeilen.Count -ge 2) { break }
            if ($zeilen.Contains($k.Text)) { continue }
            $zeilen.Add($k.Text)
            foreach ($d in (Get-Daten $k.Text)) { if (-not $daten.Contains($d)) { $daten.Add($d) } }
        }
        if ($zeilen.Count -gt 0) {
            $erg[$p.Anzeige] = [pscustomobject]@{ Zeilen = $zeilen.ToArray(); Daten = $daten.ToArray() }
        }
    }
    return $erg
}

# Grobe Einschaetzung, was fuer eine Liste das Dokument ist (nur als Hinweis)
function Get-DokArt([string]$norm) {
    if (-not $norm) { return '' }
    $arten = New-Object System.Collections.Generic.List[string]
    foreach ($w in @('unterschrift', 'teilgenommen', 'teilnehmer', 'teilnahme', 'bestaetig', 'bestatig', 'zur kenntnis', 'unterwiesen')) {
        if ($norm.Contains($w)) { $arten.Add('Teilnahme-/Unterschriftenliste?'); break }
    }
    foreach ($w in @('abgelaufen', 'faellig', 'fallig', 'gueltig bis', 'gultig bis', 'ablaufdatum', 'laeuft ab', 'erinnerung')) {
        if ($norm.Contains($w)) { $arten.Add('Fristen-/Statusliste?'); break }
    }
    return ($arten -join ', ')
}

# ------------------------------------------------------------------ Office-XML zeilenweise lesen

function Read-ZipEintrag($eintrag) {
    $sr = New-Object System.IO.StreamReader($eintrag.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

# Word/PowerPoint/OpenDocument-XML -> Text. Absaetze = Zeilen, Tabellenzeilen bleiben eine Zeile
# (Zellen getrennt durch " | "), Textlaeufe werden ohne Leerzeichen zusammengesetzt.
function Convert-XmlZuText([string]$xml) {
    $opt = [System.Text.RegularExpressions.RegexOptions]::Singleline
    $absatz = '</w:p>|</a:p>|</text:p>|</text:h>|<w:br\b[^>]*/>|<w:cr\b[^>]*/>|<a:br\b[^>]*/>|<text:line-break/>'
    $tabZeile = [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $t = [regex]::Replace($m.Value, '</w:p>|</a:p>|</text:p>|</text:h>|<w:br\b[^>]*/>|<w:cr\b[^>]*/>|<a:br\b[^>]*/>|<text:line-break/>', ' ')
        $t = [regex]::Replace($t, '</w:tc>|</a:tc>|</table:table-cell>', ' | ')
        return $t + "`n"
    }
    $xml = [regex]::Replace($xml, '<w:tr\b.*?</w:tr>|<a:tr\b.*?</a:tr>|<table:table-row\b.*?</table:table-row>', $tabZeile, $opt)
    $xml = [regex]::Replace($xml, $absatz, "`n")
    $xml = [regex]::Replace($xml, '<w:tab/>|<text:tab/>|<text:s/>|<text:s [^>]*/>', ' ')
    $xml = [regex]::Replace($xml, '<[^>]+>', '')
    return [System.Net.WebUtility]::HtmlDecode($xml)
}

# Excel: welche Zellformate (Stil-Nummern) sind Datumsformate?
function Get-XlsxDatumsStile($zip) {
    $erg = @{}
    $e = $zip.GetEntry('xl/styles.xml')
    if ($null -eq $e) { return $erg }
    $xml = Read-ZipEintrag $e
    $formate = @{}
    foreach ($id in @(14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58)) { $formate[[int]$id] = $true }
    foreach ($m in [regex]::Matches($xml, '<numFmt\b[^>]*>')) {
        $idM = [regex]::Match($m.Value, 'numFmtId="(\d+)"')
        $codeM = [regex]::Match($m.Value, 'formatCode="([^"]*)"')
        if (-not $idM.Success -or -not $codeM.Success) { continue }
        $code = [System.Net.WebUtility]::HtmlDecode($codeM.Groups[1].Value)
        $code = [regex]::Replace($code, '"[^"]*"|\[[^\]]*\]|\\.', '')
        if ($code -match '[dyjt]' -and $code -match 'm') { $formate[[int]$idM.Groups[1].Value] = $true }
        else { $formate.Remove([int]$idM.Groups[1].Value) }
    }
    $xfs = [regex]::Match($xml, '<cellXfs\b[^>]*>(.*?)</cellXfs>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($xfs.Success) {
        $i = 0
        foreach ($m in [regex]::Matches($xfs.Groups[1].Value, '<xf\b[^>]*>')) {
            $id = 0
            $idM = [regex]::Match($m.Value, 'numFmtId="(\d+)"')
            if ($idM.Success) { $id = [int]$idM.Groups[1].Value }
            if ($formate.ContainsKey($id)) { $erg[[int]$i] = $true }
            $i++
        }
    }
    return $erg
}

# Excel (xlsx/xlsm): jede Tabellenzeile als eine Textzeile, Zellen getrennt durch " | ",
# Datumszellen als TT.MM.JJJJ (in der Datei stehen sie als Zahl)
function Get-XlsxText([string]$pfad) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($pfad)
    try {
        $opt = [System.Text.RegularExpressions.RegexOptions]::Singleline
        $texte = New-Object System.Collections.Generic.List[string]
        $e = $zip.GetEntry('xl/sharedStrings.xml')
        if ($null -ne $e) {
            $xml = Read-ZipEintrag $e
            foreach ($m in [regex]::Matches($xml, '<si>(.*?)</si>|<si/>', $opt)) {
                $inhalt = [regex]::Replace($m.Groups[1].Value, '<rPh\b.*?</rPh>', '', $opt)
                $texte.Add([System.Net.WebUtility]::HtmlDecode([regex]::Replace($inhalt, '<[^>]+>', '')))
            }
        }
        $stile = Get-XlsxDatumsStile $zip
        $sb = New-Object System.Text.StringBuilder
        foreach ($blatt in $zip.Entries) {
            if ($blatt.FullName -notlike 'xl/worksheets/*.xml' -or $blatt.Length -gt (80 * $script:MB)) { continue }
            $xml = Read-ZipEintrag $blatt
            $zeilenNr = 0
            foreach ($zeile in [regex]::Matches($xml, '<row\b[^>]*?(?<!/)>(.*?)</row>', $opt)) {
                $zeilenNr++
                if ($zeilenNr -gt 20000) { break }
                $zellen = New-Object System.Collections.Generic.List[string]
                foreach ($c in [regex]::Matches($zeile.Groups[1].Value, '<c\b([^>]*?)(?:/>|>(.*?)</c>)', $opt)) {
                    $attr = $c.Groups[1].Value
                    $inhalt = $c.Groups[2].Value
                    $typ = ''
                    $tm = [regex]::Match($attr, '\st="([^"]*)"')
                    if ($tm.Success) { $typ = $tm.Groups[1].Value }
                    $wert = ''
                    if ($typ -eq 'inlineStr') {
                        $wert = [System.Net.WebUtility]::HtmlDecode([regex]::Replace($inhalt, '<[^>]+>', ''))
                    } else {
                        $v = ''
                        $vm = [regex]::Match($inhalt, '<v>(.*?)</v>', $opt)
                        if ($vm.Success) { $v = $vm.Groups[1].Value }
                        if ($typ -eq 's') {
                            $idx = 0
                            if ([int]::TryParse($v, [ref]$idx) -and $idx -ge 0 -and $idx -lt $texte.Count) { $wert = $texte[$idx] }
                        } elseif ($typ -eq '' -or $typ -eq 'n') {
                            $wert = $v
                            $sm = [regex]::Match($attr, '\ss="(\d+)"')
                            $zahl = 0.0
                            if ($v -and $sm.Success -and $stile.ContainsKey([int]$sm.Groups[1].Value)) {
                                if ([double]::TryParse($v, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$zahl) -and $zahl -ge 1 -and $zahl -lt 2958466) {
                                    try { $wert = [datetime]::FromOADate($zahl).ToString('dd.MM.yyyy') } catch { }
                                }
                            }
                        } else {
                            $wert = [System.Net.WebUtility]::HtmlDecode($v)
                        }
                    }
                    $wert = $wert.Trim()
                    if ($wert) { $zellen.Add($wert) }
                }
                if ($zellen.Count -gt 0) { [void]$sb.AppendLine(($zellen -join ' | ')) }
            }
        }
        return $sb.ToString()
    } finally { $zip.Dispose() }
}

# ------------------------------------------------------------------ Dokumente lesen (2. Phase)

function Read-Datei($d) {
    try {
        $r = Get-DokumentText $d.Pfad $d.Endung 0
        $norm = ConvertTo-Normal $r.Text
        $d.Begriffe = Find-Begriffe $norm
        $d.Namen = Find-Personen $norm $false
        $d.Anzahl = $d.Namen.Count
        $d.Kontext = Get-Kontext $r.Text $d.Namen
        $d.DokArt = Get-DokArt $norm
        $d.Quelle = $r.Quelle
        $d.Hinweis = $r.Hinweis
        $d.Zeichen = $norm.Length
        if ($r.Quelle -eq 'nicht gelesen') { $d.Status = 'nicht gelesen' } else { $d.Status = 'gelesen' }
    } catch {
        $d.Status = 'Fehler'
        $d.Quelle = 'Fehler beim Lesen'
        $d.Hinweis = $_.Exception.Message
        Write-LogDatei ('Nicht lesbar: ' + $d.Pfad + ' - ' + $_.Exception.Message)
    }
}

function Read-Dokumente {
    $schnell = New-Object System.Collections.Generic.List[object]
    $langsam = New-Object System.Collections.Generic.List[object]
    foreach ($d in @($script:Dateien.Values)) {
        if ($d.Status -ne 'offen') { continue }
        if ($script:LangsamSet.ContainsKey($d.Endung)) { $langsam.Add($d) } else { $schnell.Add($d) }
    }
    $gesamt = $schnell.Count + $langsam.Count
    if ($gesamt -eq 0) { return }
    Log ''
    Log ('Dokumente lesen: ' + $schnell.Count + ' Office-/Textdateien, ' + $langsam.Count + ' PDFs/Bilder/Mails/ZIPs') 'Cyan'
    $n = 0
    foreach ($d in $schnell) {
        $n++
        if ($n % 10 -eq 0) {
            Write-Progress -Id 1 -Activity 'Dokumente lesen' -Status ($n.ToString() + ' / ' + $gesamt + ': ' + $d.Datei) -PercentComplete ([int](100 * $n / $gesamt))
        }
        Read-Datei $d
    }
    if ($langsam.Count -gt 0) {
        if ($script:OcrAktiv) {
            Log 'Texterkennung l{ae}uft (passende und neueste Dokumente zuerst). Abbrechen mit Strg+C ist jederzeit m{oe}glich.'
        }
        $reihe = @($langsam | Sort-Object -Property @{ Expression = { $_.Signal }; Descending = $true }, @{ Expression = { $_.Datum }; Descending = $true })
        $uhr = [System.Diagnostics.Stopwatch]::StartNew()
        $k = 0
        foreach ($d in $reihe) {
            $k++
            $rest = -1
            if ($k -gt 3) { $rest = [int](($uhr.Elapsed.TotalSeconds / ($k - 1)) * ($reihe.Count - $k + 1)) }
            Write-Progress -Id 1 -Activity (T 'Dokumente lesen (Texterkennung f{ue}r PDFs und Bilder)') -Status ($k.ToString() + ' / ' + $reihe.Count + ': ' + $d.Datei) -PercentComplete ([int](100 * $k / $reihe.Count)) -SecondsRemaining $rest
            Read-Datei $d
            if ($k % 50 -eq 0) { Log ('  ' + $k + ' von ' + $reihe.Count + ' gelesen') }
        }
    }
    Write-Progress -Id 1 -Activity 'Dokumente lesen' -Completed
}

# ------------------------------------------------------------------ Treffer auswerten

function Merge-Namen($ziel, $quelle, [string]$wo) {
    if ($null -eq $quelle) { return }
    foreach ($k in @($quelle.Keys)) {
        $s = [int]$quelle[$k]
        if (-not $ziel.ContainsKey($k)) { $ziel[$k] = [pscustomobject]@{ S = $s; Wo = $wo } }
        else {
            $e = $ziel[$k]
            if ($s -gt $e.S) { $e.S = $s }
            if ((', ' + $e.Wo + ',').IndexOf(', ' + $wo + ',') -lt 0) { $e.Wo = $e.Wo + ', ' + $wo }
        }
    }
}

# Relevanz: A = UVV-Begriff + Name, B = nur Name (Vor- und Nachname), K = UVV-Begriff ohne Namen
function Get-Treffer {
    $liste = New-Object System.Collections.Generic.List[object]
    foreach ($f in $script:Funde) {
        $datei = $null
        if ($f.Hash) { $datei = $script:Dateien[$f.Hash] }
        $begriffe = New-Object System.Collections.Generic.List[string]
        $quellen = @($f.MailBegriffe) + @($f.DateiBegriffe)
        if ($null -ne $datei) { $quellen += @($datei.Begriffe) }
        foreach ($b in $quellen) { if ($b -and -not $begriffe.Contains($b)) { $begriffe.Add($b) } }
        $namen = @{}
        Merge-Namen $namen $f.MailNamen 'Mail'
        Merge-Namen $namen $f.DateiNamen 'Dateiname'
        if ($null -ne $datei) { Merge-Namen $namen $datei.Namen 'Dokument' }
        $uvv = $begriffe.Count -gt 0
        foreach ($k in @($namen.Keys)) { if ($namen[$k].S -lt 2 -and -not $uvv) { $namen.Remove($k) } }
        if ($namen.Count -gt 0 -and $uvv) { $rel = 'A' }
        elseif ($namen.Count -gt 0) { $rel = 'B' }
        elseif ($uvv) { $rel = 'K' }
        else { continue }
        # Wie viele Mitarbeiter der Liste stehen im Dokument bzw. in der Mail?
        $anzahlMail = 0
        foreach ($k in @($f.MailNamen.Keys)) { if ([int]$f.MailNamen[$k] -ge 2) { $anzahlMail++ } }
        $anzahl = $anzahlMail
        $dokArt = ''
        if ($null -ne $datei) {
            if ($datei.Anzahl -gt $anzahl) { $anzahl = $datei.Anzahl }
            $dokArt = $datei.DokArt
        }
        $liste.Add([pscustomobject]@{
                Fund = $f; Datei = $datei; Relevanz = $rel; Begriffe = $begriffe.ToArray(); Namen = $namen
                Anzahl = $anzahl; Sammel = ($anzahl -ge 2); DokArt = $dokArt
            })
    }
    return , $liste
}

# Eine Zeile pro Mitarbeiter und Treffer, mit der Zeile aus dem Dokument
function Get-PersonenZeilen($treffer) {
    $liste = New-Object System.Collections.Generic.List[object]
    foreach ($t in $treffer) {
        foreach ($person in @($t.Namen.Keys)) {
            $kontext = $null
            if ($null -ne $t.Datei -and $t.Datei.Kontext.ContainsKey($person)) { $kontext = $t.Datei.Kontext[$person] }
            elseif ($t.Fund.MailKontext.ContainsKey($person)) { $kontext = $t.Fund.MailKontext[$person] }
            $zeilen = @(); $daten = @()
            if ($null -ne $kontext) { $zeilen = @($kontext.Zeilen); $daten = @($kontext.Daten) }
            $liste.Add([pscustomobject]@{
                    Person = $person; Treffer = $t; S = $t.Namen[$person].S; Wo = $t.Namen[$person].Wo
                    Zeilen = $zeilen; Daten = $daten
                })
        }
    }
    return , $liste
}

function Get-DatumText([datetime]$d) {
    if ($d -eq [datetime]::MinValue) { return '' }
    return $d.ToString('dd.MM.yyyy')
}

function Get-StaerkeText([int]$s) {
    if ($s -ge 3) { return 'Vor- und Nachname' }
    if ($s -eq 2) { return T 'Name (Vor- und Nachname getrennt, "Klatt, A." oder "Herr Klatt")' }
    return 'nur Nachname'
}

function Get-RelevanzText([string]$r) {
    if ($r -eq 'A') { return 'UVV-Bezug + Name' }
    if ($r -eq 'B') { return 'nur Name, kein UVV-Begriff' }
    return 'UVV-Bezug, kein Name erkannt'
}

function Get-ArtText($t) {
    $s = ''
    if ($t.Relevanz -ne 'K') {
        if ($t.Sammel) {
            if ($null -ne $t.Datei) { $s = 'Sammeldokument: ' + $t.Anzahl + ' Mitarbeiter der Liste' }
            else { $s = 'Sammel-Mail: ' + $t.Anzahl + ' Mitarbeiter der Liste' }
        } else { $s = 'nur dieser Mitarbeiter der Liste gefunden' }
    }
    if ($t.DokArt) { if ($s) { $s += ' - ' }; $s += $t.DokArt }
    return $s
}

function Get-FundartText([string]$art) {
    if ($art -eq 'Mail') { return '(Mail ohne Dokument-Anhang)' }
    if ($art -eq 'Mail verschluesselt') { return (T '(verschl{ue}sselte Mail - nur Betreff gepr{ue}ft)') }
    return ''
}

function Get-LeseText($datei) {
    if ($null -eq $datei) { return '' }
    $t = $datei.Quelle
    if ($datei.Status -eq 'offen') { $t = 'nicht gelesen (Suche abgebrochen)' }
    elseif ($datei.Status -eq 'gelesen' -and $datei.Zeichen -lt 20) { $t = $t + ' - kaum Text erkannt' }
    if ($datei.Hinweis) { $t = $t + ' (' + $datei.Hinweis + ')' }
    return $t
}

function Write-Csv($zeilen, [string]$pfad) {
    if (@($zeilen).Count -eq 0) {
        [System.IO.File]::WriteAllText($pfad, (T 'Keine Eintr{ae}ge'), (New-Object System.Text.UTF8Encoding($true)))
        return
    }
    $zeilen | Export-Csv -LiteralPath $pfad -NoTypeInformation -Delimiter ';' -Encoding UTF8
}

function Write-TrefferCsv($personenZeilen, $treffer, [string]$pfad) {
    $zeilen = New-Object System.Collections.Generic.List[object]
    $sortiert = @($personenZeilen | Sort-Object -Property @{ Expression = { $_.Person } }, @{ Expression = { $_.Treffer.Relevanz } }, @{ Expression = { $_.Treffer.Fund.Datum }; Descending = $true })
    $kTreffer = @($treffer | Where-Object { $_.Relevanz -eq 'K' } | Sort-Object -Property @{ Expression = { $_.Fund.Datum }; Descending = $true })
    foreach ($z in ($sortiert + $kTreffer)) {
        if ($null -ne $z.Treffer) { $t = $z.Treffer; $person = $z.Person } else { $t = $z; $person = '(kein Name erkannt)' }
        $f = $t.Fund
        $anhang = $f.Anhang
        if (-not $anhang) { $anhang = Get-FundartText $f.Art }
        $datei = ''
        if ($null -ne $t.Datei) { $datei = $t.Datei.Pfad }
        $zeileText = ''; $datumText = ''; $treff = ''
        if ($null -ne $z.Treffer) {
            $zeileText = ($z.Zeilen -join '  //  ')
            $datumText = ($z.Daten -join ', ')
            $treff = (Get-StaerkeText $z.S) + ' - gefunden in: ' + $z.Wo
        }
        $zeilen.Add([pscustomobject]@{
                'Mitarbeiter'                     = $person
                'Bewertung'                       = (Get-RelevanzText $t.Relevanz)
                'Einzel oder Sammel'              = (Get-ArtText $t)
                'Zeile im Dokument'               = $zeileText
                'Datum in der Zeile'              = $datumText
                'Namenstreffer'                   = $treff
                'UVV-Begriffe'                    = ($t.Begriffe -join ', ')
                'Mail-Datum'                      = (Get-DatumText $f.Datum)
                'Absender'                        = $f.Absender
                'Betreff'                         = $f.Betreff
                'Anhang'                          = $anhang
                'Datei'                           = $datei
                'Text gelesen'                    = (Get-LeseText $t.Datei)
                'Ordner'                          = $f.Ordner
            })
    }
    Write-Csv $zeilen $pfad
}

function Write-ZusammenfassungCsv($personenZeilen, [string]$pfad) {
    $zeilen = New-Object System.Collections.Generic.List[object]
    foreach ($p in $script:Personen) {
        $eigene = @($personenZeilen | Where-Object { $_.Person -eq $p.Anzeige })
        $a = @($eigene | Where-Object { $_.Treffer.Relevanz -eq 'A' } | Sort-Object -Property @{ Expression = { $_.Treffer.Fund.Datum }; Descending = $true })
        $einzel = @($a | Where-Object { -not $_.Treffer.Sammel })
        $sammel = @($a | Where-Object { $_.Treffer.Sammel })
        $b = @($eigene | Where-Object { $_.Treffer.Relevanz -eq 'B' })
        $e1 = $null; if ($einzel.Count -gt 0) { $e1 = $einzel[0] }
        $s1 = $null; if ($sammel.Count -gt 0) { $s1 = $sammel[0] }
        $zeilen.Add([pscustomobject]@{
                'Mitarbeiter'                                 = $p.Anzeige
                'Treffer mit UVV-Bezug'                       = $a.Count
                'davon nur dieser Mitarbeiter'                = $einzel.Count
                'davon Sammeldokumente'                       = $sammel.Count
                'Treffer nur mit Name'                        = $b.Count
                'Neuester Einzeltreffer: Mail-Datum'          = $(if ($e1) { Get-DatumText $e1.Treffer.Fund.Datum } else { '' })
                'Neuester Einzeltreffer: Datum in der Zeile'  = $(if ($e1) { $e1.Daten -join ', ' } else { '' })
                'Neuester Einzeltreffer: Datei'               = $(if ($e1 -and $e1.Treffer.Datei) { $e1.Treffer.Datei.Pfad } elseif ($e1) { $e1.Treffer.Fund.Betreff } else { '' })
                'Neuester Sammeltreffer: Mail-Datum'          = $(if ($s1) { Get-DatumText $s1.Treffer.Fund.Datum } else { '' })
                'Neuester Sammeltreffer: Zeile im Dokument'   = $(if ($s1) { $s1.Zeilen -join '  //  ' } else { '' })
                'Neuester Sammeltreffer: Datum in der Zeile'  = $(if ($s1) { $s1.Daten -join ', ' } else { '' })
                'Neuester Sammeltreffer: Datei'               = $(if ($s1 -and $s1.Treffer.Datei) { $s1.Treffer.Datei.Pfad } elseif ($s1) { $s1.Treffer.Fund.Betreff } else { '' })
            })
    }
    Write-Csv $zeilen $pfad
}

function Write-DokumenteCsv([string]$pfad) {
    $zeilen = New-Object System.Collections.Generic.List[object]
    foreach ($f in $script:Funde) {
        if ($f.Art -ne 'Anhang') { continue }
        $d = $script:Dateien[$f.Hash]
        if ($null -eq $d) { continue }
        $datei = $d.Pfad
        if (-not $datei) { $datei = (T '(gel{oe}scht - kein Treffer)') }
        $zeilen.Add([pscustomobject]@{
                'Mail-Datum'                        = (Get-DatumText $f.Datum)
                'Absender'                          = $f.Absender
                'Betreff'                           = $f.Betreff
                'Anhang'                            = $f.Anhang
                'Datei'                             = $datei
                'Text gelesen'                      = (Get-LeseText $d)
                'Mitarbeiter der Liste im Dokument' = ((@($d.Namen.Keys) | Sort-Object) -join '; ')
                'UVV-Begriffe im Dokument'          = (@($d.Begriffe) -join ', ')
                'Ordner'                            = $f.Ordner
            })
    }
    Write-Csv $zeilen $pfad
}

# Kopien: Treffer mit UVV-Bezug in einen Ordner pro Mitarbeiter ("Sammel_" = Sammeldokument),
# UVV-Dokumente ohne erkannten Namen in einen eigenen Ordner
function Copy-Treffer($personenZeilen, $treffer) {
    $basis = [System.IO.Path]::Combine($script:AusgabeDir, 'Treffer_nach_Mitarbeiter')
    $kopiert = @{}
    foreach ($z in $personenZeilen) {
        $t = $z.Treffer
        if ($t.Relevanz -ne 'A' -or $null -eq $t.Datei -or -not $t.Datei.Pfad) { continue }
        $schl = $z.Person + '|' + $t.Datei.Hash
        if ($kopiert.ContainsKey($schl)) { continue }
        $kopiert[$schl] = $true
        $ordner = [System.IO.Path]::Combine($basis, (Get-SichererName $z.Person 60))
        $name = $t.Datei.Datei
        if ($t.Sammel) { $name = 'Sammel_' + $name }
        try {
            [void][System.IO.Directory]::CreateDirectory($ordner)
            [System.IO.File]::Copy($t.Datei.Pfad, (Get-FreierPfad $ordner $name))
        } catch { Write-LogDatei ('Kopieren fehlgeschlagen: ' + $t.Datei.Pfad + ' - ' + $_.Exception.Message) }
    }
    $ohne = [System.IO.Path]::Combine($script:AusgabeDir, 'UVV_ohne_erkannten_Namen')
    foreach ($t in $treffer) {
        if ($t.Relevanz -ne 'K' -or $null -eq $t.Datei -or -not $t.Datei.Pfad) { continue }
        $schl = '|' + $t.Datei.Hash
        if ($kopiert.ContainsKey($schl)) { continue }
        $kopiert[$schl] = $true
        try {
            [void][System.IO.Directory]::CreateDirectory($ohne)
            [System.IO.File]::Copy($t.Datei.Pfad, (Get-FreierPfad $ohne $t.Datei.Datei))
        } catch { Write-LogDatei ('Kopieren fehlgeschlagen: ' + $t.Datei.Pfad + ' - ' + $_.Exception.Message) }
    }
}

function Remove-NichtTreffer($treffer) {
    $behalten = @{}
    foreach ($t in $treffer) { if ($null -ne $t.Datei) { $behalten[$t.Datei.Hash] = $true } }
    $n = 0
    foreach ($d in @($script:Dateien.Values)) {
        if ($behalten.ContainsKey($d.Hash) -or -not $d.Pfad) { continue }
        try { [System.IO.File]::Delete($d.Pfad); $d.Pfad = ''; $n++ } catch { }
    }
    Log ($n.ToString() + ' Dokumente ohne Treffer wurden wieder gel{oe}scht (Einstellung AlleAnhaengeBehalten = $false).')
}

# ------------------------------------------------------------------ Bericht (HTML)

function Get-Html([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }

# Relativer Link vom Bericht zur Datei (funktioniert auch, wenn der Ordner verschoben wird)
function Get-Link([string]$pfad) {
    if (-not $pfad) { return '' }
    $basis = $script:AusgabeDir.TrimEnd('\') + '\'
    if (-not $pfad.StartsWith($basis, [System.StringComparison]::OrdinalIgnoreCase)) { return '' }
    $teile = $pfad.Substring($basis.Length) -split '\\'
    return (($teile | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/')
}

function Get-DokZelle($t) {
    $f = $t.Fund
    if ($null -eq $t.Datei) { return '<span class="grau">' + (Get-Html (Get-FundartText $f.Art)) + '</span>' }
    $link = Get-Link $t.Datei.Pfad
    if ($link) { $s = '<a href="' + $link + '">' + (Get-Html $f.Anhang) + '</a>' } else { $s = Get-Html $f.Anhang }
    $lese = Get-LeseText $t.Datei
    if ($lese) { $s += '<div class="klein">' + (Get-Html $lese) + '</div>' }
    return $s
}

function Get-KontextZelle($zeilen, $daten) {
    $s = ''
    foreach ($zl in $zeilen) { $s += '<div class="zeile">' + (Get-Html $zl) + '</div>' }
    if (@($daten).Count -gt 0) { $s += '<div class="klein">Datum in der Zeile: <b>' + (Get-Html (@($daten) -join ', ')) + '</b></div>' }
    if (-not $s) { $s = '<span class="klein">-</span>' }
    return $s
}

function Add-PersonenTabelle($sb, $zeilen) {
    [void]$sb.Append('<table><tr><th>Mail-Datum</th><th>Dokument</th><th>Zeile im Dokument</th><th>Einzel / Sammel</th><th>Mail</th><th>Treffer</th></tr>')
    foreach ($z in $zeilen) {
        $t = $z.Treffer
        $f = $t.Fund
        $art = Get-Html (Get-ArtText $t)
        if ($t.Sammel) { $art = '<span class="sammel">' + $art + '</span>' }
        $mail = (Get-Html $f.Betreff) + '<div class="klein">' + (Get-Html $f.Absender) + ' &middot; ' + (Get-Html $f.Ordner) + '</div>'
        $treff = (Get-Html (Get-StaerkeText $z.S)) + '<div class="klein">gefunden in: ' + (Get-Html $z.Wo)
        if ($t.Begriffe.Count -gt 0) { $treff += '<br>Begriffe: ' + (Get-Html ($t.Begriffe -join ', ')) }
        $treff += '</div>'
        [void]$sb.Append('<tr><td>' + (Get-DatumText $f.Datum) + '</td><td>' + (Get-DokZelle $t) + '</td><td>' + (Get-KontextZelle $z.Zeilen $z.Daten) + '</td><td>' + $art + '</td><td>' + $mail + '</td><td>' + $treff + '</td></tr>')
    }
    [void]$sb.Append('</table>')
}

function Write-Bericht($treffer, $personenZeilen, [datetime]$seit, [datetime]$start, [string]$pfad) {
    $css = @'
body{font-family:"Segoe UI",Arial,sans-serif;margin:24px;color:#1f2328;background:#fff;font-size:14px;line-height:1.4}
h1{font-size:22px;margin:0 0 4px}
h2{font-size:18px;margin:32px 0 8px;padding-bottom:4px;border-bottom:1px solid #d0d7de}
h3{font-size:15px;margin:18px 0 6px}
.meta{color:#57606a;margin:0 0 12px}
.hinweis{background:#fff8c5;border:1px solid #d4a72c;padding:8px 12px;margin:8px 0;border-radius:6px}
.info{background:#ddf4ff;border:1px solid #54aeff;padding:8px 12px;margin:8px 0;border-radius:6px}
table{border-collapse:collapse;width:100%;margin:6px 0 14px}
th,td{border:1px solid #d0d7de;padding:5px 8px;text-align:left;vertical-align:top}
th{background:#f6f8fa;font-weight:600}
td.zahl{text-align:right}
.klein{font-size:12px;color:#57606a}
.grau{color:#57606a}
.zeile{font-family:Consolas,monospace;font-size:12px;background:#f6f8fa;padding:2px 4px;margin:1px 0}
.sammel{color:#9a6700;font-weight:600}
.nichts{color:#cf222e;font-weight:600}
details{margin:4px 0 14px}
summary{cursor:pointer;color:#0969da}
a{color:#0969da}
'@
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>UVV-Suche Ergebnis</title><style>' + $css + '</style></head><body>')
    [void]$sb.Append('<h1>UVV-Suche &ndash; Ergebnis</h1>')
    $postf = (@($script:PostfachInfo | ForEach-Object { $_.Name }) -join ', ')
    [void]$sb.Append('<p class="meta">' + (Get-Html ((T 'Erstellt am ') + (Get-Date).ToString('dd.MM.yyyy HH:mm') + (T ' | Postf{ae}cher: ') + $postf + ' | Mails ab ' + $seit.ToString('dd.MM.yyyy'))) + '</p>')
    $gelesen = @($script:Dateien.Values | Where-Object { $_.Status -eq 'gelesen' }).Count
    $minuten = [int]((Get-Date) - $start).TotalMinutes
    [void]$sb.Append('<p class="meta">' + (Get-Html (T ('Durchsuchte Mails: ' + $script:Stat.Mails + ' | Dokument-Anh{ae}nge: ' + $script:Stat.Anhaenge + ' (davon ' + $script:Stat.Dokumente + ' verschiedene gespeichert) | Text gelesen: ' + $gelesen + ' | Dauer: ' + $minuten + ' Min.'))) + '</p>')

    [void]$sb.Append('<div class="info">' + (T ('<b>So ist der Bericht zu lesen:</b> Ein Treffer hei{ss}t nur, dass der Name in einem Dokument oder einer Mail steht. ' +
                'Ob der Mitarbeiter wirklich unterwiesen wurde (Teilnahme, Unterschrift), kann das Skript nicht erkennen.<br>' +
                '<b>Sammeldokumente</b> (mehrere Mitarbeiter der Liste im selben Dokument, z. B. Teilnehmer- oder Fristenlisten) sind gelb markiert und unten ' +
                'noch einmal gesammelt. Dort bitte im Dokument pr{ue}fen, ob die Person wirklich teilgenommen bzw. unterschrieben hat. ' +
                'Die Spalte <i>Zeile im Dokument</i> zeigt, was neben dem Namen steht (z. B. ein Datum oder &bdquo;abgelaufen&ldquo;).<br>' +
                'Auch ein Dokument, in dem nur ein Mitarbeiter der Liste gefunden wurde, kann weitere Personen enthalten, die nicht in der Namensliste stehen.')) + '</div>')

    if ($script:Abgebrochen) {
        [void]$sb.Append('<div class="hinweis">' + (T 'Die Suche wurde vorzeitig beendet (abgebrochen oder Fehler). Der Bericht enth{ae}lt nur den Stand bis dahin.') + '</div>')
    }
    if (-not $script:OcrAktiv) {
        [void]$sb.Append('<div class="hinweis">' + (Get-Html $script:OcrGrund) + (T ' PDFs und Bilder wurden deshalb nur {ue}ber Dateiname und Mailtext gepr{ue}ft.') + '</div>')
    }
    foreach ($info in $script:PostfachInfo) {
        if ($null -ne $info.Aelteste -and $info.Aelteste -gt $seit.AddDays(14)) {
            [void]$sb.Append('<div class="hinweis">' + (T ('Im Postfach <b>' + (Get-Html $info.Name) + '</b> zeigt Outlook nur Mails ab ' + $info.Aelteste.ToString('dd.MM.yyyy') + '. ' +
                        '{Ae}ltere Mails liegen vermutlich nur auf dem Server und wurden nicht durchsucht. Abhilfe: Datei &rarr; Kontoeinstellungen &rarr; Kontoeinstellungen &rarr; ' +
                        'Konto doppelklicken &rarr; Regler &bdquo;E-Mail herunterladen f{ue}r&ldquo; auf &bdquo;Alle&ldquo;, Outlook neu starten, warten bis alles geladen ist, Suche wiederholen. ' +
                        'Beim Gruppenpostfach zus{ae}tzlich unter &bdquo;Weitere Einstellungen &rarr; Erweitert&ldquo; den Haken &bdquo;Freigegebene Ordner herunterladen&ldquo; entfernen. ' +
                        'Alternativ die {ae}lteren Mails in Outlook im Web suchen.')) + '</div>')
        }
    }
    if ($script:Stat.Verschluesselt -gt 0) {
        [void]$sb.Append('<div class="hinweis">' + (T ($script:Stat.Verschluesselt.ToString() + ' verschl{ue}sselte Mails wurden nicht ge{oe}ffnet (nur der Betreff wurde gepr{ue}ft).')) + '</div>')
    }
    if ($script:Stat.Fehler -gt 0) {
        [void]$sb.Append('<div class="hinweis">' + (T ($script:Stat.Fehler.ToString() + ' Mails oder Anh{ae}nge konnten nicht gelesen werden (Details in Protokoll.txt).')) + '</div>')
    }

    # Uebersicht
    [void]$sb.Append('<h2>' + (T '{Ue}bersicht') + '</h2><table><tr><th>Mitarbeiter</th><th>Treffer mit UVV-Bezug</th><th>davon nur dieser Mitarbeiter</th><th>davon Sammeldokumente</th><th>Neuester Treffer (Mail-Datum)</th><th>Treffer nur mit Name</th></tr>')
    $nr = 0
    foreach ($p in $script:Personen) {
        $nr++
        $eigene = @($personenZeilen | Where-Object { $_.Person -eq $p.Anzeige })
        $a = @($eigene | Where-Object { $_.Treffer.Relevanz -eq 'A' } | Sort-Object -Property @{ Expression = { $_.Treffer.Fund.Datum }; Descending = $true })
        $einzel = @($a | Where-Object { -not $_.Treffer.Sammel })
        $b = @($eigene | Where-Object { $_.Treffer.Relevanz -eq 'B' })
        $anzA = '<span class="nichts">keine</span>'
        if ($a.Count -gt 0) { $anzA = $a.Count.ToString() }
        $neu = ''
        if ($a.Count -gt 0) { $neu = Get-DatumText $a[0].Treffer.Fund.Datum }
        [void]$sb.Append('<tr><td><a href="#p' + $nr + '">' + (Get-Html $p.Anzeige) + '</a></td><td class="zahl">' + $anzA + '</td><td class="zahl">' + $einzel.Count + '</td><td class="zahl">' + ($a.Count - $einzel.Count) + '</td><td>' + $neu + '</td><td class="zahl">' + $b.Count + '</td></tr>')
    }
    [void]$sb.Append('</table>')
    [void]$sb.Append('<p><a href="#sammel">Sammeldokumente</a> &middot; <a href="#ohne">' + (T 'UVV-Unterlagen ohne erkannten Namen') + '</a> &middot; ' + (T 'F{ue}r Excel: Treffer.csv und Zusammenfassung.csv im selben Ordner.') + '</p>')

    # Je Mitarbeiter
    $nr = 0
    foreach ($p in $script:Personen) {
        $nr++
        $eigene = @($personenZeilen | Where-Object { $_.Person -eq $p.Anzeige })
        $a = @($eigene | Where-Object { $_.Treffer.Relevanz -eq 'A' } | Sort-Object -Property @{ Expression = { $_.Treffer.Fund.Datum }; Descending = $true })
        $einzel = @($a | Where-Object { -not $_.Treffer.Sammel })
        $sammel = @($a | Where-Object { $_.Treffer.Sammel })
        $b = @($eigene | Where-Object { $_.Treffer.Relevanz -eq 'B' } | Sort-Object -Property @{ Expression = { $_.Treffer.Fund.Datum }; Descending = $true })
        [void]$sb.Append('<h2 id="p' + $nr + '">' + (Get-Html $p.Anzeige) + '</h2>')
        if ($a.Count -eq 0) { [void]$sb.Append('<p class="nichts">Keine Unterlagen mit UVV-Bezug gefunden.</p>') }
        if ($einzel.Count -gt 0) {
            [void]$sb.Append('<h3>Nur dieser Mitarbeiter der Liste gefunden (' + $einzel.Count + ')</h3>')
            Add-PersonenTabelle $sb $einzel
        }
        if ($sammel.Count -gt 0) {
            [void]$sb.Append('<h3 class="sammel">' + (T ('Sammeldokumente (' + $sammel.Count + ') &ndash; bitte pr{ue}fen, ob der Mitarbeiter wirklich teilgenommen hat')) + '</h3>')
            Add-PersonenTabelle $sb $sammel
        }
        if ($b.Count -gt 0) {
            [void]$sb.Append('<details><summary>' + (T ($b.Count.ToString() + ' weitere Dokumente/Mails, in denen nur der Name vorkommt (ohne UVV-Begriff)')) + '</summary>')
            Add-PersonenTabelle $sb $b
            [void]$sb.Append('</details>')
        }
    }

    # Sammeldokumente: jedes Dokument einmal, mit allen Mitarbeitern der Liste
    $gesehen = @{}
    $sammelListe = New-Object System.Collections.Generic.List[object]
    foreach ($t in @($treffer | Where-Object { $_.Relevanz -eq 'A' -and $_.Sammel } | Sort-Object -Property @{ Expression = { $_.Fund.Datum }; Descending = $true })) {
        if ($null -ne $t.Datei) { $schl = $t.Datei.Hash } else { $schl = 'mail|' + $t.Fund.Betreff + '|' + $t.Fund.Datum.ToString('s') }
        if ($gesehen.ContainsKey($schl)) { continue }
        $gesehen[$schl] = $true
        $sammelListe.Add($t)
    }
    [void]$sb.Append('<h2 id="sammel">Sammeldokumente (' + $sammelListe.Count + ')</h2>')
    [void]$sb.Append('<p>' + (T 'Jedes Dokument nur einmal (neueste Mail), mit allen Mitarbeitern der Liste und der Zeile, in der ihr Name steht. So l{ae}sst sich jedes Dokument in einem Durchgang pr{ue}fen.') + '</p>')
    foreach ($t in $sammelListe) {
        $f = $t.Fund
        [void]$sb.Append('<h3>' + (Get-DatumText $f.Datum) + ' &ndash; ' + (Get-DokZelle $t) + '</h3>')
        [void]$sb.Append('<div class="klein">Mail: ' + (Get-Html $f.Betreff) + ' &middot; ' + (Get-Html $f.Absender) + ' &middot; ' + (Get-Html $f.Ordner))
        if ($t.DokArt) { [void]$sb.Append(' &middot; ' + (Get-Html $t.DokArt)) }
        [void]$sb.Append('</div><table><tr><th>Mitarbeiter</th><th>Zeile im Dokument</th><th>Treffer</th></tr>')
        foreach ($person in @($t.Namen.Keys | Sort-Object)) {
            $kontext = $null
            if ($null -ne $t.Datei -and $t.Datei.Kontext.ContainsKey($person)) { $kontext = $t.Datei.Kontext[$person] }
            elseif ($f.MailKontext.ContainsKey($person)) { $kontext = $f.MailKontext[$person] }
            $zellen = '<span class="klein">-</span>'
            if ($null -ne $kontext) { $zellen = Get-KontextZelle $kontext.Zeilen $kontext.Daten }
            [void]$sb.Append('<tr><td>' + (Get-Html $person) + '</td><td>' + $zellen + '</td><td class="klein">' + (Get-Html ((Get-StaerkeText $t.Namen[$person].S) + ' - ' + $t.Namen[$person].Wo)) + '</td></tr>')
        }
        [void]$sb.Append('</table>')
    }

    # UVV-Unterlagen ohne erkannten Namen
    $gesehen = @{}
    $ohne = New-Object System.Collections.Generic.List[object]
    foreach ($t in @($treffer | Where-Object { $_.Relevanz -eq 'K' } | Sort-Object -Property @{ Expression = { $_.Fund.Datum }; Descending = $true })) {
        if ($null -ne $t.Datei) { $schl = $t.Datei.Hash } else { $schl = 'mail|' + $t.Fund.Betreff + '|' + $t.Fund.Datum.ToString('s') }
        if ($gesehen.ContainsKey($schl)) { continue }
        $gesehen[$schl] = $true
        $ohne.Add($t)
    }
    [void]$sb.Append('<h2 id="ohne">' + (T ('UVV-Unterlagen ohne erkannten Namen (' + $ohne.Count + ')')) + '</h2>')
    [void]$sb.Append('<p>' + (T 'Enthalten einen UVV-Begriff, aber keinen Namen aus der Liste &ndash; z. B. handschriftlich ausgef{ue}llte Listen, schlecht lesbare Scans oder Unterlagen anderer Mitarbeiter. Bei Bedarf von Hand ansehen (Kopien im Ordner UVV_ohne_erkannten_Namen).') + '</p>')
    if ($ohne.Count -gt 0) {
        [void]$sb.Append('<table><tr><th>Mail-Datum</th><th>Dokument</th><th>Mail</th><th>Begriffe</th></tr>')
        foreach ($t in $ohne) {
            $f = $t.Fund
            [void]$sb.Append('<tr><td>' + (Get-DatumText $f.Datum) + '</td><td>' + (Get-DokZelle $t) + '</td><td>' + (Get-Html $f.Betreff) + '<div class="klein">' + (Get-Html $f.Absender) + ' &middot; ' + (Get-Html $f.Ordner) + '</div></td><td class="klein">' + (Get-Html ($t.Begriffe -join ', ')) + '</td></tr>')
        }
        [void]$sb.Append('</table>')
    }
    [void]$sb.Append('<p class="klein">' + (Get-Html ('Ergebnisordner: ' + $script:AusgabeDir)) + '</p></body></html>')
    [System.IO.File]::WriteAllText($pfad, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
}

function Write-Ergebnis([datetime]$seit, [datetime]$start) {
    Write-Progress -Id 2 -ParentId 1 -Activity 'Mails' -Completed
    Write-Progress -Id 1 -Activity 'UVV-Suche' -Completed
    if (-not $script:AusgabeDir -or $null -eq $script:Funde) { return }
    Log ''
    Log 'Bericht wird erstellt ...' 'Cyan'
    $treffer = Get-Treffer
    $personenZeilen = Get-PersonenZeilen $treffer
    try { Copy-Treffer $personenZeilen $treffer } catch { Log ('Kopieren der Treffer fehlgeschlagen: ' + $_.Exception.Message) 'Red' }
    if (-not $AlleAnhaengeBehalten) { try { Remove-NichtTreffer $treffer } catch { } }
    $html = [System.IO.Path]::Combine($script:AusgabeDir, 'Bericht.html')
    try { Write-Bericht $treffer $personenZeilen $seit $start $html } catch { Log ('Bericht.html konnte nicht geschrieben werden: ' + $_.Exception.Message) 'Red' }
    try { Write-TrefferCsv $personenZeilen $treffer ([System.IO.Path]::Combine($script:AusgabeDir, 'Treffer.csv')) } catch { Log ('Treffer.csv: ' + $_.Exception.Message) 'Red' }
    try { Write-ZusammenfassungCsv $personenZeilen ([System.IO.Path]::Combine($script:AusgabeDir, 'Zusammenfassung.csv')) } catch { Log ('Zusammenfassung.csv: ' + $_.Exception.Message) 'Red' }
    try { Write-DokumenteCsv ([System.IO.Path]::Combine($script:AusgabeDir, 'Alle_Dokumente.csv')) } catch { Log ('Alle_Dokumente.csv: ' + $_.Exception.Message) 'Red' }
    try { [System.IO.Directory]::Delete($script:TmpDir, $true) } catch { }

    $mitTreffer = @($personenZeilen | Where-Object { $_.Treffer.Relevanz -eq 'A' } | ForEach-Object { $_.Person } | Sort-Object -Unique).Count
    Log ''
    Log ('Durchsuchte Mails: ' + $script:Stat.Mails + ' | gespeicherte Dokumente: ' + $script:Stat.Dokumente + ' | Fehler: ' + $script:Stat.Fehler)
    Log ('Mitarbeiter mit Unterlagen mit UVV-Bezug: ' + $mitTreffer + ' von ' + $script:Personen.Count)
    Log ('Ergebnis: ' + $script:AusgabeDir) 'Green'
    if ([System.IO.File]::Exists($html)) { try { Invoke-Item -LiteralPath $html } catch { } }
}

# ------------------------------------------------------------------ Hauptablauf

function Main {
    try { $Host.UI.RawUI.WindowTitle = 'UVV-Suche' } catch { }
    Log ''
    Log '=== UVV-Suche in Outlook ===' 'Cyan'
    Log 'Liest nur - in Outlook wird nichts ver{ae}ndert. Alle Daten bleiben auf diesem Laptop.'
    Log ''
    if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
        throw (T 'PowerShell l{ae}uft auf diesem Laptop im eingeschr{ae}nkten Modus (von der IT so eingestellt). Das Skript kann so nicht arbeiten. Bitte die Suche von Hand machen (siehe Anleitung) oder die IT fragen.')
    }

    # Namensliste
    $skriptOrdner = $PSScriptRoot
    if (-not $skriptOrdner) { $skriptOrdner = (Get-Location).Path }
    $namenPfad = $NamenDatei
    if (-not $namenPfad) { $namenPfad = Join-Path $skriptOrdner 'namen.txt' }
    if (-not (Test-Path -LiteralPath $namenPfad)) {
        New-NamenVorlage $namenPfad
        Log ('Die Namensliste fehlt noch. Sie wird jetzt im Editor ge{oe}ffnet: ' + $namenPfad) 'Yellow'
        Log 'Bitte die Mitarbeiter eintragen (eine Person pro Zeile: Nachname, Vorname), speichern und den Editor schlie{ss}en.'
        try { Start-Process -FilePath 'notepad.exe' -ArgumentList ('"' + $namenPfad + '"') -Wait } catch { }
        [void](Read-Host (T 'Danach hier Enter dr{ue}cken'))
    }
    $script:Personen = Read-Namen $namenPfad
    Log ($script:Personen.Count.ToString() + ' Mitarbeiter in der Namensliste (' + $namenPfad + ').')
    if ($script:Personen.Count -eq 0) {
        $antwort = Read-Host (T 'Keine Namen gefunden. Trotzdem nur nach UVV-Begriffen suchen? (J/N)')
        if ($antwort -notmatch '^[jJyY]') { return }
    }
    Initialize-Begriffe $Suchbegriffe
    $script:StichwortFilter = Get-StichwortFilter

    # Outlook und Postfaecher
    Log 'Verbinde mit Outlook ...'
    Connect-Outlook
    $postfaecher = Get-PostfachListe
    if ($postfaecher.Count -eq 0) { throw (T 'In Outlook wurde kein Postfach gefunden.') }
    Log ''
    Log 'Gefundene Postf{ae}cher:'
    foreach ($p in $postfaecher) {
        $typText = Get-TypText $p.Typ
        $zusatz = ''
        if ($typText) { $zusatz = '  (' + $typText + ')' }
        Log ('  [' + $p.Nr + '] ' + $p.Name + $zusatz)
    }
    $standard = @($postfaecher | Where-Object { $_.Typ -ne 2 } | ForEach-Object { $_.Nr })
    $eingabe = Read-Host (T ('Welche Postf{ae}cher durchsuchen? Nummern mit Komma, z. B. 1,2 (Enter = ' + ($standard -join ',') + ')'))
    $auswahl = @()
    if (-not $eingabe -or -not $eingabe.Trim()) { $auswahl = $standard }
    else {
        foreach ($teil in ($eingabe -split '[,; ]+')) {
            $zahl = 0
            if ([int]::TryParse($teil, [ref]$zahl) -and @($postfaecher | Where-Object { $_.Nr -eq $zahl }).Count -gt 0) { $auswahl += $zahl }
        }
    }
    if ($auswahl.Count -eq 0) { throw (T 'Kein Postfach ausgew{ae}hlt.') }

    # Zeitraum und Modus
    $seit = $null
    while ($null -eq $seit) {
        $e = Read-Host (T ('Mails ab welchem Datum durchsuchen? (TT.MM.JJJJ, Enter = ' + $StandardSeit + ')'))
        if (-not $e -or -not $e.Trim()) { $e = $StandardSeit }
        $d = [datetime]::MinValue
        $formate = [string[]]@('dd.MM.yyyy', 'd.M.yyyy', 'dd.MM.yy', 'd.M.yy', 'yyyy-MM-dd')
        if ([datetime]::TryParseExact($e.Trim(), $formate, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { $seit = $d }
        else { Log 'Datum nicht erkannt, bitte z. B. so eingeben: 01.01.2024' 'Yellow' }
    }
    $modus = Read-Host (T 'Modus: 1 = gr{ue}ndlich (mit Texterkennung f{ue}r PDFs und Scans, dauert l{ae}nger), 2 = schnell (Enter = 1)')
    $gruendlich = ($modus -ne '2')

    # Ergebnisordner
    $haupt = $AusgabeHauptordner
    if (-not $haupt) { $haupt = Join-Path $env:USERPROFILE 'UVV-Suche' }
    $script:AusgabeDir = Join-Path $haupt ((Get-Date).ToString('yyyy-MM-dd_HHmm'))
    $script:AnhangDir = Join-Path $script:AusgabeDir 'Anhaenge'
    $script:TmpDir = Join-Path $script:AusgabeDir '_tmp'
    foreach ($o in @($script:AusgabeDir, $script:AnhangDir, $script:TmpDir)) { [void][System.IO.Directory]::CreateDirectory($o) }
    $script:LogDatei = Join-Path $script:AusgabeDir 'Protokoll.txt'
    Write-LogDatei ('Postfaecher: ' + ($auswahl -join ',') + ' | ab ' + $seit.ToString('dd.MM.yyyy') + ' | gruendlich: ' + $gruendlich + ' | Personen: ' + $script:Personen.Count)

    # Texterkennung pruefen
    $script:OcrAktiv = $false
    $script:OcrGrund = 'Schnellmodus: PDFs und Bilder wurden nicht gelesen.'
    if ($gruendlich) {
        Log 'Pr{ue}fe die Windows-Texterkennung ...'
        if ((Initialize-Ocr) -and (Test-Ocr)) { Log 'Texterkennung funktioniert.' 'Green' }
        else {
            $script:OcrAktiv = $false
            Log ('Texterkennung nicht m{oe}glich: ' + $script:OcrGrund) 'Yellow'
            Log 'Es geht ohne Texterkennung weiter: PDFs und Bilder werden nur {ue}ber Dateiname und Mailtext gepr{ue}ft.' 'Yellow'
        }
    }

    Initialize-Status
    Log ''
    Log ('Ergebnisordner: ' + $script:AusgabeDir)
    Log 'Die Suche l{ae}uft. Outlook kann in dieser Zeit langsamer reagieren. Abbrechen mit Strg+C - der Bericht wird trotzdem geschrieben.'
    $start = Get-Date
    $script:Abgebrochen = $true
    try {
        foreach ($p in $postfaecher) {
            if ($auswahl -notcontains $p.Nr) { continue }
            try { Search-Postfach $p $seit }
            catch {
                $script:Stat.Fehler++
                Log ('Fehler im Postfach ' + $p.Name + ': ' + $_.Exception.Message) 'Red'
            }
        }
        Read-Dokumente
        $script:Abgebrochen = $false
    } finally {
        Write-Ergebnis $seit $start
    }
}

# Start (beim Einbinden mit ". .\UVV-Suche.ps1" wird nichts ausgefuehrt)
if ($MyInvocation.InvocationName -ne '.') {
    try { Main }
    catch {
        Write-Host ''
        Write-Host ((T 'Abbruch: ') + $_.Exception.Message) -ForegroundColor Red
        Write-LogDatei ('Abbruch: ' + $_.Exception.Message)
    }
    Write-Host ''
    [void](Read-Host (T 'Fertig. Enter dr{ue}cken, um das Fenster zu schlie{ss}en'))
}

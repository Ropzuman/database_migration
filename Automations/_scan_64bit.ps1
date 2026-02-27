$exportRoot = 'C:\database_migration\AutoCAD\exported'

$results = @()

Get-ChildItem $exportRoot -Recurse -Include '*.bas', '*.cls' | ForEach-Object {
    $file = $_.FullName
    $rel = $file.Replace($exportRoot + '\', '')
    $lines = Get-Content $file -Encoding UTF8

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $num = $i + 1

        # 1. Declare ilman PtrSafe
        if ($line -match 'Declare\s+(Function|Sub)' -and $line -notmatch 'PtrSafe') {
            $results += [PSCustomObject]@{ Tyyppi = 'DECLARE_NO_PTRSAFE'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() }
        }

        # 2. Long-tyyppiset handle-muuttujat (ei LongPtr)
        if ($line -match '\bAs\s+Long\b' -and $line -notmatch '\bAs\s+LongPtr\b' -and $line -match '\b(hWnd|hDC|hWin|hwnd|pidList|pvoid|lParam|wParam)\b') {
            $results += [PSCustomObject]@{ Tyyppi = 'LONG_HANDLE'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() }
        }

        # 3. Jet-ajuri
        if ($line -match 'Jet\.OLEDB|Microsoft\.Jet') {
            $results += [PSCustomObject]@{ Tyyppi = 'JET_DRIVER'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() }
        }

        # 4. AutoCAD COM-yhteys koodissa
        if ($line -match 'CreateObject.*AutoCAD|GetObject.*AutoCAD') {
            $results += [PSCustomObject]@{ Tyyppi = 'ACAD_COM'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() }
        }
    }
}

# Ryhmittele
$grouped = $results | Group-Object Tiedosto | Sort-Object Name

Write-Host ""
Write-Host "==================================================================="
Write-Host "  AUTOCAD VBA - 64-BITTI ANALYYSI"
Write-Host "==================================================================="
Write-Host ""
Write-Host "Tiedostoja skannattu: $(Get-ChildItem $exportRoot -Recurse -Include '*.bas','*.cls').Count)"
Write-Host "Ongelmia yhteensa:    $($results.Count)"
Write-Host ""

foreach ($group in $grouped) {
    $items = $group.Group
    Write-Host "--- $($group.Name) ---" -ForegroundColor Yellow
    foreach ($item in $items) {
        $color = switch ($item.Tyyppi) {
            'DECLARE_NO_PTRSAFE' { 'Red' }
            'LONG_HANDLE' { 'Magenta' }
            'JET_DRIVER' { 'DarkRed' }
            'ACAD_COM' { 'Cyan' }
            default { 'Gray' }
        }
        Write-Host "  [$($item.Tyyppi)] rivi $($item.Rivi): $($item.Koodi)" -ForegroundColor $color
    }
    Write-Host ""
}

# Yhteenveto tyypeittain
Write-Host "==================================================================="
Write-Host "  YHTEENVETO ONGELMITTAIN"
Write-Host "==================================================================="
$results | Group-Object Tyyppi | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "  $($_.Name.PadRight(22)) : $($_.Count) kpl"
}
Write-Host ""

# Tallenna raportti
$reportPath = "$exportRoot\_64BIT_ANALYSIS.txt"
$lines_out = @()
$lines_out += "AUTOCAD VBA 64-BITTI ANALYYSI"
$lines_out += "Aika: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$lines_out += ""
foreach ($group in $grouped) {
    $lines_out += "--- $($group.Name) ---"
    foreach ($item in $group.Group) {
        $lines_out += "  [$($item.Tyyppi)] rivi $($item.Rivi): $($item.Koodi)"
    }
    $lines_out += ""
}
$lines_out += "=== YHTEENVETO ==="
$results | Group-Object Tyyppi | Sort-Object Count -Descending | ForEach-Object {
    $lines_out += "  $($_.Name): $($_.Count) kpl"
}
$lines_out | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "Raportti tallennettu: $reportPath"

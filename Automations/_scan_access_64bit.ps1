
# ACCESS VBA - Kattava 64-bitti tarkistus
# Skannaa kaikki .bas, .cls, .vba -tiedostot Access-kansiosta

$root = 'C:\database_migration\Access'
$reportPath = 'C:\database_migration\Logs\ACCESS_64BIT_SCAN.md'
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$fileCount = 0
$backups = [System.Collections.Generic.List[string]]::new()

Get-ChildItem $root -Recurse -Include '*.bas', '*.cls', '*.vba' | Sort-Object FullName | ForEach-Object {
    $file = $_.FullName
    $rel = $file.Replace($root + '\', '')
    $fileCount++

    # Varmuuskopio-tiedostot
    if ($file -match '\.backup$|\.bak$|_OLD\.|_VANHA\.|_Back\.') {
        $backups.Add($rel)
    }

    $lines = Get-Content $file -Encoding UTF8
    $in32bitBlock = $false  # Seuraa #Else-lohkoa (32-bitti-koodi, ei vaadi PtrSafe)

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $num = $i + 1

        # Seuraa #If VBA7 / #Else / #End If -lohkoja
        if ($line -match '^\s*#If\s+VBA7\b') { $in32bitBlock = $false }
        elseif ($line -match '^\s*#Else\b') { $in32bitBlock = $true }
        elseif ($line -match '^\s*#End\s+If\b') { $in32bitBlock = $false }

        # 1. Declare ilman PtrSafe - hypätään yli 32-bitti-fallback-lohkot
        if (-not $in32bitBlock -and $line -match '^\s*#?Declare\s+(Function|Sub)' -and $line -notmatch 'PtrSafe') {
            $results.Add([PSCustomObject]@{Tyyppi = 'DECLARE_NO_PTRSAFE'; Vakavuus = 'KRIITTINEN'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() })
        }
        # 2. Jet-ajuri
        if ($line -match 'Jet\.OLEDB|Microsoft\.Jet') {
            $results.Add([PSCustomObject]@{Tyyppi = 'JET_DRIVER'; Vakavuus = 'KRIITTINEN'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() })
        }
        # 3. Handle-muuttuja As Long (ei As LongPtr) - tarkistetaan etta handle itse on Long
        # Tarkka pattern: handle-nimi heti ennen "As Long" (ei "As LongPtr")
        if ($line -match '\b(hWnd|hDC|hWin|hwnd|hHandle)\s+As\s+Long\b' -and
            $line -notmatch '\b(hWnd|hDC|hWin|hwnd|hHandle)\s+As\s+LongPtr') {
            $results.Add([PSCustomObject]@{Tyyppi = 'LONG_HANDLE'; Vakavuus = 'KRIITTINEN'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() })
        }
        if ($line -match '\b(lParam|wParam|pvoid|pidList)\s+As\s+Long\b' -and $line -notmatch 'LongPtr') {
            $results.Add([PSCustomObject]@{Tyyppi = 'LONG_HANDLE'; Vakavuus = 'KRIITTINEN'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() })
        }
        # 4. DAO ilman etuliitettä
        if ($line -match '^\s*(Private\s+|Public\s+|Dim\s+)\w.*\bAs\s+(Database|Recordset|TableDef|QueryDef|Workspace)\b' -and
            $line -notmatch '\bAs\s+DAO\.' -and $line -notmatch '\bAs\s+ADODB\.') {
            $results.Add([PSCustomObject]@{Tyyppi = 'DAO_NO_PREFIX'; Vakavuus = 'VAROITUS'; Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() })
        }
    }
}

# Laske myos Excel-tiedostot
$excelRoot = 'C:\database_migration\Excel'
$excelFiles = Get-ChildItem $excelRoot -Recurse -Include '*.bas', '*.cls', '*.vba' -ErrorAction SilentlyContinue
$excelIssues = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($ef in $excelFiles) {
    $rel = $ef.FullName.Replace($excelRoot + '\', 'Excel\')
    $lines = Get-Content $ef.FullName -Encoding UTF8
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]; $num = $i + 1
        if ($line -match 'Jet\.OLEDB|Microsoft\.Jet') {
            $excelIssues.Add([PSCustomObject]@{Tiedosto = $rel; Rivi = $num; Koodi = $line.Trim() })
        }
    }
}

# --- Tulosta ja tallenna ---
$krittiset = $results | Where-Object Vakavuus -eq 'KRIITTINEN'
$varoitukset = $results | Where-Object Vakavuus -eq 'VAROITUS'

$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine("# ACCESS VBA - 64-BITTI SKANNAUSRAPORTTI")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("**Aika:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$null = $sb.AppendLine("**Skannattu:** $fileCount Access VBA -tiedostoa")
$null = $sb.AppendLine("**Excel-tiedostoja:** $($excelFiles.Count)")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("## Yhteenveto")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("| Tyyppi | Maara | Vakavuus |")
$null = $sb.AppendLine("|--------|-------|----------|")

$results | Group-Object Tyyppi | Sort-Object Count -Descending | ForEach-Object {
    $vak = $_.Group[0].Vakavuus
    $null = $sb.AppendLine("| $($_.Name) | $($_.Count) | $vak |")
}

if ($results.Count -eq 0) {
    $null = $sb.AppendLine("| *Ei ongelmia* | 0 | OK |")
}

$null = $sb.AppendLine("")

# Kriittiset ongelmat
if ($krittiset.Count -gt 0) {
    $null = $sb.AppendLine("## Kriittiset ongelmat (KAATUU 64-bitissa)")
    $null = $sb.AppendLine("")
    $krittiset | Group-Object Tiedosto | ForEach-Object {
        $null = $sb.AppendLine("### $($_.Name)")
        $_.Group | ForEach-Object {
            $null = $sb.AppendLine("- **[$($_.Tyyppi)]** r.$($_.Rivi): ``$($_.Koodi)``")
        }
        $null = $sb.AppendLine("")
    }
}

# DAO-varoitukset
if ($varoitukset.Count -gt 0) {
    $null = $sb.AppendLine("## DAO-etuliite puuttuu (Varoitukset)")
    $null = $sb.AppendLine("")
    $varoitukset | Group-Object Tiedosto | ForEach-Object {
        $null = $sb.AppendLine("### $($_.Name)")
        $_.Group | ForEach-Object {
            $null = $sb.AppendLine("- r.$($_.Rivi): ``$($_.Koodi)``")
        }
        $null = $sb.AppendLine("")
    }
}

# Excel Jet-ongelmat
if ($excelIssues.Count -gt 0) {
    $null = $sb.AppendLine("## Excel: Jet-ajuri havaittu")
    $excelIssues | ForEach-Object {
        $null = $sb.AppendLine("- **$($_.Tiedosto)** r.$($_.Rivi): ``$($_.Koodi)``")
    }
    $null = $sb.AppendLine("")
}
else {
    $null = $sb.AppendLine("## Excel: Jet-ajuri")
    $null = $sb.AppendLine("Ei Jet-ajuriviittauksia. OK.")
    $null = $sb.AppendLine("")
}

# Varmuuskopiotiedostot
if ($backups.Count -gt 0) {
    $null = $sb.AppendLine("## Siivottavat varmuuskopiotiedostot")
    $backups | ForEach-Object { $null = $sb.AppendLine("- $_") }
    $null = $sb.AppendLine("")
}

# Lopputulos
$null = $sb.AppendLine("## Lopputulos")
if ($krittiset.Count -eq 0 -and $excelIssues.Count -eq 0) {
    $null = $sb.AppendLine("**LAEPÄISSYT** - Ei kriittisia 64-bitti-ongelmia havaittu.")
}
else {
    $null = $sb.AppendLine("**EI LAEPÄSSYT** - Kriittisia ongelmia loydetty (ks. yllä).")
}

$sb.ToString() | Out-File -FilePath $reportPath -Encoding UTF8

# Konsoliyhteenveto
Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  ACCESS VBA 64-BITTI SKANNAUS - VALMIS" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  Tiedostoja skannattu : $fileCount" -ForegroundColor White
Write-Host "  Kriittiset ongelmat  : $($krittiset.Count)" -ForegroundColor $(if ($krittiset.Count -gt 0) { 'Red' } else { 'Green' })
Write-Host "  DAO-varoitukset      : $($varoitukset.Count)" -ForegroundColor $(if ($varoitukset.Count -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "  Excel Jet-ongelmat   : $($excelIssues.Count)" -ForegroundColor $(if ($excelIssues.Count -gt 0) { 'Red' } else { 'Green' })
Write-Host "  Varmuuskopio-tiedost : $($backups.Count)" -ForegroundColor $(if ($backups.Count -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""
Write-Host "  Raportti: $reportPath" -ForegroundColor Gray
Write-Host "==================================================="

if ($results.Count -gt 0) {
    Write-Host ""
    $results | Group-Object Tyyppi | Sort-Object Count -Descending | ForEach-Object {
        $color = switch ($_.Name) { 'DECLARE_NO_PTRSAFE' { 'Red' }; 'JET_DRIVER' { 'DarkRed' }; 'LONG_HANDLE' { 'Magenta' }; 'DAO_NO_PREFIX' { 'Yellow' }; default { 'Gray' } }
        Write-Host "  [$($_.Name)] $($_.Count) kpl:" -ForegroundColor $color
        $_.Group | Select-Object -First 3 | ForEach-Object {
            Write-Host "    $($_.Tiedosto) r.$($_.Rivi): $($_.Koodi.Substring(0,[Math]::Min(80,$_.Koodi.Length)))" -ForegroundColor Gray
        }
    }
}
if ($backups.Count -gt 0) {
    Write-Host ""
    Write-Host "  Varmuuskopiot:" -ForegroundColor Yellow
    $backups | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
}

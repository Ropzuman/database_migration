$exportRoot = 'C:\database_migration\AutoCAD\exported'
$found = 0

Get-ChildItem $exportRoot -Recurse -Include '*.frm' | ForEach-Object {
    $file = $_.FullName
    $rel = $file.Replace($exportRoot + '\', '')
    $lines = Get-Content $file -Encoding UTF8

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $num = $i + 1

        if ($line -match 'Declare\s+(Function|Sub)' -and $line -notmatch 'PtrSafe') {
            if ($found -eq 0) { Write-Host "`n--- DECLARE ilman PtrSafe (.frm) ---" -ForegroundColor Yellow }
            Write-Host "  $rel rivi $num : $($line.Trim())" -ForegroundColor Red
            $found++
        }
    }
}

if ($found -eq 0) {
    Write-Host ".frm-tiedostoissa ei 64-bitti-ongelmia." -ForegroundColor Green
}
else {
    Write-Host "`nYhteensa: $found ongelmarivia .frm-tiedostoissa." -ForegroundColor Yellow
}

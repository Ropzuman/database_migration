<#
.SYNOPSIS
    Korjaa MS Access VBA -tiedostojen (Mojibake) merkistövirheet rekursiivisesti.
.DESCRIPTION
    Skripti etsii .bas, .cls, .frm, .vba ja .txt -tiedostot, korjaa rikkoutuneet ääkköset ja
    muut erikoismerkit sekä tallentaa tiedostot UTF-8 BOM -muodossa.
#>

param (
    [string]$Path = ".\",
    [switch]$DryRun
)

# Määritetään korjaussäännöt (Mojibake -> Oikea merkki)
$replacements = [ordered]@{
    'ï¿½'        = 'ä'
    'G'          = '–'
    'P\+ivitetty' = 'Päivitetty'
    'pï¿½ï¿½'    = 'pää'
    'Pï¿½ï¿½'    = 'Pää'
    'ï¿½ï¿½'     = 'ää'
}

# Haetaan kaikki kansion VBA-tiedostot
\(files = Get-ChildItem -Path\)Path -Recurse -Include *.bas, *.cls, *.frm, *.txt, *.vba

Write-Host "Löytyi \((\)files.Count) tarkistettavaa tiedostoa." -ForegroundColor Cyan

foreach (\(file in\)files) {
    # Luetaan tiedosto raw-tekstinä UTF-8-muodossa
    \(content = Get-Content -Path\)file.FullName -Raw -Encoding UTF8
    \(modified =\)false
    \(updatedContent =\)content

    # Ajetaan korjaussäännöt
    foreach (\(key in\)replacements.Keys) {
        if (\(updatedContent -match [regex]::Escape(\)key)) {\(updatedContent =\)updatedContent -replace [regex]::Escape(\(key),\)replacements[$key]
            \(modified =\)true
        }
    }

    if ($modified) {
        Write-Host "[KORJATTU] \((\)file.FullName)" -ForegroundColor Green
        if (-not $DryRun) {
            # Tallennetaan UTF-8 BOM -muodossa, jotta VS Code ja Access tunnistavat merkistön oikein
            [System.IO.File]::WriteAllText(\(file.FullName,\)updatedContent, [System.Text.Encoding]::UTF8)
        }
    } else {
        Write-Host "[OK] \((\)file.FullName)" -ForegroundColor Gray
    }
}

if ($DryRun) {
    Write-Host "`nDryRun päättyi. Mitään tiedostoja ei muutettu levyltä." -ForegroundColor Yellow
} else {
    Write-Host "`nValmis! Kaikki merkistökorjaukset tallennettu." -ForegroundColor Green
}
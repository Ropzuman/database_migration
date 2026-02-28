$ae = [char]0xe4; $oe = [char]0xf6

$replacements = @(
    @("' Updated 2025-11-11: Added DAO prefix for early binding", "' P${ae}ivitetty 2025-11-11: Lis${ae}tty DAO-etuliite aikaista sidontaa varten")
    @("' Updated 2025-11-11: Added DAO prefix for early binding", "' P${ae}ivitetty 2025-11-11: Lis${ae}tty DAO-etuliite aikaista sidontaa varten")
    @("' Updated 2025-11-11: Added DAO prefix", "' P${ae}ivitetty 2025-11-11: Lis${ae}tty DAO-etuliite")
    @("' Updated 2025-11-11: Changed from CurrentDB to CurrentDb (standard capitalization)", "' P${ae}ivitetty 2025-11-11: Muutettu CurrentDB -> CurrentDb")
    @("' Updated 2025-11-11: Changed from CurrentDb", "' P${ae}ivitetty 2025-11-11: Muutettu CurrentDb:sta")
    @("' Updated 2025-11-11: Changed from early binding (AcadApplication) to late binding (Object)", "' P${ae}ivitetty 2025-11-11: Muutettu aikainen sidonta (AcadApplication) my${oe}h${ae}iseksi sidonnaksi (Object)")
    @("' Updated 2025-11-11: Changed from early binding to late binding", "' P${ae}ivitetty 2025-11-11: Muutettu aikainen sidonta my${oe}h${ae}iseksi sidonnaksi")
    @("' Updated 2025-11-11: Added VBA7 conditional compilation for 64-bit compatibility", "' P${ae}ivitetty 2025-11-11: Lis${ae}tty VBA7-ehdollinen k${ae}${ae}nt${ae}minen 64-bit-yhteensopivuutta varten")
    @("' Updated 2025-11-11: Added DAO typing, improved comments", "' P${ae}ivitetty 2025-11-11: Lis${ae}tty DAO-tyypitys, parannettu kommentit")
    @("' Updated 2025-11-11: CurrentDB -> CurrentDb", "' P${ae}ivitetty 2025-11-11: CurrentDB -> CurrentDb")
    @("' Updated 2025-11-11:", "' P${ae}ivitetty 2025-11-11:")
    @("' Updated 2025-10-22: 64-bit compatibility, cleaner code", "' P${ae}ivitetty 2025-10-22: 64-bit-yhteensopivuus, siistimpi koodi")
    @("' Updated 2025-10-23: Changed API Declarations from Private to Public", "' P${ae}ivitetty 2025-10-23: Muutettu API-m${ae}${ae}rittelyt Private:sta Public:ksi")
    @("' Updated 2025-10-22:", "' P${ae}ivitetty 2025-10-22:")
    @("' Updated 2025-10-26:", "' P${ae}ivitetty 2025-10-26:")
    @("' Updated 2025-10-30:", "' P${ae}ivitetty 2025-10-30:")
    @("' Updated 2025-10-", "' P${ae}ivitetty 2025-10-")
    @("' Check each field (except ID and Rev) for changes", "' Tarkistetaan jokainen kentt${ae} (paitsi ID ja Rev) muutoksia varten")
    @("' Find document in DOCUMENTS table", "' Etsit${ae}${ae}n asiakirja DOCUMENTS-taulusta")
    @("' Ensure path ends with backslash", "' Varmistetaan ett${ae} polku p${ae}${ae}ttyy kenoviivaan")
    @("' Open cross-reference LISP lookup table", "' Avataan ristiviittaus LISP-hakutaulu")
    @("' Close non-loop-based files", "' Suljetaan ei-silmukkapohjaiset tiedostot")
    @("' Close loop-based files", "' Suljetaan silmukkapohjaiset tiedostot")
    @("' Generate AutoCAD script file for batch processing", "' Generoidaan AutoCAD-skriptitiedosto eras${ae}k${ae}sittelyi${ae} varten")
    @("' Cleanup on error", "' Siivotaan virhetilanteessa")
    @("' Cleanup", "' Siivotaan")
    @("Close #1  ' Close any open file handle", "Close #1  ' Suljetaan avoin tiedostok${ae}ynti")
    @("' Build reference prefix from ID fields", "' Rakennetaan viiteprefiksi ID-kent${ae}ist${ae}")
)

$count = 0
$allFiles = @(Get-ChildItem "c:\database_migration\Access" -Recurse -Include "*.bas", "*.cls", "*.vba") + @(Get-ChildItem "c:\database_migration\Excel\Moduulit" -Recurse -Include "*.bas")
$allFiles | ForEach-Object {
    $f = $_.FullName
    $orig = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)
    $c = $orig
    foreach ($r in $replacements) {
        $c = $c.Replace($r[0], $r[1])
    }
    if ($c -ne $orig) {
        [System.IO.File]::WriteAllText($f, $c, [System.Text.Encoding]::UTF8)
        $count++
        Write-Host "OK: $($_.Name)"
    }
}
Write-Host "Muutettu: $count tiedostoa"

# Fresh English comment scan — writes results to Logs\current_english.txt
$files = Get-ChildItem -Path "c:\database_migration\Access", "c:\database_migration\Excel\Moduulit" `
    -Include "*.bas", "*.cls", "*.vba" -Recurse -File

$englishPattern = [regex]"(?i)^\s*'[^']*\b(initialize|iterate|handle|handles|extract|extracts|reads?|writes?|opens?|closes?|checks?|creates?|deletes?|returns?|parses?|filters?|searches?|prevents?|builds?|generates?|converts?|replaces?|processes?|provides?|connects?|disconnects?|loads?|loops?|execute|executes?|normalize|fallback|fetch|fetches?|append|inserts?|removes?|sets?|gets?|adds?|drops?|populate|populates?|marks?|log|logging|delimit|token|format|validate|validates?)\b"

$results = [System.Collections.Generic.List[string]]::new()
foreach ($file in $files) {
    $lines = Get-Content $file.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($null -eq $lines) { continue }
    $rel = $file.FullName.Replace("c:\database_migration\", "")
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($englishPattern.IsMatch($line)) {
            $results.Add("$rel`:$($i+1): $($line.Trim())")
        }
    }
}

$results | Set-Content "c:\database_migration\Logs\current_english.txt" -Encoding UTF8
Write-Host "Found $($results.Count) potentially English comment lines"
Write-Host "Saved to Logs\current_english.txt"

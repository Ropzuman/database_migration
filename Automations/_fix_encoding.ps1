$files = @(
    'C:\database_migration\Automations\AutoCAD_DVB_Import_Run.ps1'
)
foreach ($path in $files) {
    $content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $utf8bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($path, $content, $utf8bom)
    Write-Host "BOM lisatty: $(Split-Path $path -Leaf)"
}
Write-Host "Valmis."

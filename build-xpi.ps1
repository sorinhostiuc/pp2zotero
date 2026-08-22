Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

Set-Location $PSScriptRoot

$xpiPath = Join-Path (Get-Location) 'pp2zotero-3.0.3.xpi'
if (Test-Path $xpiPath) { Remove-Item $xpiPath }

$zip = [System.IO.Compression.ZipFile]::Open($xpiPath, 'Create')

# Add root files
foreach ($f in @('manifest.json', 'bootstrap.js', 'prefs.js')) {
    $fullPath = Join-Path (Get-Location) $f
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $fullPath, $f) | Out-Null
}

# Add content/ recursively (exclude tests, node_modules, .DS_Store)
$basePath = (Get-Location).Path
Get-ChildItem -Path 'content' -Recurse -File | Where-Object {
    $_.Name -ne '.DS_Store' -and
    $_.FullName -notmatch 'node_modules' -and
    $_.FullName -notmatch '\\tests\\'
} | ForEach-Object {
    $rel = $_.FullName.Substring($basePath.Length + 1).Replace('\', '/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
}

# Add locale/ recursively
if (Test-Path 'locale') {
    Get-ChildItem -Path 'locale' -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($basePath.Length + 1).Replace('\', '/')
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
    }
}

$zip.Dispose()

$size = (Get-Item $xpiPath).Length / 1KB
Write-Host "Built: pp2zotero-3.0.3.xpi ($([math]::Round($size, 1)) KB)"

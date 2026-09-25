# build.ps1
# Packages the extension into two loadable zips under .\dist\ :
#   - onedrive-pdf-download-unlocker-chrome.zip   (uses manifest.json as-is)
#   - onedrive-pdf-download-unlocker-firefox.zip  (manifest.firefox.json renamed to manifest.json)
#
# Why two zips: a folder can only have one active manifest.json, but Chrome
# needs background.service_worker while Firefox needs background.scripts.
# This script does the rename automatically so testers just unzip and load.
#
# Usage:  powershell -ExecutionPolicy Bypass -File build.ps1

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$dist = Join-Path $root "dist"
$staging = Join-Path $root ".build-staging"

# Runtime files shared by both browsers (everything the extension needs at run time).
$files = @("background.js", "content.js", "i18n.js", "popup.html", "popup.js", "slides.js", "slidepdf.js")
$dirs = @("icons")

function Reset-Dir($p) {
    if (Test-Path $p) { Remove-Item $p -Recurse -Force }
    New-Item -ItemType Directory -Path $p | Out-Null
}

function Copy-Runtime($destDir) {
    foreach ($f in $files) { Copy-Item (Join-Path $root $f) (Join-Path $destDir $f) }
    foreach ($d in $dirs) { Copy-Item (Join-Path $root $d) (Join-Path $destDir $d) -Recurse }
}

# Zip a staging folder with FORWARD-SLASH entry names. PowerShell 5.1's
# Compress-Archive (and .NET Framework's ZipFile.CreateFromDirectory) write
# backslash separators on Windows, which breaks extraction on macOS/Linux
# (you get a literal "icons\icon16.png" file and the manifest's icon paths
# fail). Building entries by hand keeps the packages cross-platform.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
function New-Zip($srcDir, $zipPath) {
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $srcFull = (Resolve-Path $srcDir).Path.TrimEnd('\') + '\'
        Get-ChildItem $srcDir -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($srcFull.Length).Replace('\', '/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
        }
    }
    finally { $zip.Dispose() }
}

Reset-Dir $dist

# --- Chrome ---
$chromeStage = Join-Path $staging "chrome"
Reset-Dir $chromeStage
Copy-Item (Join-Path $root "manifest.json") (Join-Path $chromeStage "manifest.json")
Copy-Runtime $chromeStage
New-Zip $chromeStage (Join-Path $dist "onedrive-pdf-download-unlocker-chrome.zip")

# --- Firefox (rename manifest.firefox.json -> manifest.json inside the package) ---
$ffStage = Join-Path $staging "firefox"
Reset-Dir $ffStage
Copy-Item (Join-Path $root "manifest.firefox.json") (Join-Path $ffStage "manifest.json")
Copy-Runtime $ffStage
New-Zip $ffStage (Join-Path $dist "onedrive-pdf-download-unlocker-firefox.zip")

Remove-Item $staging -Recurse -Force

Write-Host "Built into $dist :"
Get-ChildItem $dist | ForEach-Object {
    Write-Host ("  {0}  ({1:N1} KB)" -f $_.Name, ($_.Length / 1KB))
}

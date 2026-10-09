# build.ps1
# Packages the extension into two loadable zips under .\dist\ :
#   - onedrive-pdf-download-unlocker-<version>-chrome.zip   (manifest.json as-is)
#   - onedrive-pdf-download-unlocker-<version>-firefox.zip  (Firefox manifest generated from manifest.json)
# <version> is read from manifest.json, so the file names always match the release.
#
# Why two zips: a folder can only have one active manifest.json, but Chrome
# needs background.service_worker while Firefox needs background.scripts.
#
# The Firefox manifest is GENERATED from manifest.json (there is no separate
# file to keep in sync — it used to drift, e.g. stuck on an old version):
#   - background.service_worker -> background.scripts
#   - drop the `debugger` permission and the PowerPoint slide-export content
#     script (Firefox has no debugger API, so no slide export there)
#   - drop version_name (Chrome-only key)
#   - add browser_specific_settings.gecko (add-on id, minimum Firefox version)
#
# Works in Windows PowerShell 5.1 and PowerShell 7+.
# Usage:  powershell -ExecutionPolicy Bypass -File build.ps1

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$dist = Join-Path $root "dist"
$staging = Join-Path $root ".build-staging"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# Runtime files shared by both browsers (everything the extension needs at run time).
$files = @("background.js", "content.js", "i18n.js", "util.js", "popup.html", "popup.js", "slides.js", "slides-viewer.js", "slidepdf.js")
$dirs = @("icons")

function Reset-Dir($p) {
    if (Test-Path $p) { Remove-Item $p -Recurse -Force }
    New-Item -ItemType Directory -Path $p | Out-Null
}

function Copy-Runtime($destDir) {
    foreach ($f in $files) { Copy-Item (Join-Path $root $f) (Join-Path $destDir $f) }
    foreach ($d in $dirs) { Copy-Item (Join-Path $root $d) (Join-Path $destDir $d) -Recurse }
}

function Write-Json($obj, $path) {
    $json = ConvertTo-Json -InputObject $obj -Depth 20
    # No BOM: a BOM in manifest.json is not valid JSON for every browser.
    [System.IO.File]::WriteAllText($path, $json, $utf8NoBom)
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
        $srcFull = (Resolve-Path $srcDir).Path.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
        Get-ChildItem $srcDir -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($srcFull.Length).Replace('\', '/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
        }
    }
    finally { $zip.Dispose() }
}

$manifestText = [System.IO.File]::ReadAllText((Join-Path $root "manifest.json"), $utf8NoBom)
$version = (ConvertFrom-Json $manifestText).version
$base = "onedrive-pdf-download-unlocker-$version"

Reset-Dir $dist

# --- Chrome ---
$chromeStage = Join-Path $staging "chrome"
Reset-Dir $chromeStage
Copy-Item (Join-Path $root "manifest.json") (Join-Path $chromeStage "manifest.json")
Copy-Runtime $chromeStage
New-Zip $chromeStage (Join-Path $dist "$base-chrome.zip")

# --- Firefox (manifest generated from manifest.json) ---
$ff = ConvertFrom-Json $manifestText
$ff.PSObject.Properties.Remove("version_name")
$ff.permissions = @($ff.permissions | Where-Object { $_ -ne "debugger" })
$ff.background = [ordered]@{ scripts = @("background.js") }
$ff.content_scripts = @($ff.content_scripts | Where-Object { $_.js -notcontains "slides.js" })
$ff | Add-Member -NotePropertyName "browser_specific_settings" -NotePropertyValue ([ordered]@{
    gecko = [ordered]@{
        id = "onedrive-pdf-download-unlocker@chidi0716"
        strict_min_version = "115.0"
    }
})

$ffStage = Join-Path $staging "firefox"
Reset-Dir $ffStage
Write-Json $ff (Join-Path $ffStage "manifest.json")
Copy-Runtime $ffStage
New-Zip $ffStage (Join-Path $dist "$base-firefox.zip")

Remove-Item $staging -Recurse -Force

Write-Host "Built version $version into $dist :"
Get-ChildItem $dist | ForEach-Object {
    Write-Host ("  {0}  ({1:N1} KB)" -f $_.Name, ($_.Length / 1KB))
}

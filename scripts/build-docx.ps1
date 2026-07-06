<#
Builds a .docx from an article Markdown file by hand-assembling the OOXML package.
Used because this machine has no node/python/soffice in PATH (see repo README / memory notes).

Usage:
  powershell -File scripts/build-docx.ps1 -MdPath articulos/2026-07-05-foo.md -Title "Article title" -Date 2026-07-05

Requires: perl (available via Git Bash at /usr/bin/perl), scripts/md2docx.pl, scripts/docx-template/.
#>
param(
    [Parameter(Mandatory=$true)][string]$MdPath,
    [Parameter(Mandatory=$true)][string]$Title,
    [Parameter(Mandatory=$true)][string]$Date
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$templateDir = Join-Path $scriptDir "docx-template"
$mdFull = (Get-Item $MdPath).FullName
$outDocx = [System.IO.Path]::ChangeExtension($mdFull, ".docx")

$buildDir = Join-Path $env:TEMP ("docxbuild_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
Copy-Item -Path (Join-Path $templateDir "*") -Destination $buildDir -Recurse -Force

# Render core.xml from template with title/date substitution
$coreTmpl = Get-Content (Join-Path $templateDir "docProps\core.xml.tmpl") -Raw -Encoding UTF8
$coreXml = $coreTmpl.Replace("{{TITLE}}", $Title).Replace("{{DATE}}", $Date)
$coreOut = Join-Path $buildDir "docProps\core.xml"
[System.IO.File]::WriteAllText($coreOut, $coreXml, [System.Text.UTF8Encoding]::new($false))
Remove-Item (Join-Path $buildDir "docProps\core.xml.tmpl") -Force

# Generate word/document.xml from the markdown via the perl converter.
# perl is not on PowerShell's PATH here (only in Git Bash), so locate it explicitly.
$perlCmd = Get-Command perl -ErrorAction SilentlyContinue
if ($perlCmd) {
    $perlExe = $perlCmd.Source
} elseif (Test-Path "C:\Program Files\Git\usr\bin\perl.exe") {
    $perlExe = "C:\Program Files\Git\usr\bin\perl.exe"
} else {
    throw "perl.exe not found (checked PATH and C:\Program Files\Git\usr\bin\perl.exe)"
}
$perlScript = Join-Path $scriptDir "md2docx.pl"
$documentXmlPath = Join-Path $buildDir "word\document.xml"
& $perlExe $perlScript $mdFull $documentXmlPath
if ($LASTEXITCODE -ne 0) { throw "md2docx.pl failed with exit code $LASTEXITCODE" }

# Zip the package manually (System.IO.Compression.ZipFile in PS 5.1 forces forward slashes only via manual entry naming)
if (Test-Path $outDocx) { Remove-Item $outDocx -Force }
$src = (Get-Item $buildDir).FullName
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($outDocx, [System.IO.Compression.ZipArchiveMode]::Create)
Get-ChildItem -Path $src -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($src.Length + 1).Replace([System.IO.Path]::DirectorySeparatorChar, [char]47)
    $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
    $entryStream = $entry.Open()
    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
    $entryStream.Write($bytes, 0, $bytes.Length)
    $entryStream.Close()
}
$zip.Dispose()

Remove-Item $buildDir -Recurse -Force
Write-Output "Built: $outDocx"

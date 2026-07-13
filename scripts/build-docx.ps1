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

# If the article contained ```chart``` directives, md2docx.pl left a sidecar .charts.tsv
# next to document.xml (id, type, title, labels, values) and placeholder runs
# "##CHART_PLACEHOLDER:<id>##" inside document.xml. Render each chart as a PNG with
# System.Drawing and splice it into the package (media part + relationship + content type).
$chartsTsv = "$documentXmlPath.charts.tsv"
if (Test-Path $chartsTsv) {
    . (Join-Path $scriptDir "charts.ps1")

    $mediaDir = Join-Path $buildDir "word\media"
    New-Item -ItemType Directory -Force -Path $mediaDir | Out-Null

    $documentXml = Get-Content -LiteralPath $documentXmlPath -Raw -Encoding UTF8
    $relsPath = Join-Path $buildDir "word\_rels\document.xml.rels"
    $relsXml = Get-Content -LiteralPath $relsPath -Raw -Encoding UTF8
    $ctPath = Join-Path $buildDir "[Content_Types].xml"
    $ctXml = Get-Content -LiteralPath $ctPath -Raw -Encoding UTF8

    if ($ctXml -notmatch 'Extension="png"') {
        $ctXml = $ctXml.Replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
    }

    $nextRelId = 2 # rId1 is already used by styles.xml
    $docPrId = 1

    Get-Content $chartsTsv -Encoding UTF8 | Where-Object { $_.Trim() -ne '' } | ForEach-Object {
        $fields = $_ -split "`t"
        $chartId = $fields[0]
        $chartType = $fields[1]
        $chartTitle = $fields[2]
        $labels = @($fields[3] -split ',' | ForEach-Object { $_.Trim() })
        $values = @($fields[4] -split ',' | ForEach-Object { [double]($_.Trim()) })

        $pngName = "$chartId.png"
        $pngPath = Join-Path $mediaDir $pngName

        if ($chartType -eq 'pie') {
            New-PieChartPng -Path $pngPath -Title $chartTitle -Labels $labels -Values $values
            $pxW = 780; $pxH = 520
        } else {
            New-BarChartPng -Path $pngPath -Title $chartTitle -Labels $labels -Values $values
            $pxW = 900; $pxH = 580
        }

        $emu = Get-PngEmuSize -ChartType $chartType -PixelWidth $pxW -PixelHeight $pxH
        $relId = "rId$nextRelId"
        $nextRelId++

        $relsXml = $relsXml.Replace('</Relationships>', "<Relationship Id=`"$relId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$pngName`"/></Relationships>")

        $drawing = "<w:drawing><wp:inline xmlns:wp=`"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`" distT=`"0`" distB=`"0`" distL=`"0`" distR=`"0`">" +
            "<wp:extent cx=`"$($emu.cx)`" cy=`"$($emu.cy)`"/><wp:effectExtent l=`"0`" t=`"0`" r=`"0`" b=`"0`"/>" +
            "<wp:docPr id=`"$docPrId`" name=`"$chartId`"/>" +
            "<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" noChangeAspect=`"1`"/></wp:cNvGraphicFramePr>" +
            "<a:graphic xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><a:graphicData uri=`"http://schemas.openxmlformats.org/drawingml/2006/picture`">" +
            "<pic:pic xmlns:pic=`"http://schemas.openxmlformats.org/drawingml/2006/picture`">" +
            "<pic:nvPicPr><pic:cNvPr id=`"$docPrId`" name=`"$pngName`"/><pic:cNvPicPr/></pic:nvPicPr>" +
            "<pic:blipFill><a:blip r:embed=`"$relId`"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
            "<pic:spPr><a:xfrm><a:off x=`"0`" y=`"0`"/><a:ext cx=`"$($emu.cx)`" cy=`"$($emu.cy)`"/></a:xfrm><a:prstGeom prst=`"rect`"><a:avLst/></a:prstGeom></pic:spPr>" +
            "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"

        $placeholderRun = "<w:r><w:t xml:space=`"preserve`">##CHART_PLACEHOLDER:$chartId##</w:t></w:r>"
        $imageRun = "<w:r>$drawing</w:r>"
        $documentXml = $documentXml.Replace($placeholderRun, $imageRun)

        $docPrId++
    }

    [System.IO.File]::WriteAllText($documentXmlPath, $documentXml, [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText($relsPath, $relsXml, [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText($ctPath, $ctXml, [System.Text.UTF8Encoding]::new($false))
    Remove-Item $chartsTsv -Force
}

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

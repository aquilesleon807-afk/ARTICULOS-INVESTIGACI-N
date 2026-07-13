<#
Renders simple bar/pie chart PNGs with a soft pastel palette using System.Drawing (.NET Framework,
always available on Windows PowerShell 5.1 -- no node/python/matplotlib needed).
Dot-source this file from build-docx.ps1: . (Join-Path $scriptDir "charts.ps1")
#>

Add-Type -AssemblyName System.Drawing

$script:ChartPalette = @('#AED6F1','#A9DFBF','#F9E79F','#F5B7B1','#D7BDE2','#F8C471','#A3E4D7','#F0B27A')

function Get-ChartColor {
    param([int]$Index)
    $hex = $script:ChartPalette[$Index % $script:ChartPalette.Count]
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

function New-BarChartPng {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string[]]$Labels,
        [Parameter(Mandatory=$true)][double[]]$Values
    )
    $width = 900; $height = 580
    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::White)

    $titleFont = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $labelFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $valueFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $darkBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60,60,60))

    $titleSize = $g.MeasureString($Title, $titleFont)
    $g.DrawString($Title, $titleFont, $darkBrush, [float](($width - $titleSize.Width) / 2), 15)

    $chartTop = 70
    $chartBottom = $height - 150
    $chartLeft = 60
    $chartRight = $width - 40
    $chartHeight = $chartBottom - $chartTop
    $maxVal = ($Values | Measure-Object -Maximum).Maximum
    if ($maxVal -le 0) { $maxVal = 1 }

    $axisPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180,180,180)), 1
    $g.DrawLine($axisPen, $chartLeft, $chartBottom, $chartRight, $chartBottom)

    $n = $Labels.Count
    if ($n -eq 0) { $n = 1 }
    $slot = ($chartRight - $chartLeft) / $n
    $barWidth = [Math]::Min(70, $slot * 0.6)

    for ($i = 0; $i -lt $Labels.Count; $i++) {
        $val = $Values[$i]
        $barH = [double]$val / $maxVal * $chartHeight
        $x = $chartLeft + $i * $slot + ($slot - $barWidth) / 2
        $y = $chartBottom - $barH
        $brush = New-Object System.Drawing.SolidBrush(Get-ChartColor $i)
        $g.FillRectangle($brush, [float]$x, [float]$y, [float]$barWidth, [float]$barH)
        $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140,140,140)), 1
        $g.DrawRectangle($borderPen, [float]$x, [float]$y, [float]$barWidth, [float]$barH)

        $valStr = "$val"
        $valSize = $g.MeasureString($valStr, $valueFont)
        $g.DrawString($valStr, $valueFont, $darkBrush, [float]($x + $barWidth / 2 - $valSize.Width / 2), [float]($y - $valSize.Height - 2))

        $lbl = $Labels[$i]
        $lblSize = $g.MeasureString($lbl, $labelFont)
        if ($lblSize.Width -gt ($slot - 4)) {
            $state = $g.Save()
            $g.TranslateTransform([float]($x + $barWidth / 2 + 6), [float]($chartBottom + 10))
            $g.RotateTransform(35)
            $g.DrawString($lbl, $labelFont, $darkBrush, 0, 0)
            $g.Restore($state)
        } else {
            $g.DrawString($lbl, $labelFont, $darkBrush, [float]($x + $barWidth / 2 - $lblSize.Width / 2), [float]($chartBottom + 8))
        }
    }

    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
}

function New-PieChartPng {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string[]]$Labels,
        [Parameter(Mandatory=$true)][double[]]$Values
    )
    $width = 780; $height = 520
    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::White)

    $titleFont = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $legendFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $darkBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60,60,60))

    $titleSize = $g.MeasureString($Title, $titleFont)
    $g.DrawString($Title, $titleFont, $darkBrush, [float](($width - $titleSize.Width) / 2), 15)

    $total = ($Values | Measure-Object -Sum).Sum
    if ($total -le 0) { $total = 1 }
    $cx = 250; $cy = 290; $r = 165
    $startAngle = -90.0
    for ($i = 0; $i -lt $Labels.Count; $i++) {
        $sweep = 360.0 * ($Values[$i] / $total)
        if ($sweep -le 0) { continue }
        $brush = New-Object System.Drawing.SolidBrush(Get-ChartColor $i)
        $g.FillPie($brush, [float]($cx - $r), [float]($cy - $r), [float]($r * 2), [float]($r * 2), [float]$startAngle, [float]$sweep)
        $whitePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White), 2
        $g.DrawPie($whitePen, [float]($cx - $r), [float]($cy - $r), [float]($r * 2), [float]($r * 2), [float]$startAngle, [float]$sweep)
        $startAngle += $sweep
    }

    $legendX = 450; $legendY = 130
    for ($i = 0; $i -lt $Labels.Count; $i++) {
        $brush = New-Object System.Drawing.SolidBrush(Get-ChartColor $i)
        $g.FillRectangle($brush, $legendX, ($legendY + $i * 28), 16, 16)
        $pct = [Math]::Round(100.0 * $Values[$i] / $total, 1)
        $g.DrawString("$($Labels[$i]) ($pct%)", $legendFont, $darkBrush, ($legendX + 22), ($legendY + $i * 28 - 2))
    }

    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
}

function Get-PngEmuSize {
    param(
        [Parameter(Mandatory=$true)][string]$ChartType,
        [Parameter(Mandatory=$true)][int]$PixelWidth,
        [Parameter(Mandatory=$true)][int]$PixelHeight
    )
    $targetWidthEmu = if ($ChartType -eq 'pie') { 5029200 } else { 5486400 } # 5.5in / 6.0in
    $ratio = $PixelHeight / $PixelWidth
    $cx = [int64]$targetWidthEmu
    $cy = [int64]([double]$targetWidthEmu * $ratio)
    return @{ cx = $cx; cy = $cy }
}

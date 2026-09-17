# Regenerates the app icons (System.Drawing only; run with PowerShell 7 on Windows):
#
#   pwsh tools/make-icons.ps1                      # rewrites src/SendToOneNote/Assets/app-release.ico + app-debug.ico
#   pwsh tools/make-icons.ps1 -PreviewDir <dir>    # also writes per-size PNGs and a contact sheet for review
#
# Release = purple rounded square, white "N", envelope badge from 32 px up.
# Debug   = same shape in orange, dark "D" badge from 24 px up, so the two builds are
#           told apart at a glance in the tray (colour) and everywhere else (badge).
# Frames under 256 px are stored as 32-bit BMP and the 256 px frame as PNG, the layout
# every Windows icon consumer accepts.
param(
    [string]$OutDir = (Join-Path $PSScriptRoot '..\src\SendToOneNote\Assets'),
    [string]$PreviewDir
)

Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$Sizes = 16, 20, 24, 32, 48, 64, 256
$NarrowN = 0.78   # Segoe UI Bold's N is nearly square and reads as a sideways Z; squeeze it.

function New-RoundedRect([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $d = $r * 2
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

# Fills a glyph centred on (cx, cy), sized from its real outline rather than font metrics.
function Add-Glyph($g, [string]$text, [float]$cx, [float]$cy, [float]$height, $brush, [float]$xScale = 1.0) {
    $family = [System.Drawing.FontFamily]::new('Segoe UI')
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $path.AddString($text, $family, [int][System.Drawing.FontStyle]::Bold, 100,
        [System.Drawing.PointF]::new(0, 0), [System.Drawing.StringFormat]::GenericTypographic)
    $b = $path.GetBounds()
    $scale = $height / $b.Height
    $m = [System.Drawing.Drawing2D.Matrix]::new()
    $m.Translate($cx - ($b.X + $b.Width / 2) * $scale * $xScale, $cy - ($b.Y + $b.Height / 2) * $scale)
    $m.Scale($scale * $xScale, $scale)
    $path.Transform($m)
    $g.FillPath($brush, $path)
    $path.Dispose(); $family.Dispose(); $m.Dispose()
}

function New-IconBitmap([int]$S, [bool]$IsDebug) {
    $bmp = [System.Drawing.Bitmap]::new($S, $S, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([System.Drawing.Color]::Transparent)

    if ($IsDebug) { $top = '#F59A3A'; $bottom = '#D4610A' }
    else          { $top = '#8E35C9'; $bottom = '#661A96' }
    $cTop = [System.Drawing.ColorTranslator]::FromHtml($top)
    $cBottom = [System.Drawing.ColorTranslator]::FromHtml($bottom)

    $margin = [Math]::Max(0.5, $S * 0.03)
    $body = New-RoundedRect $margin $margin ($S - 2 * $margin) ($S - 2 * $margin) ($S * 0.2)
    $grad = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
        [System.Drawing.PointF]::new(0, 0), [System.Drawing.PointF]::new(0, $S), $cTop, $cBottom)
    $g.FillPath($grad, $body)

    $white = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    # A badge smaller than this turns to mush; below it the colour alone carries the meaning.
    $hasBadge = if ($IsDebug) { $S -ge 24 } else { $S -ge 32 }

    if ($hasBadge) { Add-Glyph $g 'N' ($S * 0.40) ($S * 0.44) ($S * 0.60) $white $NarrowN }
    else           { Add-Glyph $g 'N' ($S * 0.50) ($S * 0.50) ($S * 0.66) $white $NarrowN }

    if ($hasBadge -and -not $IsDebug) {
        # Envelope, bottom-right, separated from the N by a ring of the body colour.
        $ex = $S * 0.50; $ey = $S * 0.58; $ew = $S * 0.41; $eh = $S * 0.29
        $env = New-RoundedRect $ex $ey $ew $eh ($S * 0.035)
        $cutPen = [System.Drawing.Pen]::new($cBottom, [float]($S * 0.07))
        $g.DrawPath($cutPen, $env)
        $g.FillPath($white, $env)
        $flap = [System.Drawing.Pen]::new($cBottom, [float][Math]::Max(1.0, $S * 0.035))
        $flap.LineJoin = 'Round'
        $pts = [System.Drawing.PointF[]]@(
            [System.Drawing.PointF]::new($ex + $S * 0.02, $ey + $S * 0.03),
            [System.Drawing.PointF]::new($ex + $ew / 2, $ey + $eh * 0.58),
            [System.Drawing.PointF]::new($ex + $ew - $S * 0.02, $ey + $S * 0.03))
        $g.DrawLines($flap, $pts)
        $cutPen.Dispose(); $flap.Dispose(); $env.Dispose()
    }
    if ($hasBadge -and $IsDebug) {
        # Dark "D" badge, bottom-right, kept inside the body.
        $d = $S * 0.47; $bx = $S * 0.455; $by = $S * 0.455
        $ring = [System.Drawing.SolidBrush]::new($cBottom)
        $g.FillEllipse($ring, [float]($bx - $S * 0.035), [float]($by - $S * 0.035), [float]($d + $S * 0.07), [float]($d + $S * 0.07))
        $dark = [System.Drawing.SolidBrush]::new([System.Drawing.ColorTranslator]::FromHtml('#1F1F1F'))
        $g.FillEllipse($dark, [float]$bx, [float]$by, [float]$d, [float]$d)
        Add-Glyph $g 'D' ($bx + $d / 2 + $S * 0.01) ($by + $d / 2) ($d * 0.56) $white
        $ring.Dispose(); $dark.Dispose()
    }

    $white.Dispose(); $grad.Dispose(); $body.Dispose(); $g.Dispose()
    return $bmp
}

# One ICO frame: PNG for 256 px, otherwise a 32-bit DIB (header with doubled height, bottom-up BGRA
# rows, then an all-zero AND mask because the alpha channel carries the transparency).
function Get-FrameBytes($bmp) {
    $S = $bmp.Width
    $ms = [System.IO.MemoryStream]::new()
    if ($S -ge 256) {
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        return , $ms.ToArray()
    }
    $rect = [System.Drawing.Rectangle]::new(0, 0, $S, $S)
    $data = $bmp.LockBits($rect, 'ReadOnly', [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $pixels = [byte[]]::new($S * $S * 4)
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $pixels, 0, $pixels.Length)
    $bmp.UnlockBits($data)

    $w = [System.IO.BinaryWriter]::new($ms)
    $w.Write([uint32]40); $w.Write([int32]$S); $w.Write([int32]($S * 2))
    $w.Write([uint16]1); $w.Write([uint16]32); $w.Write([uint32]0)
    $w.Write([uint32]($S * $S * 4)); $w.Write([int32]0); $w.Write([int32]0); $w.Write([uint32]0); $w.Write([uint32]0)
    for ($row = $S - 1; $row -ge 0; $row--) { $w.Write($pixels, $row * $S * 4, $S * 4) }
    $maskRow = [int]([Math]::Floor(($S + 31) / 32)) * 4
    $w.Write([byte[]]::new($maskRow * $S))
    $w.Flush()
    return , $ms.ToArray()
}

function Write-Ico([string]$path, $bitmaps) {
    $frames = foreach ($b in $bitmaps) { , (Get-FrameBytes $b) }
    $fs = [System.IO.File]::Create($path)
    $w = [System.IO.BinaryWriter]::new($fs)
    $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$bitmaps.Count)
    $offset = 6 + 16 * $bitmaps.Count
    for ($i = 0; $i -lt $bitmaps.Count; $i++) {
        $s = $bitmaps[$i].Width
        $dim = if ($s -ge 256) { 0 } else { $s }   # 0 means 256 in the directory entry
        $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
        $w.Write([uint16]1); $w.Write([uint16]32)
        $w.Write([uint32]$frames[$i].Length); $w.Write([uint32]$offset)
        $offset += $frames[$i].Length
    }
    foreach ($f in $frames) { $w.Write($f) }
    $w.Dispose(); $fs.Dispose()
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$sets = @{}
foreach ($variant in 'release', 'debug') {
    $sets[$variant] = foreach ($s in $Sizes) { , (New-IconBitmap $s ($variant -eq 'debug')) }
    Write-Ico (Join-Path $OutDir "app-$variant.ico") $sets[$variant]
}
Get-ChildItem $OutDir -Filter 'app-*.ico' | Select-Object Name, Length

if ($PreviewDir) {
    # Contact sheet: each variant on a light and a dark strip — actual sizes, 16/24/32 zoomed 6x, and the 256 frame.
    New-Item -ItemType Directory -Force $PreviewDir | Out-Null
    foreach ($variant in $sets.Keys) {
        foreach ($f in $sets[$variant]) {
            $f.Save((Join-Path $PreviewDir "$variant-$($f.Width).png"), [System.Drawing.Imaging.ImageFormat]::Png)
        }
    }
    $W = 1180; $stripH = 230
    $sheet = [System.Drawing.Bitmap]::new($W, $stripH * 4)
    $g = [System.Drawing.Graphics]::FromImage($sheet)
    $font = [System.Drawing.Font]::new('Segoe UI', 10)
    $row = 0
    foreach ($variant in 'release', 'debug') {
        foreach ($bg in '#F3F3F3', '#202020') {
            $y0 = $row * $stripH
            $g.FillRectangle([System.Drawing.SolidBrush]::new([System.Drawing.ColorTranslator]::FromHtml($bg)), 0, $y0, $W, $stripH)
            $ink = if ($bg -eq '#202020') { [System.Drawing.Brushes]::Gainsboro } else { [System.Drawing.Brushes]::DimGray }
            $g.DrawString("$variant  -  actual size", $font, $ink, 12, $y0 + 8)
            $x = 16
            $g.InterpolationMode = 'NearestNeighbor'; $g.PixelOffsetMode = 'Half'
            foreach ($f in $sets[$variant] | Where-Object { $_.Width -le 64 }) {
                $g.DrawImage($f, $x, $y0 + 40, $f.Width, $f.Height)
                $g.DrawString("$($f.Width)", $font, $ink, $x, $y0 + 40 + $f.Height + 4)
                $x += $f.Width + 22
            }
            $x += 20
            $g.DrawString('zoomed 6x: 16, 24, 32', $font, $ink, $x, $y0 + 8)
            foreach ($s in 16, 24, 32) {
                $f = $sets[$variant] | Where-Object { $_.Width -eq $s }
                $g.DrawImage($f, $x, $y0 + 32, $s * 6, $s * 6)
                $x += $s * 6 + 18
            }
            $g.InterpolationMode = 'HighQualityBicubic'
            $big = $sets[$variant] | Where-Object { $_.Width -eq 256 }
            $g.DrawString('256 (shown at 192)', $font, $ink, $W - 210, $y0 + 8)
            $g.DrawImage($big, $W - 210, $y0 + 30, 192, 192)
            $row++
        }
    }
    $sheet.Save((Join-Path $PreviewDir 'contact-sheet.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $sheet.Dispose()
    "Preview written to $PreviewDir"
}

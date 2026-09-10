#Requires -Version 5
<#
  Labelled, high-zoom contact sheet of the Genesis tiles for identifying
  sprite orientations. Reads reference\genesis-tiles\rNN_cNN.png.

  Params: -Rows "12,13"  (which atlas rows)   -Scale 8   -PerRow 12
  Output: reference\tilesheet.png
#>
param(
    [string]$Rows = '',                 # empty = all rows
    [int]$Scale = 8,
    [int]$PerRow = 12,
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$dir = Join-Path $Root 'reference\genesis-tiles'
if (-not (Test-Path (Join-Path $dir 'r00_c00.png'))) { & (Join-Path $Root 'reference\slice-tiles.ps1') }

$files = @(Get-ChildItem $dir -Filter '*.png' | Sort-Object Name)
if ($Rows) {
    $want = @($Rows.Split(',') | ForEach-Object { 'r{0:00}_' -f [int]$_ })
    $files = @($files | Where-Object { $fn = $_.Name; ($want | Where-Object { $fn.StartsWith($_) }).Count -gt 0 })
}

$tile  = [int](16 * $Scale)
$lbl   = 16
$cellW = [int]($tile + 8)
$cellH = [int]($tile + $lbl + 8)
$nTiles = $files.Count
$cols   = [int][math]::Min($PerRow, $nTiles)
$rowN   = [int][math]::Ceiling($nTiles / [double]$PerRow)

$img = New-Object System.Drawing.Bitmap ([int]($cols * $cellW + 8)), ([int]($rowN * $cellH + 8))
$g = [System.Drawing.Graphics]::FromImage($img)
$g.Clear([System.Drawing.Color]::FromArgb(255, 30, 30, 36))
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
$g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
$font = New-Object System.Drawing.Font('Consolas', 9)
$checker1 = [System.Drawing.Color]::FromArgb(255, 60, 60, 68)
$checker2 = [System.Drawing.Color]::FromArgb(255, 45, 45, 52)

for ($i = 0; $i -lt $files.Count; $i++) {
    $cx = 8 + ($i % $PerRow) * $cellW
    $cy = 8 + [math]::Floor($i / $PerRow) * $cellH
    # checker backing so transparent pixels are visible
    for ($by = 0; $by -lt $tile; $by += 8*$Scale) {
        for ($bx = 0; $bx -lt $tile; $bx += 8*$Scale) {
            $col = if ((($bx/(8*$Scale)) + ($by/(8*$Scale))) % 2) { $checker1 } else { $checker2 }
            $g.FillRectangle((New-Object System.Drawing.SolidBrush $col), $cx+$bx, $cy+$by, 8*$Scale, 8*$Scale)
        }
    }
    $t = [System.Drawing.Bitmap]::FromFile($files[$i].FullName)
    $g.DrawImage($t, $cx, $cy, $tile, $tile)
    $t.Dispose()
    $name = $files[$i].BaseName -replace '_c', ' c' -replace '^r', 'r'
    $g.DrawString($name, $font, [System.Drawing.Brushes]::White, $cx, $cy + $tile + 2)
}
$out = Join-Path $Root 'reference\tilesheet.png'
$img.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $img.Dispose()
Write-Host "wrote $out ($($files.Count) tiles, $($img.Width)x$($img.Height))"

#Requires -Version 5
<#
  Slice genesis-gauntlet-samples.png into 16x16 tiles.

  The sheet is a 44 x 14 grid of 16px tiles, pitch 18 (16px tile + 2px grey
  separator), origin (3,5). Magenta (#FF00FF) is the transparency key and is
  written out as alpha 0.

  Output: reference/genesis-tiles/r{RR}_c{CC}.png  (non-empty cells only)
  Usage:  powershell -ExecutionPolicy Bypass -File reference\slice-tiles.ps1
#>
param(
    [string]$Src = (Join-Path $PSScriptRoot 'genesis-gauntlet-samples.png'),
    [string]$Out = (Join-Path $PSScriptRoot 'genesis-tiles')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$OX = 3; $OY = 5; $PITCH = 18; $SZ = 16

$bmp = [System.Drawing.Bitmap]::FromFile($Src)
$W = $bmp.Width; $H = $bmp.Height
$rect = New-Object System.Drawing.Rectangle 0, 0, $W, $H
$d = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $d.Stride
$px = New-Object byte[] ($stride * $H)
[System.Runtime.InteropServices.Marshal]::Copy($d.Scan0, $px, 0, $px.Length)
$bmp.UnlockBits($d); $bmp.Dispose()

if (Test-Path $Out) { Remove-Item -Recurse -Force $Out }
New-Item -ItemType Directory -Path $Out | Out-Null

$cols = 1 + [int](($W - $OX - $SZ) / $PITCH)
$rows = 1 + [int](($H - $OY - $SZ) / $PITCH)
$saved = 0
for ($rr = 0; $rr -lt $rows; $rr++) {
    for ($cc = 0; $cc -lt $cols; $cc++) {
        $sx = $OX + $cc * $PITCH; $sy = $OY + $rr * $PITCH
        $tile = New-Object System.Drawing.Bitmap $SZ, $SZ, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $any = $false
        for ($y = 0; $y -lt $SZ; $y++) {
            for ($x = 0; $x -lt $SZ; $x++) {
                $i = ($sy + $y) * $stride + ($sx + $x) * 4
                $b = $px[$i]; $g = $px[$i + 1]; $r = $px[$i + 2]
                if ($r -gt 200 -and $g -lt 80 -and $b -gt 200) {
                    $tile.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, 0, 0, 0))
                } else {
                    $tile.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(255, $r, $g, $b)); $any = $true
                }
            }
        }
        if ($any) {
            $tile.Save((Join-Path $Out ('r{0:00}_c{1:00}.png' -f $rr, $cc)), [System.Drawing.Imaging.ImageFormat]::Png)
            $saved++
        }
        $tile.Dispose()
    }
}
Write-Host "grid ${cols}x${rows}, saved $saved tiles to $Out"

#Requires -Version 5
# Opens the workbook, runs frames headless, checks the render buffer + camera.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$VIEW_COLS = 24            # size we pin to for the deterministic checks
$VIEW_ROWS = 15
$MAP = 32
$MIN_COLS = 12
$MIN_ROWS = 8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$wb = $null
$fail = 0
function Check($label, $cond, $got) {
    if ($cond) { Write-Host "  ok   $label" }
    else       { Write-Host "  FAIL $label  (got: $got)"; $script:fail++ }
}
try {
    $wb = $excel.Workbooks.Open($book)
    $ws = $wb.Worksheets.Item('Screen')

    function View-Row($r) {
        $vals = $ws.Range($ws.Cells.Item($r,1), $ws.Cells.Item($r,$VIEW_COLS)).Value2
        -join (1..$VIEW_COLS | ForEach-Object { [string]$vals.GetValue(1, $_) })
    }
    function Find-Player {
        for ($r = 1; $r -le $VIEW_ROWS; $r++) {
            for ($c = 1; $c -le $VIEW_COLS; $c++) {
                if ([string]$ws.Cells.Item($r,$c).Value2 -eq '@') { return "$r,$c" }
            }
        }
        return $null
    }
    # DebugState() -> "viewCols;viewRows;camR;camC;plR;plC;state;cellPts"
    function State {
        $p = ([string]$excel.Run('DebugState')).Split(';')
        [pscustomobject]@{
            vc = [int]$p[0]; vr = [int]$p[1]; camR = [int]$p[2]; camC = [int]$p[3]
            plR = [int]$p[4]; plC = [int]$p[5]; st = $p[6]; cell = [double]$p[7]
        }
    }

    # ---- Part 0: PLAY button anchored at A1 ----
    $btn = $null
    try { $btn = $ws.Shapes.Item('btnPlay') } catch {}
    Check "btnPlay exists near A1" ($btn -and $btn.Top -lt 20 -and $btn.Left -lt 20) `
        $(if ($btn) { "top=$([int]$btn.Top) left=$([int]$btn.Left)" } else { 'missing' })

    # ---- Part 1: FitViewport produces a sane, capped size ----
    $excel.Run('GameInit')
    $excel.Run('FitViewport')
    $excel.Run('RenderInit')
    $excel.Run('CenterCamera')
    $excel.Run('RenderFrame', [double]0)
    $s = State
    Check "fit cols in [$MIN_COLS..$MAP]" ($s.vc -ge $MIN_COLS -and $s.vc -le $MAP) $s.vc
    Check "fit rows in [$MIN_ROWS..$MAP]" ($s.vr -ge $MIN_ROWS -and $s.vr -le $MAP) $s.vr
    Check "fit cell size in [12..48]pt" ($s.cell -ge 12 -and $s.cell -le 48) $s.cell
    Check "fit fills a dimension" (
        $s.vc -eq $MAP -or $s.vr -eq $MAP -or $s.cell -ge 47.9
    ) "view $($s.vc)x$($s.vr) @ $($s.cell)pt"
    Check "fit camera in range" (
        $s.camR -ge 1 -and $s.camR -le ($MAP - $s.vr + 1) -and
        $s.camC -ge 1 -and $s.camC -le ($MAP - $s.vc + 1)
    ) "$($s.camR),$($s.camC) view $($s.vc)x$($s.vr)"

    # ---- Part 2: pin to 24x15 for deterministic checks ----
    $excel.Run('DebugSetViewport', [int]$VIEW_COLS, [int]$VIEW_ROWS)
    $excel.Run('GameInit')
    $excel.Run('RenderInit')
    $excel.Run('CenterCamera')
    $excel.Run('RenderFrame', [double]0)

    Check "top row is solid wall"       ((View-Row 1) -eq ('#' * $VIEW_COLS)) (View-Row 1)
    Check "player rendered at view 2,2"  ((Find-Player) -eq '2,2')            (Find-Player)

    $w = [math]::Round($ws.Columns(1).Width, 2)
    $h = [math]::Round($ws.Rows(1).Height, 2)
    Check "cells ~square ~18pt" (($w -ge 16 -and $w -le 21) -and ($h -eq 18)) "$w x $h"
    Check "HUD health line" ([string]$ws.Cells.Item($VIEW_ROWS + 2, 1).Value2 -like 'HEALTH*SCORE*') ([string]$ws.Cells.Item($VIEW_ROWS + 2, 1).Value2)

    # scroll: step player +7 rows / +20 cols into open floor -> map (9,22)
    $excel.Run('DebugStep', [int]7, [int]20)
    $excel.Run('RenderFrame', [double]0)
    $s = State
    Check "player at map 9,22"           ($s.plR -eq 9 -and $s.plC -eq 22)     "$($s.plR),$($s.plC)"
    Check "camera clamped to 2,9"        ($s.camR -eq 2 -and $s.camC -eq 9)    "$($s.camR),$($s.camC)"
    Check "player re-rendered at 8,14"   ((Find-Player) -eq '8,14')            (Find-Player)
    Check "top row now interior floor"   ((View-Row 1) -ne ('#' * $VIEW_COLS)) (View-Row 1)

    if ($fail) { throw "$fail check(s) failed" }
    Write-Host "OK"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

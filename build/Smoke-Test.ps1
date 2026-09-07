#Requires -Version 5
# Opens the workbook, runs frames headless, checks the half-cell render + camera.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$VC = 36        # viewport half-cells (fixed: VIEW_BLOCK_COLS*2)
$VR = 20        # VIEW_BLOCK_ROWS*2
$MAP = 64       # half-cells

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
        $vals = $ws.Range($ws.Cells.Item($r,1), $ws.Cells.Item($r,$VC)).Value2
        -join (1..$VC | ForEach-Object { [string]$vals.GetValue(1, $_) })
    }
    function Find-Player {   # first '@', and total count
        $first = $null; $n = 0
        for ($r = 1; $r -le $VR; $r++) {
            for ($c = 1; $c -le $VC; $c++) {
                if ([string]$ws.Cells.Item($r,$c).Value2 -eq '@') {
                    if (-not $first) { $first = "$r,$c" }
                    $n++
                }
            }
        }
        [pscustomobject]@{ first = $first; n = $n }
    }
    # DebugState() -> "vc;vr;camR;camC;plHR;plHC;state;cellPts"
    function State {
        $p = ([string]$excel.Run('DebugState')).Split(';')
        [pscustomobject]@{
            vc=[int]$p[0]; vr=[int]$p[1]; camR=[int]$p[2]; camC=[int]$p[3]
            plHR=[int]$p[4]; plHC=[int]$p[5]; st=$p[6]; cell=[double]$p[7]
        }
    }

    # ---- PLAY button anchored at A1 ----
    $btn = $null
    try { $btn = $ws.Shapes.Item('btnPlay') } catch {}
    Check "btnPlay exists near A1" ($btn -and $btn.Top -lt 20 -and $btn.Left -lt 20) `
        $(if ($btn) { "top=$([int]$btn.Top) left=$([int]$btn.Left)" } else { 'missing' })

    # ---- init: spawn at block (2,2) = half-cell (3,3), camera clamped to (1,1) ----
    $excel.Run('GameInit')
    $excel.Run('FitViewport')
    $excel.Run('RenderInit')
    $excel.Run('CenterCamera')
    $excel.Run('RenderFrame', [double]0)
    $s = State

    Check "viewport is 36x20 half-cells" ($s.vc -eq $VC -and $s.vr -eq $VR) "$($s.vc)x$($s.vr)"
    Check "fit cell size in [8..26]pt"   ($s.cell -ge 8 -and $s.cell -le 26) $s.cell
    Check "row height tracks cell size"  ([math]::Abs([double]$ws.Rows(1).Height - $s.cell) -le 1.5) "$([math]::Round([double]$ws.Rows(1).Height,1)) vs $($s.cell)"
    Check "camera clamped to 1,1"        ($s.camR -eq 1 -and $s.camC -eq 1) "$($s.camR),$($s.camC)"

    Check "top-left block is 2x2 wall" (
        (View-Row 1).Substring(0,2) -eq '##' -and (View-Row 2).Substring(0,2) -eq '##'
    ) ((View-Row 1) + ' / ' + (View-Row 2))

    $p = Find-Player
    Check "player is a 2x2 patch" ($p.n -eq 4)          $p.n
    Check "player at view 3,3"    ($p.first -eq '3,3')   $p.first

    # ---- scroll: +14 half-rows / +30 half-cols into open floor -> half-cell (17,33), block (9,17) ----
    $excel.Run('DebugStep', [int]14, [int]30)
    $excel.Run('RenderFrame', [double]0)
    $s = State
    Check "player at half-cell 17,33" ($s.plHR -eq 17 -and $s.plHC -eq 33) "$($s.plHR),$($s.plHC)"
    Check "camera clamped to 7,15"    ($s.camR -eq 7 -and $s.camC -eq 15)  "$($s.camR),$($s.camC)"

    $p = Find-Player
    Check "player re-rendered at 11,19" ($p.first -eq '11,19' -and $p.n -eq 4) "$($p.first) n=$($p.n)"
    Check "top row now has floor"       ((View-Row 1) -match '\.')            (View-Row 1)

    if ($fail) { throw "$fail check(s) failed" }
    Write-Host "OK"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

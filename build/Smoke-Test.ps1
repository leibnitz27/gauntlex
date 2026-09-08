#Requires -Version 5
# Opens the workbook, drives it headless, checks render + M2 game logic.
$ErrorActionPreference = 'Stop'
$book = Join-Path (Split-Path -Parent $PSScriptRoot) 'gauntlex.xlsm'

$VC = 36; $VR = 20; $MAP = 64
$START_HEALTH = 2000; $START_LIVES = 3; $FOOD = 350

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
$wb = $null; $fail = 0
function Check($label, $cond, $got) {
    if ($cond) { Write-Host "  ok   $label" }
    else       { Write-Host "  FAIL $label  (got: $got)"; $script:fail++ }
}
try {
    $wb = $excel.Workbooks.Open($book)
    $ws = $wb.Worksheets.Item('Screen')

    function View-Row($r) {
        $v = $ws.Range($ws.Cells.Item($r,1), $ws.Cells.Item($r,$VC)).Value2
        -join (1..$VC | ForEach-Object { [string]$v.GetValue(1, $_) })
    }
    function Find($glyph) {
        $first = $null; $n = 0
        for ($r = 1; $r -le $VR; $r++) { for ($c = 1; $c -le $VC; $c++) {
            if ([string]$ws.Cells.Item($r,$c).Value2 -eq $glyph) { if (-not $first) { $first = "$r,$c" }; $n++ }
        } }
        [pscustomobject]@{ first = $first; n = $n }
    }
    # DebugState -> vc;vr;camR;camC;plHR;plHC;state;cell;lives;health;keys;score;grAlive;g1r;g1c
    function State {
        $p = ([string]$excel.Run('DebugState')).Split(';')
        [pscustomobject]@{
            vc=[int]$p[0]; vr=[int]$p[1]; camR=[int]$p[2]; camC=[int]$p[3]
            plHR=[int]$p[4]; plHC=[int]$p[5]; st=$p[6]; cell=[double]$p[7]
            lives=[int]$p[8]; health=[int]$p[9]; keys=[int]$p[10]; score=[int]$p[11]
            grAlive=[int]$p[12]; g1r=[int]$p[13]; g1c=[int]$p[14]
        }
    }
    function Dist($s) { [math]::Abs($s.g1r - $s.plHR) + [math]::Abs($s.g1c - $s.plHC) }

    # ---- render / camera (unchanged from M1) ----
    $excel.Run('GameInit'); $excel.Run('FitViewport'); $excel.Run('RenderInit')
    $excel.Run('CenterCamera'); $excel.Run('RenderFrame', [double]0)
    $s = State
    Check "viewport 36x20"            ($s.vc -eq $VC -and $s.vr -eq $VR) "$($s.vc)x$($s.vr)"
    $uh=[double]$excel.ActiveWindow.UsableHeight; $uw=[double]$excel.ActiveWindow.UsableWidth
    $exp=[math]::Max(8,[math]::Min(60,[math]::Min($uh/$VR,($uw-220)/$VC)))
    Check "cell fits window"          ([math]::Abs($s.cell-$exp) -le 1.0) "$($s.cell) vs ~$([math]::Round($exp,1))"
    Check "camera clamped 1,1"        ($s.camR -eq 1 -and $s.camC -eq 1) "$($s.camR),$($s.camC)"
    Check "top-left 2x2 wall"         ((View-Row 1).Substring(0,2) -eq '##' -and (View-Row 2).Substring(0,2) -eq '##') (View-Row 1)
    $p = Find '@'
    Check "player 2x2 at view 3,3"    ($p.n -eq 4 -and $p.first -eq '3,3') "$($p.first) n=$($p.n)"
    Check "HUD shows LIVES"           ([string]$ws.Cells.Item(6, $VC+2).Value2 -like 'LIVES*') ([string]$ws.Cells.Item(6, $VC+2).Value2)

    $excel.Run('DebugStep', [int]14, [int]30)      # -> half-cell (17,33)
    $excel.Run('RenderFrame', [double]0)
    $s = State
    Check "scroll: player 17,33 cam 7,15" ($s.plHR -eq 17 -and $s.plHC -eq 33 -and $s.camR -eq 7 -and $s.camC -eq 15) "$($s.plHR),$($s.plHC) / $($s.camR),$($s.camC)"
    Check "player re-rendered view 11,19"  ((Find '@').first -eq '11,19') (Find '@').first

    # ---- M2: grunts spawn ----
    $excel.Run('GameInit')
    $s = State
    Check "7 grunts spawned"          ($s.grAlive -eq 7) $s.grAlive
    Check "start health/lives"        ($s.health -eq $START_HEALTH -and $s.lives -eq $START_LIVES) "$($s.health)/$($s.lives)"
    $excel.Run('RenderFrame', [double]0)
    Check "a grunt renders (g glyph)" ((Find 'g').n -ge 1) (Find 'g').n

    # ---- M2: health drains over time (no grunt contact at spawn) ----
    $excel.Run('GameInit')
    $excel.Run('DebugUpdate', [int]2200)            # 2 x DRAIN_MS(1100)
    $s = State
    Check "health drained ~2"        ($s.health -le $START_HEALTH-2 -and $s.health -ge $START_HEALTH-4) $s.health

    # ---- M2: grunt pursues the player ----
    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]19, [int]30)
    $d0 = Dist (State)
    $excel.Run('DebugUpdate', [int]900)
    $d1 = Dist (State)
    Check "grunt closed distance"     ($d1 -lt $d0) "before $d0 after $d1"

    # ---- M2: melee kill on walking into a grunt ----
    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]11, [int]27)       # grunt spawn block (6,15) -> half-cell (11,29)
    $b = State
    $excel.Run('DebugStep', [int]0, [int]1)          # step onto (11,28), overlaps grunt
    $a = State
    Check "melee killed a grunt"      ($a.grAlive -eq $b.grAlive-1) "$($b.grAlive) -> $($a.grAlive)"
    Check "melee scored +10"          ($a.score -eq 10) $a.score

    # ---- M2: key pickup ----
    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]9, [int]43)
    $excel.Run('DebugStep', [int]0, [int]2)          # onto K at block (5,23) = half-cell (9,45)
    Check "key picked up"            ((State).keys -eq 1) (State).keys

    # ---- M2: door opens with a key, blocks without ----
    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]51, [int]57)        # above door D at block (27,29) = half-cell (53,57)
    $excel.Run('DebugStep', [int]2, [int]0)          # no key -> blocked
    Check "door blocks without key"  ((State).plHR -eq 51) (State).plHR
    $excel.Run('DebugWarp', [int]9, [int]43); $excel.Run('DebugStep', [int]0, [int]2)  # get the key
    $excel.Run('DebugWarp', [int]51, [int]57)
    $excel.Run('DebugStep', [int]2, [int]0)          # with key -> opens, moves through
    $s = State
    Check "door opens with key"      ($s.plHR -eq 53 -and $s.keys -eq 0) "plHR $($s.plHR) keys $($s.keys)"

    # ---- M2: reaching the exit wins ----
    $excel.Run('GameInit')
    $excel.Run('DebugWarp', [int]57, [int]57)        # exit block (29,29) = half-cell (57,57)
    $excel.Run('DebugUpdate', [int]50)
    $s = State
    Check "exit -> WON +100"          ($s.st -eq 'WON' -and $s.score -eq 100) "$($s.st) score $($s.score)"

    # ---- M2: death costs a life and respawns; out of lives -> OVER ----
    $excel.Run('GameInit')
    $excel.Run('DebugSetHealth', [int]2)
    $excel.Run('DebugUpdate', [int]3400)             # drains past 0
    $s = State
    Check "death: life lost + respawn" ($s.lives -eq 2 -and $s.health -eq $START_HEALTH -and $s.st -eq 'PLAY' -and $s.plHR -eq 3 -and $s.plHC -eq 3) "lives $($s.lives) hp $($s.health) $($s.st) @$($s.plHR),$($s.plHC)"
    $excel.Run('DebugSetHealth', [int]1); $excel.Run('DebugUpdate', [int]1200)
    $excel.Run('DebugSetHealth', [int]1); $excel.Run('DebugUpdate', [int]1200)
    $s = State
    Check "game over at 0 lives"      ($s.lives -eq 0 -and $s.st -eq 'OVER') "lives $($s.lives) $($s.st)"

    if ($fail) { throw "$fail check(s) failed" }
    Write-Host "OK"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
